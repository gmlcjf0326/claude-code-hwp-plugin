"""pdf_clone.ocr — PaddleOCR 파이프라인 (scanned/hybrid PDF).

함수:
- _detect_pdf_type       : 3 페이지 샘플로 PDF 분류 (native / scanned / hybrid)
- _get_ocr_singleton     : PaddleOCR lazy 로드 (singleton 캐시)
- _preprocess_for_ocr    : opencv 전처리 (gray → denoise → deskew → adaptive threshold)
- _extract_ocr_blocks    : scanned 페이지에서 TextBlock 추출

v0.7.4.9: PaddleOCR 3.x API 호환 (use_textline_orientation, cls 생략)
"""
import sys
from typing import List, Tuple, Any, Dict

from ._models import TextBlock


def _detect_pdf_type(doc, opts: Dict[str, Any]) -> str:
    """Sample 3 pages and classify PDF as native / scanned / hybrid."""
    threshold = int(opts.get("min_native_chars_per_page", 30))
    page_count = doc.page_count
    if page_count == 0:
        return "native"
    sample_idx = sorted(set([0, page_count // 2, page_count - 1]))
    char_counts: List[int] = []
    for i in sample_idx:
        try:
            txt = doc[i].get_text("text") or ""
            char_counts.append(len(txt.strip()))
        except Exception:
            char_counts.append(0)
    if all(c >= threshold for c in char_counts):
        return "native"
    if all(c < 5 for c in char_counts):
        return "scanned"
    return "hybrid"


# v0.7.4.3: PaddleOCR singleton cache — lazy import inside
_OCR_INSTANCE = None
_OCR_LANG = None


def _get_ocr_singleton(lang: str = "korean"):
    """Lazy-load PaddleOCR. First call downloads ~150MB model to ~/.paddleocr.

    v0.7.4.9 (NEW-OCR-API fix): paddleocr 3.x API 호환
    - 2.x: PaddleOCR(use_angle_cls=True, lang=lang, show_log=False)
    - 3.x: PaddleOCR(use_textline_orientation=True, lang=lang)  # show_log 제거
    - 3.x 는 use_textline_orientation 기본값 True 라 생략도 가능
    - `ocr.ocr(arr, cls=True)` 의 `cls=True` 도 3.x 에서 deprecated → 생략
    """
    global _OCR_INSTANCE, _OCR_LANG
    if _OCR_INSTANCE is not None and _OCR_LANG == lang:
        return _OCR_INSTANCE
    # Lazy import — raises ImportError if paddleocr/paddlepaddle not installed
    from paddleocr import PaddleOCR  # noqa: WPS433
    # Try 3.x API first (use_textline_orientation), fallback to 2.x API (use_angle_cls)
    try:
        _OCR_INSTANCE = PaddleOCR(lang=lang, use_textline_orientation=True)
    except TypeError:
        # 2.x fallback
        _OCR_INSTANCE = PaddleOCR(use_angle_cls=True, lang=lang, show_log=False)
    _OCR_LANG = lang
    return _OCR_INSTANCE


def _preprocess_for_ocr(pil_img):
    """v0.7.4.3: opencv pipeline — gray → denoise → deskew → adaptive threshold.
    Returns: PIL.Image in RGB mode (OCR engine expects 3-channel)."""
    try:
        import cv2  # noqa: WPS433
        import numpy as np  # noqa: WPS433
        from PIL import Image  # noqa: WPS433
    except ImportError as e:
        raise ImportError(f"전처리 의존성 미설치 (cv2/numpy/PIL): {e}")

    arr = np.array(pil_img.convert("RGB"))
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    # Denoise (h=10 = moderate strength)
    try:
        gray = cv2.fastNlMeansDenoising(gray, h=10)
    except Exception:
        pass
    # Deskew via minAreaRect of dark pixels
    try:
        coords = np.column_stack(np.where(gray < 200))
        if len(coords) > 100:
            rect = cv2.minAreaRect(coords)
            angle = rect[-1]
            if angle < -45:
                angle = -(90 + angle)
            else:
                angle = -angle
            if abs(angle) > 0.3:
                h, w = gray.shape
                M = cv2.getRotationMatrix2D((w // 2, h // 2), angle, 1.0)
                gray = cv2.warpAffine(
                    gray, M, (w, h),
                    flags=cv2.INTER_CUBIC,
                    borderMode=cv2.BORDER_REPLICATE,
                )
    except Exception:
        pass
    # Adaptive threshold — handles uneven scan illumination
    try:
        gray = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY, 31, 10,
        )
    except Exception:
        pass
    return Image.fromarray(gray).convert("RGB")


def _extract_ocr_blocks(page, page_index: int, opts: Dict[str, Any]) -> Tuple[List[TextBlock], float]:
    """Extract text blocks from a scanned page via PaddleOCR.
    Returns: (blocks, avg_confidence) — avg_confidence ∈ [0.0, 1.0]."""
    try:
        import numpy as np  # noqa: WPS433
        from PIL import Image  # noqa: WPS433
    except ImportError as e:
        raise ImportError(f"OCR 전처리 의존성 미설치 (numpy/PIL): {e}")

    blocks: List[TextBlock] = []
    confidences: List[float] = []

    # Rasterize PDF page → PIL Image at 300 DPI (200 DPI fallback for long PDFs)
    dpi = int(opts.get("ocr_dpi", 300))
    try:
        pix = page.get_pixmap(dpi=dpi, alpha=False)
    except TypeError:
        # Old PyMuPDF API — uses matrix instead of dpi kwarg
        import fitz  # noqa: WPS433
        mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
        pix = page.get_pixmap(matrix=mat, alpha=False)

    try:
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    except Exception as e:
        raise RuntimeError(f"page {page_index} rasterization failed: {e}")

    # Optional preprocessing
    if opts.get("preprocess", True):
        try:
            img = _preprocess_for_ocr(img)
        except Exception as e:
            print(f"[pdf_clone] page {page_index} preprocess failed: {e}", file=sys.stderr)

    # PaddleOCR — v0.7.4.9: 3.x API 에서 `cls=True` deprecated → 생략
    ocr = _get_ocr_singleton(opts.get("lang", "korean"))
    arr = np.array(img)
    try:
        # Try 3.x API first (no cls param)
        raw = ocr.ocr(arr)
    except TypeError:
        # 2.x fallback (with cls=True)
        raw = ocr.ocr(arr, cls=True)
    except Exception as e:
        raise RuntimeError(f"page {page_index} PaddleOCR failed: {e}")

    # raw format: [[[polygon, (text, confidence)], ...]] (outer list = batch, inner = page lines)
    lines = raw[0] if raw and isinstance(raw, list) and len(raw) > 0 else []
    if not lines:
        return blocks, 0.0

    # Scale image coords → PDF points
    scale_x = page.rect.width / pix.width if pix.width > 0 else 1.0
    scale_y = page.rect.height / pix.height if pix.height > 0 else 1.0

    for line in lines:
        if not line or len(line) < 2:
            continue
        polygon = line[0]
        try:
            text, confidence = line[1]
        except (TypeError, ValueError):
            continue
        if not text:
            continue
        text = str(text).strip()
        if not text:
            continue
        conf = float(confidence or 0.0)
        if conf < 0.5:
            continue  # drop low-confidence noise

        confidences.append(conf)

        # Polygon is 4 corner points — compute axis-aligned bbox in image coords
        xs = [float(p[0]) for p in polygon]
        ys = [float(p[1]) for p in polygon]
        x0, y0 = min(xs), min(ys)
        x1, y1 = max(xs), max(ys)

        # Scale to PDF points
        bbox_pdf = (x0 * scale_x, y0 * scale_y, x1 * scale_x, y1 * scale_y)

        # Estimate font size from bbox height (empirical 0.85 factor)
        height_pt = (y1 - y0) * scale_y
        size_pt = round(max(8.0, min(28.0, height_pt * 0.85)), 1)

        blocks.append(TextBlock(
            text=text,
            bbox=bbox_pdf,
            font="",
            size=size_pt,
            bold=False,
            italic=False,
            color=0,
            page=page_index,
        ))

    avg_conf = sum(confidences) / len(confidences) if confidences else 0.0
    return blocks, avg_conf

"""pdf_clone.native — Native PDF 추출 (text, paragraphs, tables, images, header/footer).

함수:
- _extract_native_blocks  : PyMuPDF get_text('dict') → TextBlock list
- _make_paragraph         : lines of TextBlocks → Paragraph (색상/정렬/폰트명 dominant)
- _extract_headers_footers: 상단/하단 band 에서 header/footer 텍스트 추출
- _detect_list_markers    : paragraph 시작의 bullet/번호 패턴 감지
- _detect_tables_native   : pdfplumber find_tables → TableModel list
- _extract_images_native  : get_images + extract_image → PNG + mm 크기 (aspect-ratio 보존 clamp)
"""
import os
import re
import sys
from collections import Counter
from typing import List, Optional, Tuple, Dict

from ._models import (
    TextBlock, Paragraph, TableModel,
    POINTS_TO_MM, MAX_IMG_W_MM, MAX_IMG_H_MM, MIN_IMG_MM,
)


def _extract_native_blocks(page, page_index: int) -> List[TextBlock]:
    """Extract text blocks from a native PDF page via fitz.Page.get_text('dict')."""
    blocks: List[TextBlock] = []
    try:
        raw = page.get_text("dict")
    except Exception as e:
        print(f"[pdf_clone] page {page_index} get_text dict failed: {e}", file=sys.stderr)
        return blocks
    for b in raw.get("blocks", []):
        if b.get("type") != 0:  # 0=text, 1=image
            continue
        for line in b.get("lines", []):
            for span in line.get("spans", []):
                txt = (span.get("text") or "").strip()
                if not txt:
                    continue
                flags = int(span.get("flags", 0))
                bbox = tuple(span.get("bbox", (0, 0, 0, 0)))
                if len(bbox) != 4:
                    continue
                blocks.append(TextBlock(
                    text=txt,
                    bbox=bbox,  # type: ignore
                    font=str(span.get("font", "")),
                    size=float(span.get("size", 10.0)),
                    bold=bool(flags & 16),
                    italic=bool(flags & 2),
                    color=int(span.get("color", 0)),
                    page=page_index,
                ))
    return blocks


def _make_paragraph(lines: List[List[TextBlock]],
                     page_width: Optional[float] = None) -> Paragraph:
    """Merge a list of lines (each line = list of TextBlocks) into one Paragraph.
    v0.7.4.8 Fix Group A: compute alignment, dominant color, font_name from blocks."""
    all_blocks = [b for ln in lines for b in ln]
    if not all_blocks:
        return Paragraph(text="")
    # Join text — lines separated by space (Korean docs don't always need space between lines)
    text_parts: List[str] = []
    for ln in lines:
        line_txt = "".join(b.text for b in ln).strip()
        if line_txt:
            text_parts.append(line_txt)
    full_text = " ".join(text_parts)
    # Median font size
    sizes = sorted(b.size for b in all_blocks)
    median_size = sizes[len(sizes) // 2] if sizes else 10.0
    # Bold/italic: majority vote
    bold_count = sum(1 for b in all_blocks if b.bold)
    italic_count = sum(1 for b in all_blocks if b.italic)
    bold = bold_count * 2 > len(all_blocks)
    italic = italic_count * 2 > len(all_blocks)

    # v0.7.4.8 Fix A1: dominant color — pick most frequent non-zero color across blocks
    color_counts = Counter(b.color for b in all_blocks)
    non_black = {c: n for c, n in color_counts.items() if c != 0}
    dominant_color = 0
    if non_black:
        dominant_color = max(non_black.items(), key=lambda kv: kv[1])[0]

    # v0.7.4.8 Fix A3: dominant font name — most frequent non-empty font
    font_counts = Counter(b.font for b in all_blocks if b.font)
    dominant_font = font_counts.most_common(1)[0][0] if font_counts else ""

    # Bbox extremes
    min_x = min(b.bbox[0] for b in all_blocks)
    max_x = max(b.bbox[2] for b in all_blocks)
    min_y = min(b.bbox[1] for b in all_blocks)
    max_y = max(b.bbox[3] for b in all_blocks)

    # v0.7.4.8 Fix A2: alignment detection from bbox x-position
    align = "left"
    if page_width and page_width > 0:
        para_width = max_x - min_x
        # Estimate body area width (page minus typical 15mm margins each side)
        typical_left_margin = 42.5
        typical_right_margin = 42.5
        body_left = typical_left_margin
        body_right = page_width - typical_right_margin
        body_width = body_right - body_left
        para_center = (min_x + max_x) / 2
        body_center = (body_left + body_right) / 2
        left_gap = min_x - body_left
        right_gap = body_right - max_x

        # Center: paragraph centered within body, both sides have similar gap, narrower than body
        if body_width > 0 and para_width < body_width * 0.85:
            center_offset = abs(para_center - body_center) / body_width
            gap_ratio_diff = abs(left_gap - right_gap) / max(1.0, body_width)
            if center_offset < 0.10 and gap_ratio_diff < 0.15 and left_gap > body_width * 0.10:
                align = "center"
            elif right_gap < body_width * 0.05 and left_gap > body_width * 0.15:
                align = "right"

    # v0.7.4.8 Fix D3: 들여쓰기 계산 — 첫 블록 x 와 페이지 left margin 차이
    left_margin_pt = 0.0
    if page_width and page_width > 0 and min_x > 42.5 + 10:
        left_margin_pt = max(0.0, min_x - 42.5)

    return Paragraph(
        text=full_text,
        font_size=median_size,
        bold=bold,
        italic=italic,
        align=align,
        is_title=False,
        bbox=(min_x, min_y, max_x, max_y),
        color=dominant_color,
        font_name=dominant_font,
        left_margin=left_margin_pt,
    )


def _extract_headers_footers(page, page_index: int,
                              band_height_pt: float = 28.35) -> Dict[str, str]:
    """v0.7.4.8 Fix D2: 페이지 상단/하단 band 의 텍스트를 header/footer 로 추출.

    Args:
        page: fitz.Page 객체
        page_index: 페이지 인덱스 (첫 페이지만 의미 — set_header_footer 는 문서 전체 영향)
        band_height_pt: band 높이 (기본 28.35pt ≈ 10mm)

    Returns: {"header": "...", "footer": "..."} (비어 있으면 빈 문자열)
    """
    result = {"header": "", "footer": ""}
    try:
        rect = page.rect
        raw = page.get_text("dict")
        header_lines = []
        footer_lines = []
        for b in raw.get("blocks", []):
            if b.get("type") != 0:
                continue
            for line in b.get("lines", []):
                line_text = ""
                line_bbox = line.get("bbox", (0, 0, 0, 0))
                for span in line.get("spans", []):
                    line_text += span.get("text", "")
                line_text = line_text.strip()
                if not line_text:
                    continue
                y_top = line_bbox[1]
                y_bot = line_bbox[3]
                # 상단 band: y_top < band_height_pt
                if y_top < band_height_pt:
                    header_lines.append(line_text)
                # 하단 band: y_bot > page_height - band_height_pt
                elif y_bot > (rect.height - band_height_pt):
                    footer_lines.append(line_text)
        result["header"] = " ".join(header_lines).strip()
        result["footer"] = " ".join(footer_lines).strip()
    except Exception as e:
        print(f"[pdf_clone] page {page_index} header/footer extract failed: {e}",
              file=sys.stderr)
    return result


def _detect_list_markers(text: str) -> Tuple[bool, str]:
    """v0.7.4.8 Fix D4: paragraph 텍스트 시작에서 bullet/번호 패턴 감지.

    Returns: (is_list, list_marker)
    """
    if not text:
        return False, ""
    # Bullet markers
    bullet_match = re.match(r'^([•○■▪◆◇●])\s+', text)
    if bullet_match:
        return True, bullet_match.group(1) + " "
    # Numeric markers (1. / 1) / 1-)
    num_match = re.match(r'^(\d{1,3}[.)]\s|\d{1,3}-\s|\d{1,3}\)\s)', text)
    if num_match:
        return True, num_match.group(1)
    # Korean 가. / 나. / 가) (1-3 hangul + dot or paren + space)
    kor_match = re.match(r'^([가-힣]{1,2}[.)]\s|\([가-힣]{1,2}\)\s|\(\d{1,3}\)\s)', text)
    if kor_match:
        return True, kor_match.group(1)
    return False, ""


def _detect_tables_native(pdf_path: str, page_index: int) -> List[TableModel]:
    """v0.7.4.4: native PDF 표 감지 via pdfplumber.Page.find_tables().
    pdfplumber 는 ruled-line + cell 정렬 heuristic 을 사용. 반환: TableModel list."""
    tables: List[TableModel] = []
    try:
        import pdfplumber  # noqa: WPS433
    except ImportError:
        return tables
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_index >= len(pdf.pages):
                return tables
            pg = pdf.pages[page_index]
            found = pg.find_tables() or []
            for t in found:
                try:
                    data = t.extract() or []
                    cells_2d: List[List[str]] = []
                    for row in data:
                        if not row:
                            continue
                        cells_2d.append([
                            (str(c).strip() if c is not None else "") for c in row
                        ])
                    if not cells_2d or all(not any(r) for r in cells_2d):
                        continue
                    # Has header heuristic
                    has_header = False
                    if len(cells_2d) >= 2:
                        has_header = any(c for c in cells_2d[0])
                    bbox_t = None
                    try:
                        bbox_t = (t.bbox[0], t.bbox[1], t.bbox[2], t.bbox[3])
                    except Exception:
                        pass
                    tables.append(TableModel(
                        cells_2d=cells_2d,
                        has_header=has_header,
                        bbox=bbox_t,
                    ))
                except Exception:
                    continue
    except Exception as e:
        print(f"[pdf_clone] page {page_index} find_tables failed: {e}", file=sys.stderr)
    return tables


def _extract_images_native(doc, page, page_index: int,
                             tmp_dir: str) -> List[Tuple[str, float, float]]:
    """v0.7.5.0: native PDF 이미지 추출 + aspect-ratio 보존 clamp.
    Returns: list of (temp_png_path, width_mm, height_mm)."""
    results: List[Tuple[str, float, float]] = []
    try:
        images = page.get_images(full=True) or []
    except Exception:
        return results

    for img_idx, img_info in enumerate(images):
        try:
            xref = img_info[0]
            extracted = doc.extract_image(xref)
            if not extracted:
                continue
            img_bytes = extracted.get("image")
            ext = extracted.get("ext", "png")
            if not img_bytes:
                continue
            filename = f"pdf_clone_p{page_index}_i{img_idx}.{ext}"
            filepath = os.path.join(tmp_dir, filename)
            with open(filepath, "wb") as fh:
                fh.write(img_bytes)
            # Get image bbox on page (if available)
            w_mm = 50.0  # default mm
            h_mm = 50.0
            try:
                rects = page.get_image_rects(xref) or []
                if rects:
                    r = rects[0]
                    w_mm = round(r.width * POINTS_TO_MM, 1)
                    h_mm = round(r.height * POINTS_TO_MM, 1)
            except Exception:
                pass
            # v0.7.5.0 Issue 8: aspect-ratio 보존 clamp
            if w_mm > 0 and h_mm > 0:
                ratio = w_mm / h_mm
                if w_mm > MAX_IMG_W_MM:
                    w_mm = MAX_IMG_W_MM
                    h_mm = round(w_mm / ratio, 1)
                if h_mm > MAX_IMG_H_MM:
                    h_mm = MAX_IMG_H_MM
                    w_mm = round(h_mm * ratio, 1)
                w_mm = max(MIN_IMG_MM, w_mm)
                h_mm = max(MIN_IMG_MM, h_mm)
            else:
                w_mm = max(MIN_IMG_MM, min(MAX_IMG_W_MM, w_mm))
                h_mm = max(MIN_IMG_MM, min(MAX_IMG_H_MM, h_mm))
            results.append((filepath, w_mm, h_mm))
        except Exception as e:
            print(f"[pdf_clone] page {page_index} image {img_idx} extract failed: {e}",
                  file=sys.stderr)
            continue
    return results

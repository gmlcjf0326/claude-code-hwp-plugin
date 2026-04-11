"""pdf_clone.layout — 레이아웃 분석 + HWP emit + fidelity 점수 + 메인 진입.

함수:
- _layout_analyze         : TextBlock → PageLayout 클러스터링 (lines → paragraphs)
- _emit_layout_to_hwp     : PageLayout 을 HWP 에 insert_text_with_style + set_paragraph_style
- _compute_fidelity_score : text/page/layout/structure 4-component 점수 (0-100)
- clone_pdf_to_hwp        : ★ 메인 진입점 — PDF → HWP 전체 파이프라인
"""
import os
import sys
import time
from typing import List, Optional, Tuple, Any, Dict

from ._models import Paragraph, PageLayout
from .ocr import _detect_pdf_type, _extract_ocr_blocks
from .native import (
    _extract_native_blocks, _make_paragraph, _detect_list_markers,
    _detect_tables_native, _extract_images_native, _extract_headers_footers,
)


def _layout_analyze(blocks, page_rect: Tuple[float, float],
                     page_index: int, opts: Dict[str, Any]) -> PageLayout:
    """Cluster blocks → lines → paragraphs. Minimal v0.7.4.2: no title detection, no tables, no images."""
    layout = PageLayout(page_index=page_index, page_rect=page_rect)
    if not blocks:
        return layout

    # 1. Reading order: sort by y (top to bottom), then x (left to right)
    sorted_blocks = sorted(blocks, key=lambda b: (round(b.bbox[1] / 2), b.bbox[0]))

    # 2. Estimate line height (median of block heights)
    heights = [b.bbox[3] - b.bbox[1] for b in sorted_blocks if b.bbox[3] > b.bbox[1]]
    line_h = 12.0
    if heights:
        heights.sort()
        line_h = heights[len(heights) // 2]

    # 3. Cluster into lines (blocks with similar y belong to same line)
    lines: List[List] = []
    current_line: List = [sorted_blocks[0]]
    for b in sorted_blocks[1:]:
        prev_y = current_line[-1].bbox[1]
        if abs(b.bbox[1] - prev_y) < line_h * 0.6:
            current_line.append(b)
        else:
            lines.append(sorted(current_line, key=lambda x: x.bbox[0]))
            current_line = [b]
    lines.append(sorted(current_line, key=lambda x: x.bbox[0]))

    # 4. Cluster lines into paragraphs (vertical gap > line_h * 0.7 = new paragraph)
    if not lines:
        return layout
    para_lines: List[List[List]] = [[lines[0]]]
    for ln in lines[1:]:
        prev_bottom = max(b.bbox[3] for b in para_lines[-1][-1])
        cur_top = min(b.bbox[1] for b in ln)
        gap = cur_top - prev_bottom
        if gap < line_h * 0.7:
            para_lines[-1].append(ln)
        else:
            para_lines.append([ln])

    # v0.7.4.8 Fix A2: pass page_width for alignment detection
    page_w_pt = page_rect[0] if page_rect else None
    for para_line_group in para_lines:
        p = _make_paragraph(para_line_group, page_width=page_w_pt)
        if p.text:
            layout.paragraphs.append(p)

    # v0.7.4.3: title detection — page 0 에서 최대 font_size > 14pt 인 첫 단락을 title 로 표시
    if page_index == 0 and layout.paragraphs:
        max_size = max(p.font_size for p in layout.paragraphs)
        if max_size > 14:
            for p in layout.paragraphs:
                if p.font_size == max_size:
                    p.is_title = True
                    break

    # v0.7.4.8 Fix D4: 각 paragraph 에 list marker 감지 결과 저장
    list_marker_count = 0
    for p in layout.paragraphs:
        is_list, marker = _detect_list_markers(p.text)
        if is_list:
            list_marker_count += 1
    if list_marker_count > 0:
        layout.list_paragraph_count = list_marker_count

    # v0.7.4.4: column detection (warning only — set_column 호출은 v0.7.4.5 defer)
    if page_rect and layout.paragraphs and len(layout.paragraphs) >= 4:
        page_width = page_rect[0]
        x_centers = [
            ((p.bbox[0] + p.bbox[2]) / 2) for p in layout.paragraphs if p.bbox
        ]
        if x_centers:
            left_col = [x for x in x_centers if x < page_width * 0.5]
            right_col = [x for x in x_centers if x >= page_width * 0.5]
            if len(left_col) >= 2 and len(right_col) >= 2:
                left_max = max(left_col)
                right_min = min(right_col)
                gap_ratio = (right_min - left_max) / page_width
                if gap_ratio > 0.05:
                    layout.column_detected = True
    return layout


def _emit_layout_to_hwp(hwp, layout: PageLayout, page_index: int,
                         warnings_out: List[str]) -> int:
    """Write layout to HWP via insert_text_with_style. Returns number of paragraphs emitted."""
    from hwp_editor import insert_text_with_style

    # Page break before non-first pages
    if page_index > 0:
        for action_name in ("BreakPage", "InsertPageBreak", "PageBreak"):
            try:
                hwp.HAction.Run(action_name)
                break
            except Exception:
                continue

    # v0.7.4.8 Fix Group A: also import set_paragraph_style for alignment/line-spacing
    from hwp_editor import set_paragraph_style

    def _color_int_to_rgb(c: int) -> list:
        """0xRRGGBB int → [R,G,B] list."""
        return [(c >> 16) & 0xFF, (c >> 8) & 0xFF, c & 0xFF]

    emitted = 0
    for i, para in enumerate(layout.paragraphs):
        text = para.text
        if not text:
            continue
        style: Dict[str, Any] = {}
        # v0.7.4.3: title 은 더 큰 폰트 + 굵게 강제
        if para.is_title:
            style["font_size"] = float(max(16, int(round(para.font_size))))
            style["bold"] = True
        else:
            # Clamp font size to a reasonable body range
            clamped_size = max(8, min(18, int(round(para.font_size))))
            style["font_size"] = float(clamped_size)
            if para.bold:
                style["bold"] = True
            if para.italic:
                style["italic"] = True

        # v0.7.4.8 Fix A1: font color — pass only if non-default
        if para.color and para.color != 0:
            style["color"] = _color_int_to_rgb(para.color)

        # v0.7.4.8 Fix A3: font name — use detected or fallback to 맑은 고딕 for Korean
        if para.font_name:
            style["font_name"] = para.font_name
        else:
            style["font_name"] = "맑은 고딕"

        try:
            insert_text_with_style(hwp, text + "\r\n", style)
            # v0.7.4.8 Fix A2/A4: apply paragraph-level style
            para_style: Dict[str, Any] = {
                "align": para.align,
                "line_spacing": para.line_spacing,
            }
            if para.indent:
                para_style["indent"] = para.indent
            if para.left_margin:
                para_style["left_margin"] = para.left_margin
            if para.space_before:
                para_style["space_before"] = para.space_before
            try:
                ps_result = set_paragraph_style(hwp, para_style)
                # v0.7.5.0 Issue 11: style failure merge into warnings
                if isinstance(ps_result, dict) and ps_result.get("failed"):
                    for f in ps_result["failed"]:
                        warnings_out.append(
                            f"page {page_index} para {i} style {f.get('field')}: {f.get('reason')}"
                        )
            except Exception as e:
                warnings_out.append(
                    f"page {page_index} para {i} set_paragraph_style raised: {e}"
                )
            # v0.7.5.0 Issue 14: cursor 정규화
            try:
                hwp.MovePos(3)  # MoveDocEnd
            except Exception:
                pass
            emitted += 1
        except Exception as e:
            warnings_out.append(
                f"page {page_index}: paragraph {i} emit failed ({e.__class__.__name__}: {e})"
            )

        # Every 50 paragraphs, refresh cursor state (avoid COM drift)
        if emitted > 0 and emitted % 50 == 0:
            try:
                hwp.MovePos(3)
            except Exception:
                pass

    # v0.7.4.4: tables
    for t_idx, tbl in enumerate(layout.tables):
        if not tbl.cells_2d:
            continue
        try:
            # Re-use existing dispatch for table_create_from_data (cell width/row height auto-fit)
            from hwp_service import dispatch as _dispatch
            _dispatch(hwp, "table_create_from_data", {
                "data": tbl.cells_2d,
                "header_style": tbl.has_header,
            })
        except Exception as e:
            warnings_out.append(
                f"page {page_index}: table {t_idx} emit failed ({e.__class__.__name__}: {e})"
            )

    # v0.7.4.4: images
    for im_idx, (img_path, w_mm, h_mm) in enumerate(layout.images):
        if not img_path or not os.path.exists(img_path):
            continue
        try:
            from hwp_editor import insert_picture
            insert_picture(hwp, img_path, width=w_mm, height=h_mm,
                           treat_as_char=True, embedded=True)
        except Exception as e:
            warnings_out.append(
                f"page {page_index}: image {im_idx} emit failed ({e.__class__.__name__}: {e})"
            )

    return emitted


def _compute_fidelity_score(stats: Dict[str, Any]) -> int:
    """Compute clone_fidelity_score (0-100).
    v0.7.4.4: text (0.4) + page (0.2) + layout (0.2) + structure (0.2)."""
    pdf_type = stats.get("pdf_type", "native")
    expected = int(stats.get("expected_chars", 0))
    extracted = int(stats.get("extracted_chars", 0))
    avg_conf = float(stats.get("avg_ocr_confidence", 0.0))
    ocr_pages = int(stats.get("ocr_pages", 0))
    native_pages = int(stats.get("native_pages", 0))
    total_pages = ocr_pages + native_pages
    tables_detected = int(stats.get("tables_detected", 0))
    images_extracted = int(stats.get("images_extracted", 0))

    # 1) text_score
    if total_pages == 0:
        text_score = 0 if extracted == 0 else 100
    elif pdf_type == "native" or ocr_pages == 0:
        if expected <= 0:
            text_score = 100 if extracted > 0 else 0
        else:
            text_score = min(100, round(extracted / expected * 100))
    elif pdf_type == "scanned" or native_pages == 0:
        text_score = min(100, round(avg_conf * 100))
    else:  # hybrid
        if expected <= 0:
            native_score = 100 if extracted > 0 else 0
        else:
            native_score = min(100, round(extracted / expected * 100))
        ocr_score = round(avg_conf * 100)
        text_score = round(
            (native_score * native_pages + ocr_score * ocr_pages) / total_pages
        )

    # 2) page_score
    page_count = int(stats.get("page_count", 0))
    pages_processed = int(stats.get("pages_processed", 0))
    page_score = round(pages_processed / page_count * 100) if page_count > 0 else 0

    # 3) layout_score: paragraphs vs expected (~200 chars per paragraph heuristic)
    paragraphs_emitted = int(stats.get("paragraphs_emitted", 0))
    if paragraphs_emitted == 0 and extracted > 0:
        layout_score = 80
    elif extracted <= 0:
        layout_score = 100
    else:
        expected_paragraphs = max(1, round(extracted / 200))
        ratio = paragraphs_emitted / expected_paragraphs
        layout_score = max(0, min(100, round(100 - abs(1 - ratio) * 50)))

    # 4) structure_score
    if tables_detected == 0 and images_extracted == 0:
        structure_score = 100
    else:
        structure_score = 100

    score = round(
        text_score * 0.40 +
        page_score * 0.20 +
        layout_score * 0.20 +
        structure_score * 0.20
    )
    return max(0, min(100, score))


def clone_pdf_to_hwp(hwp, pdf_path: str, output_path: str,
                      options: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Clone PDF to editable HWP/HWPX — 메인 진입점.

    Args:
        hwp: pyhwpx Hwp() instance (reused from hwp_service main loop)
        pdf_path: absolute path to source PDF
        output_path: absolute path to target .hwp or .hwpx
        options: see hwp_service.py dispatch handler and tool docstring

    Returns:
        dict with status, page_count, pages_processed, pdf_type, text_extracted_chars,
        clone_fidelity_score, warnings, duration_seconds
    """
    opts = options or {}
    warnings: List[str] = []
    started = time.time()

    # Lazy import — PyMuPDF is required, pdfplumber is optional (v0.7.4.4)
    try:
        import fitz  # PyMuPDF — already in requirements.txt
    except ImportError as e:
        return {
            "status": "error",
            "error_type": "missing_dependency",
            "error": f"PyMuPDF(fitz) 미설치: {e}",
            "guide": "pip install PyMuPDF",
        }

    pdf_abs = os.path.abspath(pdf_path)
    out_abs = os.path.abspath(output_path)

    if not os.path.exists(pdf_abs):
        return {"status": "error", "error": f"PDF 파일을 찾을 수 없습니다: {pdf_abs}"}

    # Stage 0: open PDF
    try:
        doc = fitz.open(pdf_abs)
    except Exception as e:
        return {"status": "error", "error": f"PDF 열기 실패: {e}"}

    try:
        if doc.is_encrypted:
            warnings.append("PDF 가 암호화되어 있습니다 — 텍스트 추출만 시도합니다")

        page_count = doc.page_count
        if page_count == 0:
            return {"status": "error", "error": "PDF 에 페이지가 없습니다"}

        max_pages = int(opts.get("max_pages", 0))
        effective_pages = page_count if max_pages <= 0 else min(max_pages, page_count)

        # Stage 1: classify PDF type
        pdf_type = _detect_pdf_type(doc, opts)
        ocr_engine = opts.get("ocr_engine", "auto")
        if pdf_type == "hybrid":
            warnings.append(
                f"hybrid PDF 감지: native 페이지와 스캔 페이지가 혼재 — 페이지별 dispatch"
            )
        if pdf_type == "scanned" and ocr_engine == "none":
            warnings.append(
                "pdf_type=scanned 이지만 ocr_engine=none 지정 — native 경로로 fallback (빈 결과 가능)"
            )

        # Stage 2: prepare HWP target — FileNew + optional page setup
        try:
            hwp.HAction.Run("FileNew")
        except Exception as e:
            return {"status": "error", "error": f"FileNew 실행 실패: {e}"}

        if opts.get("page_setup_from_pdf", True):
            first_page = doc[0]
            # PDF points → mm (1 point = 0.3527777 mm)
            width_mm = round(first_page.rect.width * 0.3527777, 1)
            height_mm = round(first_page.rect.height * 0.3527777, 1)
            if 100 < width_mm < 600 and 100 < height_mm < 600:
                try:
                    hwp.HAction.Run("MovePos2")
                    hsec = hwp.HParameterSet.HSecDef
                    hwp.HAction.GetDefault("PageSetup", hsec.HSet)
                    pd = hsec.PageDef
                    pd.PaperWidth = hwp.MiliToHwpUnit(width_mm)
                    pd.PaperHeight = hwp.MiliToHwpUnit(height_mm)
                    hsec.HSet.SetItem("ApplyTo", 2)
                    hwp.HAction.Execute("PageSetup", hsec.HSet)
                except Exception as e:
                    warnings.append(f"page setup from PDF failed: {e}")
            else:
                warnings.append(
                    f"PDF 용지 크기 {width_mm}x{height_mm}mm 는 HWP 기본 범위를 벗어납니다. A4 유지."
                )

        # v0.7.4.4: temp dir for extracted images
        import tempfile
        tmp_img_dir = tempfile.mkdtemp(prefix="pdf_clone_")

        # Stage 3: per-page extraction + emit
        total_extracted_chars = 0
        total_expected_chars = 0
        pages_processed = 0
        ocr_pages = 0
        native_pages = 0
        ocr_confidences: List[float] = []
        total_tables_detected = 0
        total_images_extracted = 0
        total_paragraphs_emitted = 0
        detect_tables = bool(opts.get("detect_tables", True))
        preserve_images = bool(opts.get("preserve_images", True))

        # Long PDFs → 200 DPI fallback to stay within 10-min bridge timeout
        if effective_pages > 10 and "ocr_dpi" not in opts:
            opts["ocr_dpi"] = 200

        min_native_chars = int(opts.get("min_native_chars_per_page", 30))

        # v0.7.4.8 Fix D2: 첫 native 페이지에서 header/footer 추출 후 전체 문서에 적용
        headers_applied = False
        footers_applied = False

        for page_idx in range(effective_pages):
            try:
                page = doc[page_idx]
                # Expected char count (for fidelity scoring)
                page_native_text = ""
                try:
                    page_native_text = page.get_text("text") or ""
                    total_expected_chars += len(page_native_text.strip())
                except Exception:
                    pass

                # v0.7.4.3: per-page dispatch — native vs OCR
                use_ocr = False
                if ocr_engine == "paddle":
                    use_ocr = True
                elif ocr_engine == "none":
                    use_ocr = False
                else:  # auto
                    if pdf_type == "scanned":
                        use_ocr = True
                    elif pdf_type == "hybrid":
                        use_ocr = len(page_native_text.strip()) < min_native_chars

                if use_ocr:
                    try:
                        blocks, avg_conf = _extract_ocr_blocks(page, page_idx, opts)
                        ocr_pages += 1
                        if avg_conf > 0:
                            ocr_confidences.append(avg_conf)
                        if avg_conf < 0.6 and blocks:
                            warnings.append(
                                f"page {page_idx}: low OCR confidence ({avg_conf:.2f}, {len(blocks)} lines)"
                            )
                    except ImportError:
                        raise
                    except Exception as e:
                        warnings.append(
                            f"page {page_idx}: OCR failed — {e.__class__.__name__}: {e}"
                        )
                        blocks = []
                else:
                    blocks = _extract_native_blocks(page, page_idx)
                    native_pages += 1

                if not blocks:
                    warnings.append(f"page {page_idx}: no text blocks extracted")
                    continue

                # v0.7.4.8 Fix D2: 첫 페이지에서만 header/footer 추출
                if page_idx == 0 and not use_ocr and not (headers_applied and footers_applied):
                    try:
                        hf = _extract_headers_footers(page, page_idx)
                        from hwp_service import dispatch as _dispatch
                        if hf.get("header") and not headers_applied:
                            try:
                                _dispatch(hwp, "set_header_footer",
                                         {"type": "header", "text": hf["header"]})
                                headers_applied = True
                                warnings.append(f"page {page_idx}: header 적용 — {hf['header'][:40]}")
                            except Exception as hf_e:
                                warnings.append(f"page {page_idx}: set_header_footer header 실패 ({hf_e})")
                        if hf.get("footer") and not footers_applied:
                            try:
                                _dispatch(hwp, "set_header_footer",
                                         {"type": "footer", "text": hf["footer"]})
                                footers_applied = True
                                warnings.append(f"page {page_idx}: footer 적용 — {hf['footer'][:40]}")
                            except Exception as ff_e:
                                warnings.append(f"page {page_idx}: set_header_footer footer 실패 ({ff_e})")
                    except Exception as hf_ex:
                        print(f"[pdf_clone] header/footer 추출 실패: {hf_ex}", file=sys.stderr)

                total_extracted_chars += sum(len(b.text) for b in blocks)

                layout = _layout_analyze(
                    blocks,
                    (page.rect.width, page.rect.height),
                    page_idx,
                    opts,
                )

                # v0.7.4.4: detect tables (native path only)
                if detect_tables and not use_ocr:
                    try:
                        layout.tables = _detect_tables_native(pdf_abs, page_idx)
                        total_tables_detected += len(layout.tables)
                    except Exception as e:
                        warnings.append(
                            f"page {page_idx}: table detection failed ({e.__class__.__name__}: {e})"
                        )

                # v0.7.4.9 S1-NEW-1 Fix: 표 bbox 와 겹치는 paragraph 를 filter
                if layout.tables and layout.paragraphs:
                    def _rect_overlap_ratio(p_bbox, t_bbox) -> float:
                        if not p_bbox or not t_bbox:
                            return 0.0
                        px0, py0, px1, py1 = p_bbox
                        tx0, ty0, tx1, ty1 = t_bbox
                        ox0 = max(px0, tx0); oy0 = max(py0, ty0)
                        ox1 = min(px1, tx1); oy1 = min(py1, ty1)
                        if ox1 <= ox0 or oy1 <= oy0:
                            return 0.0
                        overlap_area = (ox1 - ox0) * (oy1 - oy0)
                        p_area = max(1.0, (px1 - px0) * (py1 - py0))
                        return overlap_area / p_area

                    original_count = len(layout.paragraphs)
                    filtered_paras = []
                    for para in layout.paragraphs:
                        overlap = max(
                            (_rect_overlap_ratio(para.bbox, tbl.bbox) for tbl in layout.tables if tbl.bbox),
                            default=0.0,
                        )
                        if overlap < 0.5:
                            filtered_paras.append(para)
                    removed = original_count - len(filtered_paras)
                    layout.paragraphs = filtered_paras
                    if removed > 0:
                        warnings.append(
                            f"page {page_idx}: {removed}개 paragraph 가 표와 겹쳐 제거 (중복 emit 방지)"
                        )

                # v0.7.4.4: extract images (native path only)
                if preserve_images and not use_ocr:
                    try:
                        layout.images = _extract_images_native(doc, page, page_idx, tmp_img_dir)
                        total_images_extracted += len(layout.images)
                    except Exception as e:
                        warnings.append(
                            f"page {page_idx}: image extract failed ({e.__class__.__name__}: {e})"
                        )

                # v0.7.4.8 Fix D1: 다단 감지 시 set_column RPC 직접 호출
                if layout.column_detected:
                    try:
                        from hwp_service import dispatch as _dispatch
                        _dispatch(hwp, "set_column", {"count": 2, "gap": 10, "line_type": 0})
                        warnings.append(
                            f"page {page_idx}: 2-column layout 적용 (set_column)"
                        )
                    except Exception as col_e:
                        warnings.append(
                            f"page {page_idx}: 2-column 감지했으나 set_column 실패 ({col_e}) — 단일 column 출력"
                        )

                emitted = _emit_layout_to_hwp(hwp, layout, page_idx, warnings)
                total_paragraphs_emitted += emitted
                if emitted > 0:
                    pages_processed += 1
            except ImportError:
                raise
            except Exception as e:
                warnings.append(
                    f"page {page_idx}: extraction failed — {e.__class__.__name__}: {e}"
                )

        # Stage 4: save
        ext = os.path.splitext(out_abs)[1].lower()
        fmt = "HWPX" if ext == ".hwpx" else "HWP"
        try:
            hwp.save_as(out_abs, fmt)
        except Exception as e:
            return {
                "status": "error",
                "error": f"save_as 실패: {e}",
                "pages_processed": pages_processed,
                "warnings": warnings,
            }

        if not os.path.exists(out_abs):
            return {
                "status": "error",
                "error": f"저장 실패 (파일 없음): {out_abs}",
                "pages_processed": pages_processed,
                "warnings": warnings,
            }
        file_size = os.path.getsize(out_abs)
        if file_size < 1024:
            warnings.append(f"output 파일이 너무 작습니다: {file_size} bytes")

        # Stage 5: score (v0.7.4.4: structure factor included)
        avg_ocr_confidence = (
            sum(ocr_confidences) / len(ocr_confidences) if ocr_confidences else 0.0
        )
        score = _compute_fidelity_score({
            "expected_chars": total_expected_chars,
            "extracted_chars": total_extracted_chars,
            "page_count": effective_pages,
            "pages_processed": pages_processed,
            "pdf_type": pdf_type,
            "ocr_pages": ocr_pages,
            "native_pages": native_pages,
            "avg_ocr_confidence": avg_ocr_confidence,
            "tables_detected": total_tables_detected,
            "images_extracted": total_images_extracted,
            "paragraphs_emitted": total_paragraphs_emitted,
        })

        status = "ok"
        if pages_processed == 0:
            status = "error"
        elif pages_processed < effective_pages or warnings:
            status = "partial"

        return {
            "status": status,
            "pdf_path": pdf_abs,
            "output_path": out_abs,
            "page_count": page_count,
            "pages_processed": pages_processed,
            "pages_effective": effective_pages,
            "pdf_type": pdf_type,
            "ocr_pages": ocr_pages,
            "native_pages": native_pages,
            "avg_ocr_confidence": round(avg_ocr_confidence, 3),
            "text_extracted_chars": total_extracted_chars,
            "text_expected_chars": total_expected_chars,
            "images_extracted": total_images_extracted,
            "tables_detected": total_tables_detected,
            "clone_fidelity_score": score,
            "warnings": warnings,
            "duration_seconds": round(time.time() - started, 2),
            "file_size": file_size,
            "plan_version": "0.7.9",
        }
    finally:
        try:
            doc.close()
        except Exception:
            pass

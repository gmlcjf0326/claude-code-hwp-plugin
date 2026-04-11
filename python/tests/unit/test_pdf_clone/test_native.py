"""Unit tests for pdf_clone.native module.

Pure function tests — no PyMuPDF / pdfplumber required for these specific tests.
- _detect_list_markers: regex-based bullet/number detection
- _make_paragraph: alignment + dominant color/font from TextBlocks

v0.7.9 Phase 4 (pdf_clone.py 1211L → 5 files) 회귀 안전망.
"""
import pytest

from pdf_clone._models import TextBlock
from pdf_clone.native import _detect_list_markers, _make_paragraph


# ───────────────────────────────────────────────────────────────────
# _detect_list_markers: paragraph 시작에서 bullet/번호 패턴 감지
# Returns: (is_list, list_marker)
# ───────────────────────────────────────────────────────────────────

class TestDetectListMarkers:
    def test_bullet_circle(self):
        is_list, marker = _detect_list_markers("● 첫 번째 항목")
        assert is_list is True
        assert "●" in marker

    def test_bullet_square(self):
        is_list, marker = _detect_list_markers("■ 항목")
        assert is_list is True
        assert "■" in marker

    def test_bullet_diamond(self):
        is_list, marker = _detect_list_markers("◆ 항목")
        assert is_list is True

    def test_numeric_dot(self):
        is_list, marker = _detect_list_markers("1. 첫 항목")
        assert is_list is True
        assert "1." in marker

    def test_numeric_paren(self):
        is_list, marker = _detect_list_markers("1) 첫 항목")
        assert is_list is True

    def test_korean_dot(self):
        is_list, marker = _detect_list_markers("가. 첫 항목")
        assert is_list is True
        assert "가." in marker

    def test_korean_paren(self):
        is_list, marker = _detect_list_markers("(가) 첫 항목")
        assert is_list is True

    def test_paren_decimal(self):
        is_list, marker = _detect_list_markers("(1) 첫 항목")
        assert is_list is True

    def test_no_marker(self):
        is_list, marker = _detect_list_markers("일반 단락 텍스트")
        assert is_list is False
        assert marker == ""

    def test_empty(self):
        is_list, marker = _detect_list_markers("")
        assert is_list is False
        assert marker == ""

    def test_no_space_after_marker(self):
        # marker 뒤에 공백 없으면 매칭 안 됨 (regex 제약)
        is_list, marker = _detect_list_markers("●항목")
        assert is_list is False


# ───────────────────────────────────────────────────────────────────
# _make_paragraph: list of lines (each line = list of TextBlocks)
# → Paragraph (alignment, dominant color, font detection)
# ───────────────────────────────────────────────────────────────────

def _block(text, x0=0, y0=0, x1=100, y1=12, font="", size=10.0,
           bold=False, italic=False, color=0):
    """Helper: TextBlock 생성."""
    return TextBlock(
        text=text,
        bbox=(float(x0), float(y0), float(x1), float(y1)),
        font=font,
        size=size,
        bold=bold,
        italic=italic,
        color=color,
    )


class TestMakeParagraph:
    def test_empty_lines(self):
        para = _make_paragraph([])
        assert para.text == ""

    def test_single_block(self):
        line = [_block("hello")]
        para = _make_paragraph([line])
        assert para.text == "hello"

    def test_multi_line_joined_with_space(self):
        line1 = [_block("first")]
        line2 = [_block("second")]
        para = _make_paragraph([line1, line2])
        assert "first" in para.text
        assert "second" in para.text

    def test_median_font_size(self):
        # 3 blocks: 10, 12, 14 → median 12
        line = [
            _block("a", size=10.0),
            _block("b", size=12.0),
            _block("c", size=14.0),
        ]
        para = _make_paragraph([line])
        assert para.font_size == 12.0

    def test_bold_majority(self):
        # 3 blocks: 2 bold, 1 not → bold = True (majority)
        line = [
            _block("a", bold=True),
            _block("b", bold=True),
            _block("c", bold=False),
        ]
        para = _make_paragraph([line])
        assert para.bold is True

    def test_bold_minority(self):
        # 1 bold, 2 not → bold = False
        line = [
            _block("a", bold=True),
            _block("b", bold=False),
            _block("c", bold=False),
        ]
        para = _make_paragraph([line])
        assert para.bold is False

    def test_dominant_color_non_black(self):
        # 검정(0) + 빨강 + 빨강 → dominant 빨강
        red = 0xFF0000
        line = [
            _block("a", color=0),
            _block("b", color=red),
            _block("c", color=red),
        ]
        para = _make_paragraph([line])
        assert para.color == red

    def test_dominant_color_all_black(self):
        # 모두 검정 → dominant 0
        line = [
            _block("a", color=0),
            _block("b", color=0),
        ]
        para = _make_paragraph([line])
        assert para.color == 0

    def test_dominant_font_name(self):
        # "맑은 고딕" 2개, "함초롬바탕" 1개 → 맑은 고딕
        line = [
            _block("a", font="맑은 고딕"),
            _block("b", font="맑은 고딕"),
            _block("c", font="함초롬바탕"),
        ]
        para = _make_paragraph([line])
        assert para.font_name == "맑은 고딕"

    def test_empty_font_skipped(self):
        # font 가 빈 문자열인 블록은 카운트 안 함
        line = [
            _block("a", font=""),
            _block("b", font="Arial"),
            _block("c", font=""),
        ]
        para = _make_paragraph([line])
        assert para.font_name == "Arial"

    def test_alignment_left_default(self):
        # page_width 없으면 left
        line = [_block("text", x0=42.5, x1=200)]
        para = _make_paragraph([line])
        assert para.align == "left"

    def test_alignment_center_detection(self):
        # 페이지 폭 600pt 가정, body width = 600 - 85 = 515
        # paragraph bbox 가 가운데에 있고 양쪽 gap 비슷하면 center
        page_width = 600.0
        # 가운데 200pt 폭 paragraph (left=200, right=400)
        line = [_block("centered", x0=200, x1=400)]
        para = _make_paragraph([line], page_width=page_width)
        # body_left=42.5, body_right=557.5, body_width=515
        # left_gap = 200-42.5=157.5, right_gap = 557.5-400=157.5
        # para_center = 300, body_center = 300
        # center_offset = 0/515 = 0 < 0.10 ✓
        # gap_ratio_diff = 0/515 = 0 < 0.15 ✓
        # left_gap > body_width*0.10 (51.5) ✓
        assert para.align == "center"

    def test_bbox_extremes(self):
        # 여러 block 의 bbox 가 합쳐져야 함
        line = [
            _block("a", x0=10, y0=20, x1=50, y1=30),
            _block("b", x0=60, y0=15, x1=100, y1=25),
        ]
        para = _make_paragraph([line])
        assert para.bbox is not None
        min_x, min_y, max_x, max_y = para.bbox
        assert min_x == 10
        assert min_y == 15
        assert max_x == 100
        assert max_y == 30

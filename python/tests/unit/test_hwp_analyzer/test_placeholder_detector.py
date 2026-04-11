"""Unit tests for hwp_analyzer.placeholder_detector (Phase 5B).

Rule 기반 placeholder 감지의 R1-R6 rule 정확도 검증.
"""
import pytest

from hwp_analyzer.placeholder_detector import (
    BULLET_MARKERS,
    detect_placeholders,
    detect_primary_marker,
)


# ============================================================================
# R1 bare_marker
# ============================================================================
class TestR1BareMarker:
    def test_bare_marker_simple(self):
        result = detect_placeholders("◦")
        assert any(p["type"] == "bare_marker" for p in result)

    def test_bare_marker_with_space(self):
        result = detect_placeholders("◦ \n  - \n*")
        types = [p["type"] for p in result]
        assert types.count("bare_marker") == 3

    def test_not_bare_marker_with_text(self):
        result = detect_placeholders("◦ 본 사업은 AI 기반")
        types = [p["type"] for p in result]
        assert "bare_marker" not in types

    def test_hollow_circle_also_detected(self):
        result = detect_placeholders("○")
        assert any(p["type"] == "bare_marker" for p in result)


# ============================================================================
# R2 colon_label
# ============================================================================
class TestR2ColonLabel:
    def test_colon_label_simple(self):
        result = detect_placeholders("◦ 안정성 :")
        assert any(p["type"] == "colon_label" for p in result)

    def test_colon_label_fullwidth_colon(self):
        result = detect_placeholders("◦ 성장성：")
        assert any(p["type"] == "colon_label" for p in result)

    def test_colon_label_too_long(self):
        long_label = "◦ " + "가" * 20 + " :"
        result = detect_placeholders(long_label)
        types = [p["type"] for p in result]
        assert "colon_label" not in types

    def test_not_colon_label_with_body_after(self):
        # colon 이 있어도 뒤에 본문이 있으면 placeholder 아님
        result = detect_placeholders("◦ 안정성 : 의료 산업은 경기 무관")
        types = [p["type"] for p in result]
        assert "colon_label" not in types


# ============================================================================
# R3 bracket_heading
# ============================================================================
class TestR3BracketHeading:
    def test_bracket_heading_simple(self):
        result = detect_placeholders("[인프라 측면]")
        assert any(p["type"] == "bracket_heading" for p in result)

    def test_bracket_heading_u22c5(self):
        result = detect_placeholders("[과학⋅기술적 측면]")
        assert any(p["type"] == "bracket_heading" for p in result)

    def test_bracket_heading_u2024(self):
        result = detect_placeholders("[경제적․사회적 측면]")
        assert any(p["type"] == "bracket_heading" for p in result)

    def test_fullwidth_bracket(self):
        result = detect_placeholders("【주의사항】")
        assert any(p["type"] == "bracket_heading" for p in result)


# ============================================================================
# R4 numbered_section (v0.8.2: was number_short)
# ============================================================================
class TestR4NumberedSection:
    def test_number_paren(self):
        result = detect_placeholders("1) 투자계획")
        assert any(p["type"] == "numbered_section" for p in result)

    def test_number_double_paren(self):
        result = detect_placeholders("(1) 산업")
        assert any(p["type"] == "numbered_section" for p in result)

    def test_korean_alpha(self):
        result = detect_placeholders("가. 추진")
        assert any(p["type"] == "numbered_section" for p in result)

    def test_circled_number(self):
        result = detect_placeholders("① 첫째")
        assert any(p["type"] == "numbered_section" for p in result)


# ============================================================================
# R5 marker_section (v0.8.2: was marker_short)
# ============================================================================
class TestR5MarkerSection:
    def test_marker_section_simple(self):
        result = detect_placeholders("◦ 해외 동향")
        assert any(p["type"] == "marker_section" for p in result)

    def test_marker_section_dash(self):
        result = detect_placeholders("- 국내 현황")
        assert any(p["type"] == "marker_section" for p in result)

    def test_marker_not_section_if_too_long(self):
        long_text = "◦ " + "가" * 40
        result = detect_placeholders(long_text)
        types = [p["type"] for p in result]
        assert "marker_section" not in types


# ============================================================================
# R6 blank
# ============================================================================
class TestR6Blank:
    def test_blank_line(self):
        result = detect_placeholders("\n\n")
        blanks = [p for p in result if p["type"] == "blank"]
        assert len(blanks) >= 2

    def test_whitespace_only(self):
        result = detect_placeholders("   \n\t")
        blanks = [p for p in result if p["type"] == "blank"]
        assert len(blanks) >= 1


# ============================================================================
# Primary marker detection
# ============================================================================
class TestPrimaryMarker:
    def test_primary_marker_filled(self):
        text = "◦ A\n◦ B\n◦ C\n- D"
        result = detect_primary_marker(text)
        assert result["primary"] == "◦"
        assert result["counts"].get("◦", 0) == 3

    def test_primary_marker_dash(self):
        text = "- A\n- B\n- C\n◦ D"
        result = detect_primary_marker(text)
        assert result["primary"] == "-"

    def test_empty_text_default(self):
        result = detect_primary_marker("")
        assert result["primary"] == "◦"


# ============================================================================
# Integration: 혼합 양식
# ============================================================================
class TestIntegration:
    def test_mixed_form(self):
        text = """
◦ 본 사업은 AI 기반 진단 플랫폼
- 글로벌 시장: 5000억 달러
[인프라 측면]
◦ 안정성 :
1) 투자계획
◦ 해외 동향
◦
""".strip()
        result = detect_placeholders(text)
        types = [p["type"] for p in result]
        # bracket_heading + colon_label + numbered_section + marker_section + bare_marker 모두 감지
        assert "bracket_heading" in types
        assert "colon_label" in types
        assert "numbered_section" in types
        assert "marker_section" in types
        assert "bare_marker" in types

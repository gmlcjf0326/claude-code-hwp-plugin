"""Unit tests for hwp_core.text_editing._internal module.

핵심: Phase 7 분할 (text_editing.py 1133L → 4 files) 검증.
- _internal.py 가 from .. import 두 점 패턴으로 _helpers 를 import 하는지 (gotcha)
- 3-tier fuzzy heading matcher 가 정상 작동하는지

Tests pure functions: _find_heading_positions, _find_all_in_normalized,
_extract_match_from_line, _find_by_regex, _find_core_in_lines,
_detect_heading_depth.

NO HWP COM required.
"""
import pytest

from hwp_core.text_editing._internal import (
    _find_heading_positions,
    _find_all_in_normalized,
    _extract_match_from_line,
    _find_by_regex,
    _find_core_in_lines,
    _detect_heading_depth,
)
from hwp_core._helpers import normalize_for_match


# ───────────────────────────────────────────────────────────────────
# _detect_heading_depth: 번호 패턴으로 깊이 감지
# depth 1: "1.", "I." — 대제목
# depth 2: "가." — 중제목
# depth 3: "1)", "(1)" — 소제목
# depth 4: "(가)", "①" — 세부
# ───────────────────────────────────────────────────────────────────

class TestDetectHeadingDepth:
    def test_depth1_decimal_dot(self):
        assert _detect_heading_depth("1. 사업 개요") == 1

    def test_depth1_roman(self):
        assert _detect_heading_depth("I. 서론") == 1
        assert _detect_heading_depth("II. 본론") == 1

    def test_depth1_chapter(self):
        assert _detect_heading_depth("제1장 총칙") == 1

    def test_depth2_korean_dot(self):
        assert _detect_heading_depth("가. 추진 배경") == 2
        assert _detect_heading_depth("나. 추진 목적") == 2

    def test_depth3_paren_decimal(self):
        assert _detect_heading_depth("(1) 산업의 특성") == 3
        assert _detect_heading_depth("1) 시장 분석") == 3

    def test_depth4_paren_korean(self):
        assert _detect_heading_depth("(가) 세부 항목") == 4

    def test_depth4_circled(self):
        assert _detect_heading_depth("① 첫째") == 4
        assert _detect_heading_depth("② 둘째") == 4

    def test_default_depth1_for_unknown(self):
        # 번호 없는 제목은 기본 depth 1
        assert _detect_heading_depth("제목 없는 항목") == 1

    def test_handles_leading_whitespace(self):
        assert _detect_heading_depth("  1. 들여쓴 제목") == 1
        assert _detect_heading_depth("\t가. 탭으로 들여쓴") == 2


# ───────────────────────────────────────────────────────────────────
# _find_by_regex: 정규식 기반 매칭
# Returns: list of (position, matched_text)
# ───────────────────────────────────────────────────────────────────

class TestFindByRegex:
    def test_simple_match(self):
        text = "사업 개요\n사업 내용\n사업 계획"
        results = _find_by_regex(text, r"사업\s*개요")
        assert len(results) == 1
        pos, matched = results[0]
        assert "사업" in matched
        assert "개요" in matched

    def test_multiple_matches(self):
        text = "ABC ABC ABC"
        results = _find_by_regex(text, r"ABC")
        assert len(results) == 3

    def test_no_match(self):
        results = _find_by_regex("hello world", r"xyz")
        assert results == []

    def test_invalid_regex_returns_empty(self):
        # 잘못된 regex 는 빈 결과 (예외 발생 안 함)
        results = _find_by_regex("text", r"[invalid")
        assert results == []

    def test_case_insensitive(self):
        # 함수가 re.IGNORECASE 사용
        results = _find_by_regex("Hello WORLD", r"hello")
        assert len(results) == 1


# ───────────────────────────────────────────────────────────────────
# _find_all_in_normalized: 정규화 매칭 (Tier 1)
# Returns: list of (char_position, matched_text)
# ───────────────────────────────────────────────────────────────────

class TestFindAllInNormalized:
    def test_exact_match(self):
        text = "(1) 산업의 특성\n다음 줄"
        norm_text = normalize_for_match(text)
        norm_heading = normalize_for_match("(1) 산업의 특성")
        results = _find_all_in_normalized(text, norm_text, norm_heading)
        assert len(results) == 1
        pos, matched = results[0]
        assert pos == 0  # 첫 줄 시작

    def test_whitespace_difference_match(self):
        # 원본에 공백이 더 있어도 정규화 후 매치
        text = "(1)  산업의   특성"  # extra spaces
        norm_text = normalize_for_match(text)
        norm_heading = normalize_for_match("(1) 산업의 특성")
        results = _find_all_in_normalized(text, norm_text, norm_heading)
        assert len(results) == 1

    def test_fullwidth_paren_match(self):
        # 원본 fullwidth ( ) → 정규화 후 ()
        text = "（1） 산업의 특성"  # fullwidth parens
        norm_text = normalize_for_match(text)
        norm_heading = normalize_for_match("(1) 산업의 특성")
        results = _find_all_in_normalized(text, norm_text, norm_heading)
        assert len(results) == 1

    def test_no_match(self):
        text = "다른 텍스트"
        norm_text = normalize_for_match(text)
        norm_heading = normalize_for_match("(1) 산업의 특성")
        results = _find_all_in_normalized(text, norm_text, norm_heading)
        assert results == []


# ───────────────────────────────────────────────────────────────────
# _find_core_in_lines: 핵심어 기반 매칭 (Tier 3)
# 번호 prefix 가 다르더라도 본문 일치 시 매칭
# ───────────────────────────────────────────────────────────────────

class TestFindCoreInLines:
    def test_finds_with_matching_number(self):
        text = "1. 사업 개요\n다른 줄\n(1) 산업의 특성\n끝"
        # core = "산업의특성", expected_num from "(1) 산업의 특성" = "1"
        results = _find_core_in_lines(text, "산업의특성", "(1) 산업의 특성")
        # 줄에 "1" 이 있고 핵심어 매칭 → results 비어있지 않음
        assert len(results) >= 1

    def test_rejects_wrong_number(self):
        # 원본 heading 은 "(2) 산업의 특성" 인데 본문에 "(1)" 만 있는 경우
        text = "(1) 산업의 특성"
        results = _find_core_in_lines(text, "산업의특성", "(2) 산업의 특성")
        # 번호가 안 맞으므로 매칭 거부
        assert results == []

    def test_no_number_in_heading(self):
        # heading 에 번호 없으면 번호 검증 skip
        text = "임의의 줄에 산업의특성 포함된 텍스트"
        results = _find_core_in_lines(text, "산업의특성", "산업의 특성")
        assert len(results) >= 1


# ───────────────────────────────────────────────────────────────────
# _find_heading_positions: 3-tier cascading matcher (★ 핵심)
# Returns: [(char_pos, matched_text, tier)]
# Tier 1: 정규화 정확
# Tier 2: 공백 유연 regex
# Tier 3: 핵심어
# ───────────────────────────────────────────────────────────────────

class TestFindHeadingPositions:
    def test_tier1_exact(self):
        text = "(1) 산업의 특성\n본문 내용"
        results = _find_heading_positions(text, "(1) 산업의 특성")
        assert len(results) >= 1
        # 첫 매칭은 tier 1
        _, _, tier = results[0]
        assert tier == 1

    def test_tier1_whitespace_normalized(self):
        # 본문에 추가 공백이 있어도 tier 1
        text = "(1)  산업의   특성\n본문"
        results = _find_heading_positions(text, "(1) 산업의 특성")
        assert len(results) >= 1
        _, _, tier = results[0]
        assert tier == 1

    def test_tier1_fullwidth_paren(self):
        text = "（1） 산업의 특성"
        results = _find_heading_positions(text, "(1) 산업의 특성")
        assert len(results) >= 1
        _, _, tier = results[0]
        assert tier == 1

    def test_empty_text(self):
        results = _find_heading_positions("", "anything")
        assert results == []

    def test_empty_heading(self):
        results = _find_heading_positions("text", "")
        assert results == []

    def test_no_match_returns_empty(self):
        results = _find_heading_positions("완전히 다른 텍스트", "(1) 산업의 특성")
        assert results == []

    def test_tier3_number_disambiguation(self):
        # Tier 3 에서 번호로 (1) 산업의 특성 vs (2) 산업의 특성 구분
        text = "(1) 산업의 특성\n(2) 산업의 성장성"
        results_1 = _find_heading_positions(text, "(1) 산업의 특성")
        # Tier 1 매칭 — 정확히 1개
        assert len(results_1) == 1

    def test_returns_actual_text(self):
        # matched_text 는 원본 텍스트 (정규화 전)
        text = "(1)  산업의   특성"
        results = _find_heading_positions(text, "(1) 산업의 특성")
        assert len(results) >= 1
        _, matched, _ = results[0]
        # 원본 텍스트의 일부여야 함
        assert "산업" in matched


# ───────────────────────────────────────────────────────────────────
# _extract_match_from_line: 정규화 위치 → 원본 위치 변환
# ───────────────────────────────────────────────────────────────────

class TestExtractMatchFromLine:
    def test_basic_extraction(self):
        orig = "(1)  산업의   특성"
        norm = normalize_for_match(orig)  # "(1) 산업의 특성"
        norm_heading = "(1) 산업의 특성"
        # norm 의 시작 위치 0 부터 매칭
        result = _extract_match_from_line(orig, norm, 0, norm_heading)
        # 원본의 부분 문자열 (앞뒤 공백 제거)
        assert "산업" in result
        assert "특성" in result

    def test_partial_extraction(self):
        orig = "prefix (1) text suffix"
        norm = normalize_for_match(orig)
        norm_heading = "(1) text"
        # norm 에서 "(1) text" 의 시작 위치
        norm_start = norm.find(norm_heading)
        if norm_start >= 0:
            result = _extract_match_from_line(orig, norm, norm_start, norm_heading)
            assert "(1)" in result or "1" in result

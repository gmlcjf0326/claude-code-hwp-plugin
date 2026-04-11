"""Unit tests for hwp_analyzer.label module.

Tests pure functions: _normalize, _canonical_label, _match_label, classify_table_type.
No HWP COM required — runs on pure Python.

v0.7.9 Phase 10: Phase 5 (hwp_analyzer.py 705L → 5 files) 회귀 안전망.
"""
import pytest

from hwp_analyzer.label import (
    _normalize,
    _canonical_label,
    _match_label,
    classify_table_type,
)


# ───────────────────────────────────────────────────────────────────
# _normalize: 모든 공백 제거
# ───────────────────────────────────────────────────────────────────

class TestNormalize:
    def test_removes_spaces(self):
        assert _normalize("기 업 명") == "기업명"

    def test_removes_tabs(self):
        assert _normalize("기업\t명") == "기업명"

    def test_removes_newlines(self):
        assert _normalize("기업\n명") == "기업명"

    def test_removes_mixed_whitespace(self):
        assert _normalize("  기  업\t명\n  ") == "기업명"

    def test_empty_string(self):
        assert _normalize("") == ""

    def test_no_whitespace(self):
        assert _normalize("기업명") == "기업명"


# ───────────────────────────────────────────────────────────────────
# _canonical_label: 별칭 → 표준명 매핑
# ───────────────────────────────────────────────────────────────────

class TestCanonicalLabel:
    def test_canonical_returns_self(self):
        # 정규명 자체는 그대로 반환
        assert _canonical_label("기업명") == "기업명"

    def test_alias_to_canonical(self):
        # "회사명" 은 "기업명" 의 별칭
        assert _canonical_label("회사명") == "기업명"

    def test_alias_with_whitespace(self):
        # 공백 제거 후 매칭
        assert _canonical_label("회 사 명") == "기업명"

    def test_unknown_label_returns_normalized(self):
        # 별칭에 없는 라벨은 정규화된 원본 반환
        assert _canonical_label("커스텀필드") == "커스텀필드"

    def test_address_aliases(self):
        # 사업장주소의 다양한 alias
        assert _canonical_label("주소") == "사업장주소"
        assert _canonical_label("소재지") == "사업장주소"
        assert _canonical_label("본사") == "사업장주소"

    def test_ceo_aliases(self):
        # 대표자성명의 다양한 alias
        assert _canonical_label("대표자") == "대표자성명"
        assert _canonical_label("대표이사") == "대표자성명"
        assert _canonical_label("CEO") == "대표자성명"


# ───────────────────────────────────────────────────────────────────
# _match_label: 셀 텍스트 vs 검색 라벨 매칭
# Returns: (is_match, is_exact, ratio)
# ───────────────────────────────────────────────────────────────────

class TestMatchLabel:
    def test_exact_match(self):
        is_match, is_exact, ratio = _match_label("기업명", "기업명")
        assert is_match is True
        assert is_exact is True
        assert ratio == 1.0

    def test_whitespace_normalized_match(self):
        # 공백만 다른 경우 — exact match
        is_match, is_exact, ratio = _match_label("기 업 명", "기업명")
        assert is_match is True
        assert is_exact is True
        assert ratio == 1.0

    def test_alias_match(self):
        # 별칭 매칭
        is_match, is_exact, ratio = _match_label("회사명", "기업명")
        assert is_match is True
        assert is_exact is True  # 별칭도 exact 로 간주
        assert ratio == 1.0

    def test_partial_match_substring(self):
        # 셀에 라벨이 포함됨
        is_match, is_exact, ratio = _match_label("기업명(법인)", "기업명")
        assert is_match is True
        assert is_exact is False
        assert 0.0 < ratio < 1.0

    def test_no_match(self):
        is_match, is_exact, ratio = _match_label("전화번호", "기업명")
        assert is_match is False
        assert is_exact is False
        assert ratio == 0.0

    def test_empty_inputs(self):
        is_match, _, ratio = _match_label("", "기업명")
        assert is_match is False
        assert ratio == 0.0
        is_match, _, ratio = _match_label("기업명", "")
        assert is_match is False
        assert ratio == 0.0

    def test_partial_returns_lower_ratio_for_longer_cell(self):
        # 짧은 라벨이 긴 셀에 포함되면 ratio 가 낮아짐
        _, _, ratio_short_in_long = _match_label("기업명입니다정말로긴문자열", "기업명")
        _, _, ratio_short_in_short = _match_label("기업명(주식)", "기업명")
        assert ratio_short_in_long < ratio_short_in_short


# ───────────────────────────────────────────────────────────────────
# classify_table_type: 표 유형 자동 분류
# Returns: 'info_form' | 'financial' | 'timeline' | 'checklist' |
#          'comparison' | 'guide' | 'data_table'
# ───────────────────────────────────────────────────────────────────

class TestClassifyTableType:
    def test_guide_keyword_priority(self):
        # 작성요령 키워드는 최우선
        table = {
            "headers": ["작성요령", "내용"],
            "data": [["1페이지 이내", "기재사항"]],
            "rows": 2,
            "cols": 2,
        }
        assert classify_table_type(table) == "guide"

    def test_financial_table(self):
        table = {
            "headers": ["항목", "금액(원)", "비중(%)"],
            "data": [
                ["인건비", "1,000,000", "50"],
                ["재료비", "500,000", "25"],
            ],
            "rows": 3,
            "cols": 3,
        }
        assert classify_table_type(table) == "financial"

    def test_timeline_table(self):
        table = {
            "headers": ["월", "추진일정", "수행내용"],
            "data": [
                ["M+1", "기획", "분기 1차 마일스톤"],
                ["M+3", "개발", "차년도 준비"],
            ],
            "rows": 3,
            "cols": 3,
        }
        assert classify_table_type(table) == "timeline"

    def test_comparison_table(self):
        table = {
            "headers": ["국내", "해외", "합계"],
            "data": [
                ["100", "200", "300"],
                ["50", "150", "200"],
            ],
            "rows": 3,
            "cols": 3,
        }
        assert classify_table_type(table) == "comparison"

    def test_info_form_pattern(self):
        # 첫 열이 짧은 한글 라벨, 다른 열은 값
        table = {
            "headers": [],
            "data": [
                ["기업명", "(주)테스트"],
                ["대표자", "홍길동"],
                ["연락처", "02-1234-5678"],
                ["주소", "서울시 강남구"],
            ],
            "rows": 4,
            "cols": 2,
        }
        result = classify_table_type(table)
        # 짧은 한글 라벨 + 값 조합 → info_form
        assert result == "info_form"

    def test_data_table_fallback(self):
        # 분류 안 되는 일반 데이터 표
        table = {
            "headers": ["X", "Y"],
            "data": [["1", "2"], ["3", "4"]],
            "rows": 3,
            "cols": 2,
        }
        result = classify_table_type(table)
        # data_table 또는 info_form (양쪽 다 OK — fallback 동작)
        assert result in ("data_table", "info_form")

    def test_empty_table(self):
        # 빈 표는 fallback
        table = {"headers": [], "data": [], "rows": 0, "cols": 0}
        assert classify_table_type(table) == "data_table"

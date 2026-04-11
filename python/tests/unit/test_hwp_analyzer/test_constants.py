"""Unit tests for hwp_analyzer._constants module.

핵심 검증:
- _LABEL_ALIASES 사전 정합성 (round-trip)
- _ALIAS_LOOKUP 역방향 매핑 정확성
- _TABLE_TYPE_KEYWORDS 구조 검증
- MAX_TABLES 안전 한계
"""
import pytest

from hwp_analyzer._constants import (
    _LABEL_ALIASES,
    _ALIAS_LOOKUP,
    _TABLE_TYPE_KEYWORDS,
    MAX_TABLES,
)
from hwp_analyzer.label import _normalize, _canonical_label


class TestLabelAliases:
    def test_dict_not_empty(self):
        assert len(_LABEL_ALIASES) >= 25  # 27 canonical 라벨

    def test_canonical_keys_are_strings(self):
        for canonical in _LABEL_ALIASES.keys():
            assert isinstance(canonical, str)
            assert len(canonical) > 0

    def test_aliases_are_lists(self):
        for canonical, aliases in _LABEL_ALIASES.items():
            assert isinstance(aliases, list), f"{canonical} aliases not a list"
            assert all(isinstance(a, str) for a in aliases)

    def test_business_terms_present(self):
        # v0.7.4.8 Fix B2: 사업계획서 전용 확장 라벨이 있어야 함
        expected = ["주요사업", "사업목표", "사업계획", "기술현황",
                    "자금조달계획", "경쟁분석", "시장규모"]
        for term in expected:
            assert term in _LABEL_ALIASES, f"Missing business term: {term}"

    def test_basic_company_info_present(self):
        # 기본 기업 정보 라벨
        expected = ["기업명", "사업자등록번호", "대표자성명", "사업장주소", "연락처".replace("연락처", "대표전화번호")]
        for term in expected:
            assert term in _LABEL_ALIASES, f"Missing basic term: {term}"


class TestAliasLookup:
    def test_lookup_not_empty(self):
        assert len(_ALIAS_LOOKUP) >= 100  # 200+ aliases expected

    def test_round_trip_canonical(self):
        # 모든 canonical 라벨이 자기 자신에 매핑되어야 함
        for canonical in _LABEL_ALIASES.keys():
            norm = _normalize(canonical)
            assert _ALIAS_LOOKUP.get(norm) == norm, f"{canonical} not self-mapped"

    def test_round_trip_aliases(self):
        # 모든 alias 가 canonical 로 매핑되어야 함
        for canonical, aliases in _LABEL_ALIASES.items():
            norm_canonical = _normalize(canonical)
            for alias in aliases:
                norm_alias = _normalize(alias)
                mapped = _ALIAS_LOOKUP.get(norm_alias)
                assert mapped == norm_canonical, (
                    f"{alias} → {mapped} (expected {norm_canonical})"
                )

    def test_known_aliases(self):
        # 대표 케이스 5 개
        cases = [
            ("회사명", "기업명"),
            ("대표이사", "대표자성명"),
            ("주소", "사업장주소"),
            ("매출", "매출액"),
            ("기술력", "기술현황"),
        ]
        for alias, expected_canonical in cases:
            assert _canonical_label(alias) == _normalize(expected_canonical)


class TestTableTypeKeywords:
    def test_categories_present(self):
        expected_types = ["financial", "timeline", "checklist", "comparison", "guide"]
        for ttype in expected_types:
            assert ttype in _TABLE_TYPE_KEYWORDS

    def test_all_keywords_are_strings(self):
        for ttype, keywords in _TABLE_TYPE_KEYWORDS.items():
            assert isinstance(keywords, list)
            assert all(isinstance(kw, str) for kw in keywords)
            assert len(keywords) > 0

    def test_financial_keywords_meaningful(self):
        kws = _TABLE_TYPE_KEYWORDS["financial"]
        # 실제 금액/매출 관련 키워드가 있어야 함
        assert any(k in kws for k in ["금액", "매출", "비용"])

    def test_guide_keywords_minimal(self):
        # guide 는 매우 좁은 키워드 (작성요령, 유의사항)
        kws = _TABLE_TYPE_KEYWORDS["guide"]
        assert "작성요령" in kws or "유의사항" in kws


class TestMaxTables:
    def test_max_tables_reasonable(self):
        # 통장사본 등 반복 표 방지를 위한 상한
        assert isinstance(MAX_TABLES, int)
        assert 10 <= MAX_TABLES <= 200  # 합리적인 범위

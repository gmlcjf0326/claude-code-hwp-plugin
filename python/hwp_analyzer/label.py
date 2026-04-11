"""hwp_analyzer.label — 라벨 정규화 + 매칭 + 표 타입 분류.

함수:
- _normalize       : 공백 제거 (비교용)
- _canonical_label : 라벨 → 표준명 (별칭 사전 경유)
- _match_label     : 셀 텍스트 vs 검색 라벨 매칭 (exact / alias / partial)
- classify_table_type : 표 유형 자동 분류 (v0.7.8)
"""
import re

from ._constants import _ALIAS_LOOKUP, _TABLE_TYPE_KEYWORDS


# ── 공백 정규화 ──
def _normalize(text):
    """모든 공백(스페이스, 탭, NBSP 등)을 제거하여 비교용 문자열 반환."""
    return re.sub(r"\s+", "", text)


def _canonical_label(label):
    """라벨을 정규화하고 표준명으로 변환. 별칭 없으면 정규화된 원본 반환."""
    norm = _normalize(label)
    return _ALIAS_LOOKUP.get(norm.upper(), _ALIAS_LOOKUP.get(norm, norm))


def _match_label(cell_text, search_label):
    """셀 텍스트와 검색 라벨이 같은 의미인지 판단.

    Returns: (is_match, is_exact, ratio)
      - is_match: 매칭 여부
      - is_exact: exact match 여부 (정규화 후 완전 일치)
      - ratio: 매칭률 (0.0~1.0, exact이면 1.0)
    """
    norm_cell = _normalize(cell_text)
    norm_label = _normalize(search_label)

    if not norm_cell or not norm_label:
        return False, False, 0.0

    # 1) 정규화 후 exact match (공백만 달랐던 경우)
    if norm_cell == norm_label:
        return True, True, 1.0

    # 2) 별칭 매칭: 둘 다 같은 표준명으로 매핑되는지
    canon_cell = _canonical_label(cell_text)
    canon_label = _canonical_label(search_label)
    if canon_cell == canon_label:
        return True, True, 1.0

    # 3) 정규화된 문자열 포함 관계 (partial match)
    if norm_label in norm_cell:
        return True, False, len(norm_label) / len(norm_cell)
    if norm_cell in norm_label:
        return True, False, len(norm_cell) / len(norm_label)

    return False, False, 0.0


# ---------------------------------------------------------------------------
# v0.7.8 — 표 유형 자동 분류 (2A)
# ---------------------------------------------------------------------------

def classify_table_type(table_info):
    """표 유형 자동 분류 (v0.7.8).

    Returns: 'info_form' | 'financial' | 'timeline' | 'checklist' |
             'comparison' | 'guide' | 'data_table'
    """
    headers = table_info.get("headers", []) or []
    data = table_info.get("data", []) or []
    rows = table_info.get("rows", 0)
    cols = table_info.get("cols", 0)

    # 모든 텍스트 합쳐서 키워드 스캔
    all_text = " ".join(str(h or "") for h in headers)
    if data:
        for row in data[:3]:  # 상위 3행만
            all_text += " " + " ".join(str(c or "") for c in row)

    # 키워드 점수 계산
    scores = {}
    for ttype, keywords in _TABLE_TYPE_KEYWORDS.items():
        score = sum(1 for kw in keywords if kw in all_text)
        if score > 0:
            scores[ttype] = score

    # guide 는 최우선 (작성요령 표)
    if scores.get("guide", 0) > 0:
        return "guide"

    # 최고 점수 유형
    if scores:
        best = max(scores, key=scores.get)
        if scores[best] >= 2:
            return best

    # 레이블-값 쌍 패턴 감지 → info_form
    if cols >= 2 and rows >= 3:
        first_col_texts = [str(headers[0] or "")] if headers else []
        for row in data[:5]:
            if row:
                first_col_texts.append(str(row[0] or ""))
        label_like = sum(
            1 for t in first_col_texts
            if 2 <= len(t.strip()) <= 15 and not t.strip().isdigit()
        )
        if label_like >= len(first_col_texts) * 0.6:
            return "info_form"

    return "data_table"

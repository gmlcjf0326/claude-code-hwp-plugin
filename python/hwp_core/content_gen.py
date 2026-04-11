"""HWP Core — Content Generation Context Builder (v0.7.8).

AI(Claude host)가 각 섹션 본문을 생성할 때 필요한 구조화된 컨텍스트 제공.
이 모듈은 콘텐츠를 직접 생성하지 않음 — LLM 호스트에 최적화된 입력을 구성.

기능:
- map_reference_to_sections: 참고자료→섹션 매핑
- build_section_context: 섹션별 AI 컨텍스트 빌더
"""
import re
import sys
from . import register
from ._helpers import validate_params, normalize_for_match


# ---------------------------------------------------------------------------
# 섹션 주제 alias (v0.7.8 — 2B)
# ---------------------------------------------------------------------------

_SECTION_ALIASES = {
    "사업개요": ["사업내용", "사업목적", "과제개요", "개요", "요약", "사업 개요"],
    "추진배경": ["추진 배경", "배경", "필요성", "추진배경 및 필요성"],
    "시장분석": ["시장현황", "시장규모", "시장전망", "산업분석", "시장환경",
                "시장현황 및 사업화 전망", "국내외 시장현황"],
    "기술현황": ["기술개발", "기술동향", "핵심기술", "R&D", "기술경쟁력",
                "대상기술", "실증대상기술", "사업화 실증대상기술"],
    "계획제품": ["제품개요", "제품", "목표제품", "계획제품 개요", "경쟁제품"],
    "추진계획": ["사업계획", "추진일정", "실행계획", "추진전략", "세부계획",
                "추진체계", "실증참여체계", "실증 추진계획"],
    "자금계획": ["자금조달", "소요예산", "투자계획", "재무계획", "예산편성",
                "연구개발비", "사업비"],
    "인력계획": ["인력현황", "조직", "참여인력", "연구인력", "사업화조직"],
    "기대효과": ["파급효과", "성과목표", "경제적효과", "기대성과", "기대결과",
                "활용방안", "연구개발성과"],
    "지식재산": ["특허", "지식재산권", "특허현황", "내재화", "IP"],
    "일정": ["추진일정", "마일스톤", "사업화 추진일정", "실증 추진일정"],
    "안전보안": ["안전조치", "보안조치", "안전관리", "보안관리", "연구기자재"],
}

# 섹션 유형 분류 (heading 키워드 → type)
_SECTION_TYPE_KEYWORDS = {
    "overview": ["사업 개요", "과제 개요", "사업개요"],
    "background": ["추진배경", "배경", "필요성"],
    "market_analysis": ["시장현황", "시장규모", "시장전망", "시장분석"],
    "technology": ["기술", "대상기술", "핵심경쟁", "기술동향"],
    "product": ["계획제품", "제품개요", "경쟁제품"],
    "strategy": ["추진전략", "추진체계", "추진방안", "실증참여"],
    "timeline": ["추진일정", "일정", "마일스톤"],
    "budget": ["자금", "예산", "투자계획", "매출", "재무"],
    "team": ["인력", "조직", "참여연구원"],
    "expected_outcomes": ["기대효과", "파급효과", "성과목표", "활용방안"],
    "ip": ["특허", "지식재산", "내재화"],
    "safety": ["안전", "보안", "기자재"],
}


def _classify_section_type(heading_text):
    """제목 텍스트로 섹션 유형 분류."""
    norm = normalize_for_match(heading_text)
    for stype, keywords in _SECTION_TYPE_KEYWORDS.items():
        for kw in keywords:
            if kw in norm:
                return stype
    return "general"


def _match_ref_to_section(ref_text_chunk, section_heading):
    """참고자료 텍스트 청크가 특정 섹션과 관련 있는지 점수 매김."""
    norm_ref = normalize_for_match(ref_text_chunk)
    norm_heading = normalize_for_match(section_heading)
    score = 0

    # 직접 매칭: 제목 핵심어가 참고자료에 있으면
    core = re.sub(r'^[\d\s().가-힣]+[\s.)]+', '', norm_heading)
    core_words = [w for w in core.split() if len(w) >= 2]
    for w in core_words:
        if w in norm_ref:
            score += 3

    # alias 매칭
    for topic, aliases in _SECTION_ALIASES.items():
        topic_in_heading = topic in norm_heading or any(a in norm_heading for a in aliases)
        topic_in_ref = topic in norm_ref or any(a in norm_ref for a in aliases)
        if topic_in_heading and topic_in_ref:
            score += 5

    return score


@register("map_reference_to_sections")
def map_reference_to_sections(hwp, params):
    """참고자료 데이터를 양식 섹션에 매핑 (v0.7.8).

    Input:
      - reference_data: {text, tables[]} — ref_reader 출력
      - template_sections: [{heading, level}] — extract_template_structure 출력
      - guide_constraints: [{heading, constraints}] — extract_guide_text 출력

    Output:
      - section_mappings: [{heading, relevant_data, relevant_tables, guide}]
    """
    ref_data = params.get("reference_data", {}) or {}
    sections = params.get("template_sections", []) or []
    guides = params.get("guide_constraints", []) or []

    ref_text = str(ref_data.get("full_text", "") or ref_data.get("text", "") or "")
    ref_tables = ref_data.get("tables", []) or []

    # 참고자료 텍스트를 단락 단위로 분할 (빈 줄 기준)
    ref_chunks = [c.strip() for c in re.split(r'\n\s*\n', ref_text) if c.strip()]

    # guide_constraints 를 heading 기준으로 dict화
    guide_map = {}
    for g in guides:
        h = normalize_for_match(g.get("heading", ""))
        if h:
            guide_map[h] = g

    mappings = []
    for sec in sections:
        heading = sec.get("heading", "") or sec.get("title", "")
        if not heading:
            continue

        section_type = _classify_section_type(heading)

        # 참고자료 텍스트 매칭
        scored_chunks = []
        for chunk in ref_chunks:
            score = _match_ref_to_section(chunk, heading)
            if score > 0:
                scored_chunks.append((score, chunk))
        scored_chunks.sort(key=lambda x: -x[0])
        relevant_text = "\n\n".join(c for _, c in scored_chunks[:3])  # 상위 3개

        # 참고자료 표 매칭 (표 유형 vs 섹션 유형)
        relevant_tables = []
        for rt in ref_tables:
            rt_type = rt.get("table_type", "data_table")
            if _table_type_matches_section(rt_type, section_type):
                relevant_tables.append({
                    "index": rt.get("index", 0),
                    "headers": rt.get("headers", []),
                    "rows": rt.get("rows", 0),
                    "table_type": rt_type,
                })

        # guide 매칭
        norm_heading = normalize_for_match(heading)
        guide = None
        for gh, gdata in guide_map.items():
            if gh in norm_heading or norm_heading in gh:
                guide = gdata.get("constraints", {})
                break

        mappings.append({
            "heading": heading,
            "section_type": section_type,
            "relevant_data": relevant_text[:2000] if relevant_text else "",
            "relevant_tables": relevant_tables,
            "guide_constraints": guide or {},
            "match_score": scored_chunks[0][0] if scored_chunks else 0,
        })

    return {
        "status": "ok",
        "section_mappings": mappings,
        "total_sections": len(mappings),
        "sections_with_data": sum(1 for m in mappings if m["relevant_data"]),
    }


def _table_type_matches_section(table_type, section_type):
    """표 유형이 섹션 유형과 관련 있는지."""
    mapping = {
        "financial": ["budget", "expected_outcomes"],
        "timeline": ["timeline", "strategy"],
        "checklist": ["safety"],
        "comparison": ["market_analysis", "technology", "product"],
        "info_form": ["overview", "background"],
    }
    return section_type in mapping.get(table_type, [])


# ---------------------------------------------------------------------------
# v0.7.8 — 섹션 컨텍스트 빌더 (2C)
# ---------------------------------------------------------------------------

# 섹션별 권장 구조 템플릿
_SECTION_STRUCTURE_TEMPLATES = {
    "overview": "사업 목적 → 핵심 내용 요약 → 기대 성과",
    "background": "현황 분석 → 문제점 제시 → 지원 필요성 → 기대 파급효과",
    "market_analysis": "시장 규모(국내/해외) → 성장률/트렌드 → 경쟁 구도 → 진입 기회",
    "technology": "기술 개요 → 핵심 우수성 → 기술 동향(국내/해외) → 차별화 포인트",
    "product": "제품 개요(그림/사진) → 주요 기능/사양 → 경쟁제품 비교",
    "strategy": "추진 체계 → 단계별 계획 → 역할 분담 → 마일스톤",
    "timeline": "단계 구분 → 월별 세부 일정 → 주요 산출물",
    "budget": "총 사업비 → 항목별 내역 → 연차별 계획 → 자금 조달",
    "team": "조직도 → 핵심 인력 → 역할 및 기여도",
    "expected_outcomes": "경제적 효과(매출/수출) → 기술적 효과 → 사회적 효과 → 인프라",
    "ip": "기존 특허 현황 → 내재화 전략 → 출원 계획 → 파급효과",
    "safety": "안전 책임자 → 교육 계획 → 안전관리비 → 사고 대응",
    "general": "서론 → 현황 분석 → 계획/방안 → 기대 효과",
}

CHARS_PER_PAGE = 1100  # 사업계획서 기준 (A4, 휴먼명조 12pt, 줄간 160%)


@register("build_section_context")
def build_section_context(hwp, params):
    """AI 호스트를 위한 섹션별 구조화 컨텍스트 생성 (v0.7.8).

    Input:
      - section_mappings: map_reference_to_sections 출력
      - template_style: snapshot_template_style 출력 (optional)

    Output:
      - contexts[]: 섹션별 {heading, section_type, guide_constraints,
                    reference_data, suggested_length, structure_template, ...}
    """
    mappings = params.get("section_mappings", []) or []
    style = params.get("template_style", {}) or {}

    contexts = []
    prev_summary = ""

    for i, m in enumerate(mappings):
        heading = m.get("heading", "")
        section_type = m.get("section_type", "general")
        guide = m.get("guide_constraints", {})

        # 권장 글자수 계산
        max_pages = guide.get("max_pages")
        suggested_chars = max_pages * CHARS_PER_PAGE if max_pages else CHARS_PER_PAGE

        # 권장 구조
        structure = _SECTION_STRUCTURE_TEMPLATES.get(section_type,
                        _SECTION_STRUCTURE_TEMPLATES["general"])

        ctx = {
            "index": i,
            "heading": heading,
            "section_type": section_type,
            "guide_constraints": guide,
            "reference_data_relevant": m.get("relevant_data", ""),
            "relevant_tables": m.get("relevant_tables", []),
            "preceding_section_summary": prev_summary[:200] if prev_summary else "",
            "suggested_length_chars": suggested_chars,
            "suggested_length_pages": max_pages or 1,
            "structure_template": structure,
            "format_hints": guide.get("format_hints", []),
        }
        contexts.append(ctx)

        # 다음 섹션을 위한 이전 섹션 요약
        prev_summary = f"[{heading}] 섹션 — {section_type}"

    return {
        "status": "ok",
        "contexts": contexts,
        "total": len(contexts),
        "sections_with_reference": sum(
            1 for c in contexts if c["reference_data_relevant"]
        ),
        "total_suggested_chars": sum(c["suggested_length_chars"] for c in contexts),
    }

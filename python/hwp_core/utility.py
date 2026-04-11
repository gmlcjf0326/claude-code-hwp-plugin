"""HWP Core — Utility handlers.

hwp_service.py 에서 이관된 단순 유틸리티 메서드.
v0.7.6.0 P1-3 첫 이관 — pure read 함수만 먼저 (위험 최소화).

이 모듈은 pyhwpx 상태를 변경하지 않는 read-only 또는 lightweight 메서드만 포함합니다.
"""
from . import register


@register("ping")
def ping(hwp, params):
    """Health check. 서버 살아있는지 확인."""
    return {"status": "ok", "message": "HWP Service is running (via hwp_core.utility)"}


@register("get_font_list")
def get_font_list(hwp, params):
    """폰트 목록 조회. presets.KOREAN_FONTS 기반.

    params:
        category: "serif" | "sans" | "sans_bold" | "handwriting" | ... (선택)
        gov_only: True 면 정부 표준 폰트만 (선택)
    """
    from presets import get_font_list as _get_font_list
    category = params.get("category")
    gov_only = params.get("gov_only", False)
    fonts = _get_font_list(category=category, gov_only=gov_only)
    return {"status": "ok", "fonts": fonts, "count": len(fonts)}


@register("get_preset_list")
def get_preset_list(hwp, params):
    """프리셋 목록 조회 — DOCUMENT_PRESETS + TABLE_STYLES + KOREAN_BUSINESS_DEFAULTS."""
    from presets import DOCUMENT_PRESETS, TABLE_STYLES
    doc_presets = [
        {"name": k, "page": v.get("page", {})}
        for k, v in DOCUMENT_PRESETS.items()
    ]
    table_styles = [
        {"name": k, "header_bg": v.get("header_bg")}
        for k, v in TABLE_STYLES.items()
    ]
    # v0.7.5.4: KOREAN_BUSINESS_DEFAULTS 도 함께 노출
    try:
        from presets import KOREAN_BUSINESS_DEFAULTS
        korean_business = [
            {"name": k, "description": v.get("description", "")}
            for k, v in KOREAN_BUSINESS_DEFAULTS.items()
        ]
    except ImportError:
        korean_business = []
    return {
        "status": "ok",
        "document_presets": doc_presets,
        "table_styles": table_styles,
        "korean_business_defaults": korean_business,
    }

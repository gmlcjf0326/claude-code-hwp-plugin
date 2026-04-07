"""v0.7.2.13: HWPX 본문 단락 서식 XML 직접 읽기.

목적: hwp.HAction.GetDefault("ParaShape", ...) 이 현재 cursor state 가 아닌
시스템 default 만 반환하는 pyhwpx/COM API 한계를 우회.

.hwpx 파일의 section0.xml 첫 본문 단락에서 paraPrIDRef + charPrIDRef 를 찾고,
header.xml 의 paraPr/charPr 정의를 직접 파싱한다.

기존 analyze_writing_patterns COM 경로는 .hwp 용 fallback 으로 유지.
"""
import os
import re
import zipfile
import sys


def read_body_style(file_path: str) -> dict:
    """본문 첫 단락의 paraPr + charPr 정의를 읽어 dict 반환.

    Args:
        file_path: .hwpx 파일의 절대 경로

    Returns:
        {
            "char": {font_size, bold, italic, color, ...},
            "para": {align, line_spacing, left_margin, ...},
            "paraPrIDRef": "3",
            "charPrIDRef": "0",
        }
    """
    if not file_path.lower().endswith(".hwpx"):
        raise ValueError("hwpx_reader: .hwpx 파일만 지원")

    with zipfile.ZipFile(file_path) as z:
        with z.open("Contents/section0.xml") as f:
            section = f.read().decode("utf-8")
        with z.open("Contents/header.xml") as f:
            header = f.read().decode("utf-8")

    # 본문 첫 단락 = 첫 번째 hp:p (제목일 수도, secPr 이 박혀 있을 수도)
    p_match = re.search(
        r'<hp:p\s[^>]*paraPrIDRef="(\d+)"[^>]*>(.*?)</hp:p>',
        section,
        re.DOTALL,
    )
    if not p_match:
        return {"char": {}, "para": {}, "paraPrIDRef": None, "charPrIDRef": None}

    para_id = p_match.group(1)
    p_body = p_match.group(2)

    # 첫 hp:run 의 charPrIDRef
    run_match = re.search(r'<hp:run\s[^>]*charPrIDRef="(\d+)"', p_body)
    char_id = run_match.group(1) if run_match else "0"

    para_def = _find_element_def(header, "paraPr", para_id)
    char_def = _find_element_def(header, "charPr", char_id)

    return {
        "char": _parse_char_pr(char_def),
        "para": _parse_para_pr(para_def),
        "paraPrIDRef": para_id,
        "charPrIDRef": char_id,
    }


def _find_element_def(header: str, tag: str, id_: str) -> str:
    """header.xml 에서 <hh:{tag} id="{id_}">...</hh:{tag}> 또는 self-closing 블록."""
    sc = re.search(rf'<hh:{tag}\s[^>]*id="{id_}"[^>]*/>', header)
    if sc:
        return sc.group(0)
    m = re.search(
        rf'<hh:{tag}\s[^>]*id="{id_}"[^>]*>.*?</hh:{tag}>',
        header,
        re.DOTALL,
    )
    return m.group(0) if m else ""


def _parse_para_pr(xml: str) -> dict:
    """paraPr XML → align/line_spacing/margin 등 dict."""
    if not xml:
        return {}
    result = {}

    # align
    m = re.search(r'<hh:align\s[^/]*horizontal="(\w+)"', xml)
    if m:
        result["align"] = m.group(1).lower()

    # lineSpacing (hp:default 안쪽 기준)
    m = re.search(r'<hp:default>.*?<hh:lineSpacing[^/]*value="(\d+)"', xml, re.DOTALL)
    if not m:
        m = re.search(r'<hh:lineSpacing[^/]*value="(\d+)"', xml)
    if m:
        result["line_spacing"] = int(m.group(1))

    m = re.search(r'<hh:lineSpacing[^/]*type="(\w+)"', xml)
    if m:
        result["line_spacing_type"] = m.group(1).lower()

    # margin (hp:default 우선, 없으면 hp:case 에서)
    default_block = re.search(r'<hp:default>(.*?)</hp:default>', xml, re.DOTALL)
    target = default_block.group(1) if default_block else xml

    for key, pattern in [
        ("indent", r'<hc:intent\s[^/]*value="(-?\d+)"'),
        ("left_margin", r'<hc:left\s[^/]*value="(-?\d+)"'),
        ("right_margin", r'<hc:right\s[^/]*value="(-?\d+)"'),
        ("space_before", r'<hc:prev\s[^/]*value="(-?\d+)"'),
        ("space_after", r'<hc:next\s[^/]*value="(-?\d+)"'),
    ]:
        m = re.search(pattern, target)
        if m:
            result[key] = int(m.group(1))

    # keep_with_next, keep_lines_together, widow_orphan, page_break_before
    m = re.search(r'<hh:breakSetting\s[^/]*/?>', xml)
    if m:
        bs = m.group(0)
        for key, attr in [
            ("keep_with_next", "keepWithNext"),
            ("keep_lines_together", "keepLines"),
            ("widow_orphan", "widowOrphan"),
            ("page_break_before", "pageBreakBefore"),
        ]:
            am = re.search(rf'{attr}="(\d)"', bs)
            if am:
                result[key] = am.group(1) == "1"

    return result


def _parse_char_pr(xml: str) -> dict:
    """charPr XML → font_size/bold/italic/color 등 dict."""
    if not xml:
        return {}
    result = {}

    # height → font_size (height=1000 = 10pt)
    m = re.search(r'\sheight="(\d+)"', xml)
    if m:
        result["font_size"] = int(m.group(1)) / 100

    # textColor
    m = re.search(r'\stextColor="(#[0-9A-Fa-f]+)"', xml)
    if m:
        result["color"] = m.group(1)

    # bold, italic, strikeout 등은 하위 엘리먼트 또는 속성으로 올 수 있음
    # 간단한 버전: bold 속성 + hh:bold 엘리먼트 둘 다 확인
    m = re.search(r'\sbold="(\d)"', xml)
    if m:
        result["bold"] = m.group(1) == "1"
    elif re.search(r'<hh:bold\s', xml):
        result["bold"] = True

    m = re.search(r'\sitalic="(\d)"', xml)
    if m:
        result["italic"] = m.group(1) == "1"
    elif re.search(r'<hh:italic\s', xml):
        result["italic"] = True

    # underline type
    m = re.search(r'<hh:underline[^/]*type="(\w+)"', xml)
    if m:
        t = m.group(1).upper()
        result["underline"] = 0 if t == "NONE" else 1

    return result

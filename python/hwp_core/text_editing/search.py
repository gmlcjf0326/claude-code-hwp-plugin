"""hwp_core.text_editing.search — 검색/치환 handlers.

Handlers:
- text_search        : 텍스트 검색 (full_text 직접 + COM FindReplace fallback)
- find_replace       : 단건 치환
- find_replace_multi : 일괄 치환
- find_replace_nth   : N번째 match 만 치환 (마커 기법)
"""
from .. import register  # 두 점!
from .._helpers import validate_params, _execute_all_replace  # 두 점!


@register("text_search")
def text_search(hwp, params):
    """텍스트 검색 (full_text 직접 + COM FindReplace fallback)."""
    validate_params(params, ["search"], "text_search")
    search_text = params["search"]
    max_results = min(max(params.get("max_results", 50), 1), 1000)

    # 방법 1: 전체 텍스트 직접 검색 (COM FindReplace 반환값 불신뢰 대안)
    full_text = ""
    try:
        full_text = hwp.get_text_file("TEXT", "")
    except Exception:
        pass

    if full_text and search_text in full_text:
        results = []
        pos = 0
        idx = 0
        while idx < max_results:
            found = full_text.find(search_text, pos)
            if found == -1:
                break
            start = max(0, found - 20)
            end = min(len(full_text), found + len(search_text) + 20)
            results.append({
                "index": idx + 1,
                "matched_text": search_text,
                "context": full_text[start:end],
            })
            pos = found + len(search_text)
            idx += 1
        return {
            "search": search_text,
            "total_found": len(results),
            "results": results,
        }

    # 방법 2: COM FindReplace fallback
    hwp.MovePos(2)
    results = []
    for i in range(max_results):
        act = hwp.HAction
        pset = hwp.HParameterSet.HFindReplace
        act.GetDefault("FindReplace", pset.HSet)
        pset.FindString = search_text
        pset.Direction = 0
        pset.IgnoreMessage = 1
        act.Execute("FindReplace", pset.HSet)
        context = ""
        try:
            context = hwp.GetTextFile("TEXT", "saveblock").strip()[:200]
        except Exception:
            pass
        if not context:
            break
        hwp.HAction.Run("Cancel")
        results.append({
            "index": i + 1,
            "matched_text": context[:50] if context else search_text,
        })
    return {
        "search": search_text,
        "total_found": len(results),
        "results": results,
    }


@register("find_replace")
def find_replace(hwp, params):
    """단건 치환."""
    validate_params(params, ["find", "replace"], "find_replace")
    use_regex = params.get("use_regex", False)
    case_sensitive = params.get("case_sensitive", True)
    replaced = _execute_all_replace(hwp, params["find"], params["replace"], use_regex, case_sensitive)
    return {"status": "ok", "find": params["find"], "replace": params["replace"], "replaced": replaced}


@register("find_replace_multi")
def find_replace_multi(hwp, params):
    """일괄 치환."""
    validate_params(params, ["replacements"], "find_replace_multi")
    use_regex = params.get("use_regex", False)
    results = []
    hwp.MovePos(2)
    for item in params["replacements"]:
        replaced = _execute_all_replace(hwp, item["find"], item["replace"], use_regex)
        results.append({"find": item["find"], "replaced": replaced})
    return {
        "status": "ok",
        "results": results,
        "total": len(results),
        "success": sum(1 for r in results if r["replaced"]),
    }


@register("find_replace_nth")
def find_replace_nth(hwp, params):
    """N번째 match 만 치환."""
    validate_params(params, ["find", "replace", "nth"], "find_replace_nth")
    find_text = params["find"]
    replace_text = params["replace"]
    nth = params["nth"]  # 1-based
    if nth < 1 or nth > 10000:
        raise ValueError("nth must be between 1 and 10000")

    before = ""
    try:
        before = hwp.get_text_file("TEXT", "")
    except Exception:
        pass
    count = before.count(find_text)
    if count < nth:
        return {"status": "not_found", "find": find_text, "searched": count, "nth": nth}

    # AllReplace 기반 n번째 치환: 마커 치환 → n번째만 replace → 복원
    import uuid
    marker = f"@@NTH{uuid.uuid4().hex[:6]}@@"
    _execute_all_replace(hwp, find_text, marker)

    hwp.MovePos(2)
    found_count = 0
    for i in range(count):
        act = hwp.HAction
        pset = hwp.HParameterSet.HFindReplace
        act.GetDefault("FindReplace", pset.HSet)
        pset.FindString = marker
        pset.ReplaceString = replace_text if (i == nth - 1) else find_text
        pset.Direction = 0
        pset.IgnoreMessage = 1
        pset.ReplaceMode = 1
        act.Execute("FindReplace", pset.HSet)
        act.GetDefault("FindReplace", pset.HSet)
        pset.FindString = marker
        pset.Direction = 0
        pset.IgnoreMessage = 1
        act.Execute("FindReplace", pset.HSet)
        found_count += 1
    _execute_all_replace(hwp, marker, find_text)

    after = ""
    try:
        after = hwp.get_text_file("TEXT", "")
    except Exception:
        pass
    replaced = replace_text in after
    return {
        "status": "ok" if replaced else "uncertain",
        "find": find_text,
        "replace": replace_text,
        "nth": nth,
        "replaced": replaced,
    }

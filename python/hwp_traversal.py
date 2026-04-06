"""HWP HeadCtrl Traversal Module (v0.6.6 B2).

ctrl = hwp.HeadCtrl 부터 ctrl.Next 로 모든 컨트롤을 순회.
표/그림/머리말/꼬리말/각주/미주/누름틀/하이퍼링크/책갈피/수식 등을
정확히 식별하고 위치 정보를 수집.

References:
- guide-06-hwp-오브젝트내부-이벤트-CtrlID-최종보강.md §1.1~1.9
- hwp-control-HwpCtrl-API.md §컨트롤 트래버스
- 한컴-개발자-API-전체-활용가이드.md §3.1

Safety:
- visited < MAX_VISITS 상한 (무한 루프 방지)
- try/finally로 ctrl 접근 안전 보장
- ctrl 속성 접근 시 모두 try/except (한글 버전 차이 흡수)
"""
import sys

try:
    from hwp_constants import CTRL_ID, CTRL_FILTER_DEFAULT, CTRL_FILTER_ALL
except ImportError:
    # 폴백: hwp_constants 없이도 동작
    CTRL_FILTER_DEFAULT = {"tbl", "gso", "head", "foot", "fn", "en"}
    CTRL_FILTER_ALL = None  # None = 모두


MAX_VISITS = 5000  # 무한 루프 방지 상한


def _safe_get(ctrl, attr, default=None):
    """ctrl 객체의 속성을 안전하게 가져옴 (None/예외 모두 흡수)."""
    try:
        v = getattr(ctrl, attr, default)
        return v if v is not None else default
    except Exception:
        return default


def _param_get(params, name, default=None):
    """ParameterSet에서 값 읽기 (Item 방식 우선, attribute fallback).

    pyhwpx docstring: `prop.SetItem("Rows", 3)` (core.py line 1293).
    읽기는 `Item("Rows")`가 표준. 일부 ParameterSet은 direct attribute도 지원.
    """
    if params is None:
        return default
    # 1. ParameterSet.Item(name) 방식 (표준)
    try:
        v = params.Item(name)
        if v is not None:
            return v
    except Exception:
        pass
    # 2. Direct attribute 방식 (pyhwpx wrapper 지원 시)
    try:
        v = getattr(params, name, None)
        if v is not None:
            return v
    except Exception:
        pass
    return default


def _ctrl_to_dict(ctrl, idx, hwp=None, include_pos=False):
    """단일 ctrl 객체를 직렬화 가능한 dict로 변환.

    수집 정보:
        - ctrl_id: CtrlID 문자열 ("tbl", "gso", "head" 등)
        - user_desc: 사용자 설명 (있으면)
        - has_list: 리스트 보유 여부 (ctrl.HasList property)
        - index: 트래버스 순서 (0-based)
        - table: (표 전용) 행/열 정보 — ParameterSet.Item("Rows"/"Cols") 방식
    """
    cid = _safe_get(ctrl, "CtrlID", "")
    if cid is None:
        cid = ""
    cid = str(cid).strip()

    # HasList: pyhwpx CtrlCode property (core.py line 591-592: return self._com_obj.HasList)
    has_list = False
    try:
        has_list = bool(ctrl.HasList)
    except Exception:
        pass

    info = {
        "index": idx,
        "ctrl_id": cid,
        "user_desc": str(_safe_get(ctrl, "UserDesc", "") or ""),
        "has_list": has_list,
    }

    # 표(tbl) 메타: ctrl.Properties는 ParameterSet (pyhwpx docstring core.py line 1291-1295)
    # 읽기는 Item("Rows"/"Cols") 방식. v0.6.7: _param_get로 defensive 접근.
    if cid == "tbl":
        try:
            props = ctrl.Properties
            if props is not None:
                rows_raw = _param_get(props, "Rows", 0)
                cols_raw = _param_get(props, "Cols", 0)
                try:
                    rows = int(rows_raw or 0)
                except (ValueError, TypeError):
                    rows = 0
                try:
                    cols = int(cols_raw or 0)
                except (ValueError, TypeError):
                    cols = 0
                info["table"] = {"rows": rows, "cols": cols}
        except Exception:
            pass

    return info


def traverse_all_ctrls(hwp, include_ids=None, max_visits=MAX_VISITS):
    """HeadCtrl부터 ctrl.Next를 따라 모든 컨트롤 순회.

    Args:
        hwp: pyhwpx Hwp 인스턴스
        include_ids: 필터링할 ctrl_id 집합 (None = 기본 필터, "all" = 전체)
        max_visits: 무한 루프 방지 상한

    Returns:
        {
            "controls": [...],          # 필터링된 컨트롤 dict 리스트
            "total_visited": int,       # 실제 방문한 ctrl 개수
            "truncated": bool,          # max_visits 초과 여부
            "by_type": {ctrl_id: count} # 타입별 카운트
        }

    Notes:
        - 일부 한글 버전에서 첫 ctrl이 항상 'head'/'foot'으로 시작 가능
        - HasList=True인 ctrl은 내부에 추가 ctrl 보유 (예: 표 안의 그림)
        - 본 함수는 최상위만 순회 (재귀 X)
    """
    if include_ids == "all":
        include_set = None  # 모두
    elif include_ids is None:
        include_set = CTRL_FILTER_DEFAULT
    elif isinstance(include_ids, (list, tuple, set)):
        include_set = set(include_ids)
    else:
        include_set = CTRL_FILTER_DEFAULT

    controls = []
    by_type = {}
    visited = 0
    truncated = False

    try:
        ctrl = _safe_get(hwp, "HeadCtrl", None)
    except Exception as e:
        print(f"[WARN] HeadCtrl access failed: {e}", file=sys.stderr)
        return {
            "controls": [],
            "total_visited": 0,
            "truncated": False,
            "by_type": {},
            "error": f"HeadCtrl access failed: {e}",
        }

    while ctrl is not None:
        if visited >= max_visits:
            truncated = True
            print(f"[WARN] traverse_all_ctrls: max_visits {max_visits} reached",
                  file=sys.stderr)
            break

        try:
            cid = str(_safe_get(ctrl, "CtrlID", "") or "").strip()
            # 카운트 (필터와 무관하게 전체)
            by_type[cid] = by_type.get(cid, 0) + 1

            # 필터 적용
            if include_set is None or cid in include_set:
                info = _ctrl_to_dict(ctrl, visited, hwp=hwp)
                controls.append(info)
        except Exception as e:
            print(f"[WARN] ctrl info extract failed at idx {visited}: {e}",
                  file=sys.stderr)

        visited += 1

        # 다음 ctrl로 이동
        try:
            next_ctrl = _safe_get(ctrl, "Next", None)
            if next_ctrl is None:
                break
            ctrl = next_ctrl
        except Exception as e:
            print(f"[WARN] ctrl.Next failed at idx {visited}: {e}", file=sys.stderr)
            break

    return {
        "controls": controls,
        "total_visited": visited,
        "truncated": truncated,
        "by_type": by_type,
    }


def count_ctrls_by_type(hwp, max_visits=MAX_VISITS):
    """전체 컨트롤을 타입별로 카운트만 반환 (디테일 없음, 빠름).

    Returns:
        {ctrl_id: count, ...}
    """
    result = traverse_all_ctrls(hwp, include_ids="all", max_visits=max_visits)
    return result.get("by_type", {})


def find_ctrls_by_id(hwp, target_id, max_visits=MAX_VISITS):
    """특정 ctrl_id를 가진 컨트롤만 반환.

    Args:
        target_id: 찾을 ctrl_id 문자열 (예: "tbl", "gso")

    Returns:
        해당 ctrl_id의 컨트롤 dict 리스트
    """
    result = traverse_all_ctrls(hwp, include_ids={target_id}, max_visits=max_visits)
    return result.get("controls", [])

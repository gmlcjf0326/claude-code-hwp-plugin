"""HWP Constants and Context Managers (v0.6.6+).

한컴 공식 가이드 (docs/원본가이드소스/) 기반 상수 + 컨텍스트 매니저.
모든 신규 모듈은 여기서 import. raw 매직 넘버 금지.

References:
- guide-01-hwp-com-api-python-실전가이드.md (MovePos 상수)
- guide-05-hwp-누락API-고급기능-보강.md (SetMessageBoxMode, InitScan)
- guide-06-hwp-오브젝트내부-이벤트-CtrlID-최종보강.md (CtrlID 매핑)
"""
import sys
from contextlib import contextmanager


# ─────────────────────────────────────────────────────────────────────────
# MovePos 상수 (hwp.MovePos(N))
# 출처: guide-01 §MovePos 상수
# ─────────────────────────────────────────────────────────────────────────
MOVE_POS = {
    "DOC_BEGIN": 2,           # 문서 처음 (커서 초기화 용도)
    "DOC_END_EXIT_TABLE": 3,  # 문서 끝 (표 안에 있어도 무조건 탈출 — 메모리 #2)
    "CELL_BEGIN": 100,        # 현재 셀 처음
    "CELL_END": 101,          # 현재 셀 끝
    "PARA_BEGIN": 102,        # 현재 단락 처음
    "PARA_END": 103,          # 현재 단락 끝
    "LINE_BEGIN": 104,        # 현재 줄 처음
    "LINE_END": 105,          # 현재 줄 끝
    "WORD_BEGIN": 106,        # 현재 단어 처음
    "WORD_END": 107,          # 현재 단어 끝
}


# ─────────────────────────────────────────────────────────────────────────
# CtrlID 매핑 (HeadCtrl 순회 시 ctrl.CtrlID 비교용)
# 출처: guide-06 §1.1~1.9
# ─────────────────────────────────────────────────────────────────────────
CTRL_ID = {
    "TABLE": "tbl",         # 표
    "SHAPE": "gso",         # 그리기 개체 (그림/도형)
    "HEADER": "head",       # 머리말
    "FOOTER": "foot",       # 꼬리말
    "FOOTNOTE": "fn",       # 각주
    "ENDNOTE": "en",        # 미주
    "DATE_TIME": "%dte",    # 현재 날짜/시간
    "DATE_DOC": "%ddt",     # 작성 날짜
    "FIELD_CLICK": "%clk",  # 누름틀 필드
    "HYPERLINK": "%hlk",    # 하이퍼링크
    "EQUATION": "eqed",     # 수식
    "BOOKMARK": "bokm",     # 책갈피
}

# 트래버스 시 자주 쓰는 필터 프리셋
CTRL_FILTER_DEFAULT = {"tbl", "gso", "head", "foot", "fn", "en"}
CTRL_FILTER_FIELDS = {"%clk", "%hlk", "bokm"}
CTRL_FILTER_ALL = set(CTRL_ID.values())


# ─────────────────────────────────────────────────────────────────────────
# SetMessageBoxMode 값 (hwp.SetMessageBoxMode(N))
# 출처: guide-05 §메시지박스 모드
# ─────────────────────────────────────────────────────────────────────────
MSG_BOX_MODE_AUTO_OK = 0x00010000   # 모든 다이얼로그 자동 OK (자동화 안전)
MSG_BOX_MODE_RESTORE = 0xFFFFFFFF   # 원래대로 복원

# v0.6.6 dispatch 진입부에서 멱등 적용 (B1)
MSG_BOX_MODE_DEFAULT = MSG_BOX_MODE_AUTO_OK


# ─────────────────────────────────────────────────────────────────────────
# InitScan flags (hwp.InitScan(option, range))
# 출처: guide-05 §검색 시스템
# ─────────────────────────────────────────────────────────────────────────
SCAN_FLAG_DEFAULT = 0x0077      # 일반 텍스트 + 표 셀 + 컨트롤 텍스트
SCAN_FLAG_TEXT_ONLY = 0x0001    # 일반 텍스트만
SCAN_FLAG_TABLE_ONLY = 0x0010   # 표 셀만


# ─────────────────────────────────────────────────────────────────────────
# Context Managers (try/finally 자동화)
# ─────────────────────────────────────────────────────────────────────────

@contextmanager
def hwp_safe_context(hwp):
    """SetMessageBoxMode 자동 적용 + finally에서 무시 (한글 버전 차이 흡수).

    사용 예:
        with hwp_safe_context(hwp):
            hwp.open(path)
            ...

    멱등 안전. 한글 미설치 / 구버전이면 try/except로 무시.
    """
    try:
        try:
            hwp.SetMessageBoxMode(MSG_BOX_MODE_DEFAULT)
        except Exception as e:
            print(f"[WARN] SetMessageBoxMode failed (older HWP?): {e}",
                  file=sys.stderr)
        yield hwp
    finally:
        # 명시적 복원은 하지 않음 (다음 호출에서 재설정되므로 멱등)
        pass


@contextmanager
def scan_context(hwp, flags=SCAN_FLAG_DEFAULT, scan_range=None):
    """InitScan / GetText / ReleaseScan 자동 보장.

    v0.7.2.10: scan_range default 를 None 으로 변경.
    이전 default=0 (현재 선택 영역) 이 word_count 등에서 본문을 못 읽는 hidden bug 였음.
    이제 None 이면 positional 호출 (hwp_analyzer 와 동일 패턴) → pyhwpx default Range
    사용 (전체 문서). scan_range 명시 시에만 named 호출.

    사용 예:
        with scan_context(hwp) as scan:
            while True:
                state, text = hwp.GetText()
                if state <= 1:
                    break
                ...

    예외 발생 시에도 ReleaseScan() 보장 (메모리 누수 / 한글 멈춤 방지).
    """
    initialized = False
    try:
        try:
            if scan_range is None:
                # v0.7.2.10: hwp_analyzer 와 동일한 positional 호출 (전체 문서 default)
                hwp.InitScan(flags)
            else:
                hwp.InitScan(Range=scan_range, SpecialChar=flags)
            initialized = True
        except Exception:
            # 일부 인자 시그니처가 다른 한글 버전 — positional 2-arg 폴백
            try:
                hwp.InitScan(flags, scan_range if scan_range is not None else 0xff)
                initialized = True
            except Exception as e:
                print(f"[WARN] InitScan failed: {e}", file=sys.stderr)
                raise
        yield hwp
    finally:
        if initialized:
            try:
                hwp.ReleaseScan()
            except Exception as e:
                print(f"[WARN] ReleaseScan failed: {e}", file=sys.stderr)


# ─────────────────────────────────────────────────────────────────────────
# Helper: dispatch 진입부에서 호출용 (B1)
# ─────────────────────────────────────────────────────────────────────────
def apply_safe_mode(hwp):
    """dispatch() 진입부에서 호출. 멱등, 실패 시 무시.

    기존 hwp_service.py:182 `XHwpMessageBoxMode = 1`은 폴백으로 유지.
    """
    try:
        hwp.SetMessageBoxMode(MSG_BOX_MODE_DEFAULT)
    except Exception:
        # 한글 버전 차이 / 호출 가능 시점 등 — 조용히 무시
        pass

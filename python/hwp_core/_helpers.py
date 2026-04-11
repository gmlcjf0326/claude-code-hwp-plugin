"""HWP Core Helpers — 공유 유틸리티 함수.

hwp_service.py 의 모듈화 지원. 여러 handler 가 공통으로 사용하는 저수준 함수 모음.

이 파일은:
- pyhwpx COM 객체 조작을 최소화 (각 handler 에서 직접 사용)
- 순수 유틸리티 (경로 검증, 파라미터 검증, 텍스트 비교)
- 표 안전 탈출, AllReplace 공통 패턴 등
"""
import os
import re
import sys
import unicodedata


# ---------------------------------------------------------------------------
# 텍스트 정규화 (v0.7.7 — 1A)
# ---------------------------------------------------------------------------

# Fullwidth → halfwidth 괄호/기호 맵
_BRACKET_MAP = {
    '（': '(', '）': ')', '＜': '<', '＞': '>',
    '【': '[', '】': ']', '｛': '{', '｝': '}',
    '「': '[', '」': ']', '『': '[', '』': ']',
    '〈': '<', '〉': '>',
}


def normalize_unicode(text: str) -> str:
    """NFKC + fullwidth→halfwidth 괄호 정규화. 비파괴적.

    모든 매칭 함수의 기초 단계. NFKC 는 fullwidth 숫자/영문도 halfwidth 로 변환.
    """
    text = unicodedata.normalize("NFKC", text)
    for k, v in _BRACKET_MAP.items():
        text = text.replace(k, v)
    return text


def normalize_for_match(text: str) -> str:
    """공백 축소 + NFKC + 괄호 정규화. 비교용.

    heading/label 매칭 시 양쪽을 이 함수로 정규화 후 비교.
    """
    return re.sub(r'\s+', ' ', normalize_unicode(text)).strip()


def normalize_for_display(text: str) -> str:
    """공백 정리만. UI 표시용."""
    return re.sub(r'\s+', ' ', text).strip()


def _exit_table_safely(hwp):
    """표에서 안전하게 탈출. MovePos(3)으로 문서 끝(표 밖)으로 이동.

    Cancel() 은 표 탈출에 실패 — v0.5 메모리 기록 참조.
    MovePos(3) 는 문서 마지막 위치로 이동 → 표 밖으로 자연 탈출.
    """
    try:
        if hwp.is_cell():
            hwp.MovePos(3)  # 문서 마지막 위치 (표 밖으로 탈출)
    except Exception:
        pass
    # 표 밖 확인 후 새 문단 생성
    try:
        if not hwp.is_cell():
            hwp.HAction.Run("BreakPara")
    except Exception:
        pass


def validate_file_path(file_path, must_exist=True):
    """경로 보안 검증.

    - 심볼릭 링크 거부 (보안)
    - must_exist=True: 파일 존재 필수
    - must_exist=False: 저장 대상 — 디렉토리 존재 + 쓰기 권한 + 잠금 확인
    - HWP/HWPX 파일은 한글이 잠금 보유 중이므로 잠금 사전 확인 skip

    한글 에러 대화상자를 사전에 방지해서 COM freeze 없이 Python 에러 반환.
    """
    real = os.path.abspath(file_path)
    if os.path.islink(file_path):
        raise ValueError(f"심볼릭 링크는 허용되지 않습니다: {file_path}")
    if must_exist and not os.path.exists(real):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {real}")
    if not must_exist:
        # 저장 대상 경로: 디렉토리 존재 + 쓰기 권한 사전 확인
        dir_path = os.path.dirname(real) or '.'
        if not os.path.exists(dir_path):
            raise FileNotFoundError(f"저장 디렉토리가 존재하지 않습니다: {dir_path}")
        if not os.access(dir_path, os.W_OK):
            raise PermissionError(f"디렉토리 쓰기 권한이 없습니다: {dir_path}")
        # 기존 파일 덮어쓰기: 쓰기 권한 + 잠금 확인
        if os.path.exists(real):
            if not os.access(real, os.W_OK):
                raise PermissionError(f"파일 쓰기 권한이 없습니다 (읽기 전용 또는 잠김): {real}")
            # 파일 잠금 사전 확인 — HWP/HWPX 파일은 제외 (한글이 잠금 보유 중)
            ext = os.path.splitext(real)[1].lower()
            if ext not in ('.hwp', '.hwpx'):
                try:
                    with open(real, 'a'):
                        pass
                except (PermissionError, IOError):
                    raise PermissionError(f"파일이 다른 프로그램에서 사용 중입니다: {real}")
    return real


def _execute_all_replace(hwp, find_str, replace_str, use_regex=False, case_sensitive=True):
    """AllReplace 공통 함수.

    전후 텍스트 비교로 실제 치환 발생 여부 판단 (COM 반환값 신뢰 불가).
    H4: 타임아웃 시 낙관적 가정 (Execute 는 실행됨, 결과만 못 가져옴).

    Returns: bool (치환 발생했는지)
    """
    # 치환 전 텍스트 캡처
    before = None
    try:
        before = hwp.get_text_file("TEXT", "")
    except Exception as e:
        print(f"[WARN] get_text_file before failed: {e}", file=sys.stderr)

    act = hwp.HAction
    pset = hwp.HParameterSet.HFindReplace
    act.GetDefault("AllReplace", pset.HSet)
    pset.FindString = find_str
    pset.ReplaceString = replace_str
    pset.IgnoreMessage = 1
    pset.Direction = 0
    pset.FindRegExp = 1 if use_regex else 0
    pset.FindJaso = 0
    # 대소문자 구분 옵션
    try:
        pset.MatchCase = 1 if case_sensitive else 0
    except Exception:
        pass  # 일부 한글 버전에서 미지원
    pset.AllWordForms = 0
    pset.SeveralWords = 0
    act.Execute("AllReplace", pset.HSet)

    # 치환 후 텍스트 비교로 실제 변경 여부 판단
    after = None
    try:
        after = hwp.get_text_file("TEXT", "")
    except Exception as e:
        print(f"[WARN] get_text_file after failed: {e}", file=sys.stderr)

    # 2C3: 텍스트 캡처 실패 시 구분
    if before is None and after is None:
        # 양쪽 모두 실패 → Execute는 실행됨 → 낙관적 True
        return True
    if before is None or after is None:
        # 한쪽만 실패 → 비교 불가하지만 Execute는 실행됨 → 로깅 후 True
        print(
            f"[WARN] Partial text capture: "
            f"before={'ok' if before is not None else 'fail'}, "
            f"after={'ok' if after is not None else 'fail'}",
            file=sys.stderr,
        )
        return True
    return before != after


def validate_params(params, required_keys, method_name):
    """필수 파라미터 존재 검증. 누락 시 ValueError."""
    missing = [k for k in required_keys if k not in params]
    if missing:
        raise ValueError(f"{method_name}: missing required params: {', '.join(missing)}")

"""hwp_core.text_editing._internal — 비공개 helpers.

v0.7.9 Phase 7: text_editing.py 에서 분리

내용:
- Windows API 대화상자 자동 처리 (v0.7.9 핵심 fix)
  - _find_hwp_confirm_dialog, _auto_dismiss_hwp_dialog, _with_auto_dismiss
- 3-tier fuzzy heading matcher (v0.7.7)
  - _find_heading_positions, _find_all_in_normalized, _extract_match_from_line,
    _find_by_regex, _find_core_in_lines
- 마커/번호 자동 내어쓰기 (v0.7.9)
  - _apply_indent_at_caret, _detect_heading_depth

이 모듈은 @register 없음. search.py / insertions.py 가 import.
"""
import re
import sys
import time
import threading

from .._helpers import normalize_for_match  # 두 점!


# v0.7.9: Windows API 로 HWP 대화상자 자동 처리
# "본문을 [바탕글] 스타일 모양으로 덮어 쓸까요?" — XHwpMessageBoxMode 로 못 잡는
# 특정 confirm 대화상자를 Enter 키 전송으로 "덮어씀(Y)" 자동 클릭.
try:
    import ctypes
    from ctypes import wintypes
    _user32 = ctypes.windll.user32
    _HAS_CTYPES = True
    # EnumWindows callback type
    _WNDENUMPROC = ctypes.WINFUNCTYPE(
        ctypes.c_bool, wintypes.HWND, wintypes.LPARAM
    )
    # 상수 (Win32 messages + button IDs)
    _WM_KEYDOWN = 0x0100
    _WM_KEYUP = 0x0101
    _WM_COMMAND = 0x0111
    _VK_RETURN = 0x0D
    _BM_CLICK = 0x00F5
    _BN_CLICKED = 0
    _IDOK = 1
    _IDYES = 6
    _IDCLOSE = 8
    # 버튼 style + GetWindowLongW index
    _BS_DEFPUSHBUTTON = 0x01
    _BS_PUSHBUTTON = 0x00
    _GWL_STYLE = -16
    # v0.7.9-postfix4 (Tier 4): EnumChildWindows callback for button auto-detect
    _ENUMCHILDPROC = ctypes.WINFUNCTYPE(
        ctypes.c_bool, wintypes.HWND, wintypes.LPARAM
    )
except Exception:
    _HAS_CTYPES = False
    _user32 = None


def _find_hwp_confirm_dialog():
    """현재 열려있는 HWP 확인 대화상자 HWND 찾기.

    v0.7.9-postfix4 (Firefox/Chrome 등 다른 앱 차단):
    - postfix3 가 keywords ("스타일","확인","한컴","알림","덮어","바탕글") 로 매칭했는데
      Firefox/Chrome 페이지 title 에 한국어 키워드 포함 시 잘못 매칭됨 (사용자 보고 + stderr 진단)
    - postfix4: class blocklist 확장 + title 키워드를 매우 narrow phrase 로 좁힘

    Class blocklist (메인 앱 창 차단):
    - HwpApp* (한컴 메인)
    - Mozilla* (Firefox)
    - Chrome_*, ApplicationFrameWindow (Chrome, Edge)
    - CabinetWClass (탐색기), Shell_*
    - Code, Notepad, OpusApp, Word*, Excel*, PowerPoint*
    - WindowsForms* (.NET 앱)

    Title narrow phrase (dialog 의 unique 한 long phrase):
    - "스타일 모양" — "본문을 [바탕글] 스타일 모양으로 덮어 쓸까요?" 의 unique 부분
    - "덮어 쓸까" — confirm dialog 의 unique 부분
    - "모두 허용" — 보안 dialog
    - "선택하신 스타일" — 다른 스타일 dialog

    Returns: HWND or 0
    """
    if not _HAS_CTYPES:
        return 0
    found_hwnd = [0]
    # 키워드 — 원래 keyword 유지 (narrow 하면 한컴 dialog 도 못 잡음)
    # postfix4 의 핵심: class blocklist 로 다른 앱 (Firefox/Chrome 등) 차단 후
    # 키워드는 원래대로 사용
    keywords = ("스타일", "확인", "한컴", "알림", "덮어", "바탕글")
    # 메인 앱 창 class blocklist (확장 — Firefox/Chrome/Edge/탐색기/VSCode 등)
    blocked_class_prefixes = (
        "HwpApp",            # 한컴 메인
        "Mozilla",           # Firefox (사용자 보고 hwnd 197284)
        "Chrome_",           # Chrome / Electron 앱
        "ApplicationFrame",  # Edge UWP
        "Windows.UI",        # UWP (CoreWindow)
        "XamlWindow",        # WinUI
        "CabinetWClass",     # 탐색기
        "Shell_",            # 작업표시줄
        "WindowsForms",      # .NET WinForms
        "Code",              # VSCode
        "OpusApp",           # Word
        "XLMAIN",            # Excel
        "PPTFrameClass",     # PowerPoint
        "Notepad",
        "Progman",           # 데스크톱
        "WorkerW",           # 데스크톱
        "EVERYTHING",        # Everything 검색기
    )

    def _callback(hwnd, lparam):
        try:
            if not _user32.IsWindowVisible(hwnd):
                return True
            class_buf = ctypes.create_unicode_buffer(256)
            _user32.GetClassNameW(hwnd, class_buf, 256)
            class_name = class_buf.value or ""
            for prefix in blocked_class_prefixes:
                if class_name.startswith(prefix):
                    return True  # 메인 앱 창 차단
            length = _user32.GetWindowTextLengthW(hwnd)
            if length == 0:
                return True
            buf = ctypes.create_unicode_buffer(length + 1)
            _user32.GetWindowTextW(hwnd, buf, length + 1)
            title = buf.value or ""
            # title pattern blocklist — 안전 식별자
            if "한컴오피스" in title or ".hwp [" in title or ".hwpx [" in title:
                return True
            for kw in keywords:
                if kw in title:
                    found_hwnd[0] = hwnd
                    return False  # stop enum
        except Exception:
            pass
        return True

    try:
        _user32.EnumWindows(_WNDENUMPROC(_callback), 0)
    except Exception:
        pass
    return found_hwnd[0]


def _find_default_button_in_dialog(parent_hwnd):
    """v0.7.9-postfix2 Tier 4: Dialog 내부의 default button (BS_DEFPUSHBUTTON) 찾기.

    EnumChildWindows 로 dialog 의 모든 child window 순회하면서:
    1. class 가 "Button" 인지
    2. style 에 BS_DEFPUSHBUTTON (0x01) 비트 있는지

    찾으면 해당 button HWND 반환. 없으면 첫 번째 button HWND 반환 (fallback).

    Returns: button HWND or 0
    """
    if not _HAS_CTYPES or not parent_hwnd:
        return 0
    default_button = [0]
    first_button = [0]

    def _callback(hwnd, lparam):
        try:
            # class 확인 — "Button" 만
            cls_buf = ctypes.create_unicode_buffer(64)
            _user32.GetClassNameW(hwnd, cls_buf, 64)
            if cls_buf.value != "Button":
                return True
            # style 확인 — GetWindowLongW(hwnd, GWL_STYLE)
            style = _user32.GetWindowLongW(hwnd, _GWL_STYLE)
            # 첫 번째 button 기록 (fallback)
            if first_button[0] == 0:
                first_button[0] = hwnd
            # BS_DEFPUSHBUTTON 비트 확인 (lower nibble)
            if (style & 0x0F) == _BS_DEFPUSHBUTTON:
                default_button[0] = hwnd
                return False  # 찾으면 stop
        except Exception:
            pass
        return True

    try:
        _user32.EnumChildWindows(parent_hwnd, _ENUMCHILDPROC(_callback), 0)
    except Exception:
        pass
    # default 우선, 없으면 first
    return default_button[0] or first_button[0]


def _dump_all_visible_windows():
    """v0.7.9-postfix3-diag: 모든 visible top-level window 의 class+title 를 stderr 로 dump.

    dialog 진단용 — 어떤 window 들이 EnumWindows 로 보이는지 확인.
    """
    if not _HAS_CTYPES:
        return
    dumped = []

    def _cb(hwnd, lparam):
        try:
            if not _user32.IsWindowVisible(hwnd):
                return True
            cls = ctypes.create_unicode_buffer(256)
            _user32.GetClassNameW(hwnd, cls, 256)
            length = _user32.GetWindowTextLengthW(hwnd)
            title = ""
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                _user32.GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value or ""
            dumped.append(f"  hwnd={hwnd} class={cls.value!r} title={title!r}")
        except Exception:
            pass
        return True

    try:
        _user32.EnumWindows(_WNDENUMPROC(_cb), 0)
    except Exception:
        pass
    print(f"[DIALOG-DIAG] visible windows ({len(dumped)}):", file=sys.stderr)
    for line in dumped[:30]:
        print(line, file=sys.stderr)
    sys.stderr.flush()


def _auto_dismiss_hwp_dialog(max_wait_sec=8):
    """백그라운드 스레드: HWP 확인 대화상자를 4-tier 전략으로 자동 닫기.

    v0.7.9-postfix2 (사용자 보고: "덮어씀", "모두 허용" 여전히 뜸):
    - Tier 1 (postfix3 유지): _find_hwp_confirm_dialog 가 dialog HWND 찾기
    - Tier 2 (~70%): SendMessageW (sync) — PostMessageW 대체
    - Tier 3 (~90%): WM_COMMAND BN_CLICKED + IDYES(6)/IDOK(1) 직접 클릭
    - Tier 4 (~95%): EnumChildWindows + BS_DEFPUSHBUTTON 자동 감지 → BM_CLICK

    v0.7.9-postfix3 (diagnostic): stderr 로 dialog 활동 로깅.
    """
    if not _HAS_CTYPES:
        return
    deadline = time.time() + max_wait_sec
    dismissed = 0
    poll_count = 0
    last_dump = 0
    while time.time() < deadline:
        try:
            hwnd = _find_hwp_confirm_dialog()
            poll_count += 1
            # 매 1초마다 한 번 (20 polls × 50ms) 모든 visible window dump
            if poll_count - last_dump >= 20:
                _dump_all_visible_windows()
                last_dump = poll_count
                print(f"[DIALOG-DIAG] poll {poll_count}: confirm_dialog hwnd={hwnd}", file=sys.stderr)
                sys.stderr.flush()
            if hwnd:
                print(f"[DIALOG-DIAG] FOUND dialog hwnd={hwnd}, applying Tier 2+3+4", file=sys.stderr)
                sys.stderr.flush()
                # === Tier 4 (가장 강력): default button 자동 감지 + BM_CLICK ===
                btn = _find_default_button_in_dialog(hwnd)
                print(f"[DIALOG-DIAG] Tier4 default_button hwnd={btn}", file=sys.stderr)
                if btn:
                    try:
                        _user32.SendMessageW(btn, _BM_CLICK, 0, 0)
                        print(f"[DIALOG-DIAG] Tier4 SendMessage(button, BM_CLICK) sent", file=sys.stderr)
                    except Exception as e:
                        print(f"[DIALOG-DIAG] Tier4 ERR: {e}", file=sys.stderr)

                # === Tier 3: WM_COMMAND IDYES + IDOK 직접 클릭 ===
                for button_id in (_IDYES, _IDOK):
                    try:
                        wParam = (_BN_CLICKED << 16) | button_id
                        _user32.SendMessageW(hwnd, _WM_COMMAND, wParam, 0)
                    except Exception:
                        pass
                print(f"[DIALOG-DIAG] Tier3 WM_COMMAND IDYES/IDOK sent", file=sys.stderr)

                # === Tier 2: SendMessage WM_KEYDOWN VK_RETURN (sync, fallback) ===
                try:
                    _user32.SendMessageW(hwnd, _WM_KEYDOWN, _VK_RETURN, 0)
                    _user32.SendMessageW(hwnd, _WM_KEYUP, _VK_RETURN, 0)
                except Exception:
                    pass
                print(f"[DIALOG-DIAG] Tier2 SendMessage VK_RETURN sent", file=sys.stderr)
                sys.stderr.flush()

                dismissed += 1
                time.sleep(0.2)  # 다음 dialog 대기
                if dismissed >= 5:  # 최대 5개 연속 dialog 처리
                    print(f"[DIALOG-DIAG] dismissed limit reached (5)", file=sys.stderr)
                    return
        except Exception as e:
            print(f"[DIALOG-DIAG] outer loop ERR: {e}", file=sys.stderr)
        time.sleep(0.05)
    print(f"[DIALOG-DIAG] deadline reached, total polls={poll_count}, dismissed={dismissed}", file=sys.stderr)
    sys.stderr.flush()


def _with_auto_dismiss(hwp, callable_fn, *args, **kwargs):
    """함수 호출 중 백그라운드에서 대화상자 자동 처리 + MessageBoxMode 재설정.

    사용: _with_auto_dismiss(hwp, hwp.set_style, "바탕글")
    """
    # 1) XHwpMessageBoxMode + SetMessageBoxMode 재적용 (set_style 이 리셋하는 버그 우회)
    try:
        hwp.XHwpMessageBoxMode = 1
    except Exception:
        pass
    try:
        hwp.SetMessageBoxMode(0x00010000)
    except Exception:
        pass

    # 2) 백그라운드 스레드로 대화상자 감시 시작
    t = threading.Thread(
        target=_auto_dismiss_hwp_dialog,
        args=(5,),
        daemon=True,
    )
    t.start()

    # 3) 실제 함수 호출
    try:
        return callable_fn(*args, **kwargs)
    finally:
        # 스레드는 daemon 이고 자체 타임아웃으로 종료됨
        pass


# ---------------------------------------------------------------------------
# v0.7.7 — 3-tier fuzzy heading matcher
# ---------------------------------------------------------------------------

def _find_heading_positions(full_text, heading):
    """3-tier cascading heading search.

    Returns list of (char_position, matched_text, tier).
    Tier 1: 정규화 후 정확 일치 (공백/괄호 정규화)
    Tier 2: 공백 유연 regex (whitespace-flexible)
    Tier 3: 핵심어 추출 포함 검색 (번호 제거 후 한글 핵심어)

    상위 tier 에서 매칭되면 하위 tier 는 시도하지 않음.
    """
    if not full_text or not heading:
        return []

    # ── Tier 1: 정규화 후 정확 일치 ──
    norm_text = normalize_for_match(full_text)
    norm_heading = normalize_for_match(heading)
    results = _find_all_in_normalized(full_text, norm_text, norm_heading)
    if results:
        return [(pos, txt, 1) for pos, txt in results]

    # ── Tier 2: 공백 유연 regex ──
    # 각 비공백 토큰 사이에 \s* 삽입
    tokens = norm_heading.split()
    if len(tokens) >= 2:
        pattern = r'\s*'.join(re.escape(t) for t in tokens)
        results = _find_by_regex(full_text, pattern)
        if results:
            return [(pos, txt, 2) for pos, txt in results]

    # ── Tier 3: 핵심어 추출 포함 검색 ──
    # 번호 prefix 제거 → 한글 핵심어만 추출
    core = re.sub(r'^[\d\s().·\-\[\]가나다라마바사아자차카타파하]+[\s.)]+', '', norm_heading)
    core = re.sub(r'\s+', '', core)  # 공백 모두 제거
    if len(core) >= 3:
        results = _find_core_in_lines(full_text, core, heading)
        if results:
            return [(pos, txt, 3) for pos, txt in results]

    return []


def _find_all_in_normalized(original, norm_text, norm_heading):
    """줄 단위로 정규화 비교하여 원본 위치 + 매칭 텍스트 반환.

    position mapping 대신 줄 단위 탐색으로 정확도 보장.
    """
    results = []
    lines = original.split('\n')
    char_pos = 0
    for line in lines:
        norm_line = normalize_for_match(line)
        search_pos = 0
        while True:
            idx = norm_line.find(norm_heading, search_pos)
            if idx < 0:
                break
            # 원본 줄에서 매칭 부분 추출
            matched = _extract_match_from_line(line, norm_line, idx, norm_heading)
            results.append((char_pos, matched))
            search_pos = idx + len(norm_heading)
        char_pos += len(line) + 1  # +1 for \n
    return results


def _extract_match_from_line(orig_line, norm_line, norm_start, norm_heading):
    """정규화된 줄의 위치를 원본 줄의 부분문자열로 변환."""
    # 정규화 줄에서 매칭된 부분 앞의 비공백 문자 수 카운트
    prefix_non_space = len(norm_line[:norm_start].replace(' ', ''))
    # 원본 줄에서 같은 수의 비공백 문자 위치 찾기
    orig_start = 0
    count = 0
    for i, c in enumerate(orig_line):
        if c not in ' \t\r':
            if count >= prefix_non_space:
                orig_start = i
                break
            count += 1
    else:
        orig_start = len(orig_line)
    # 매칭 텍스트의 비공백 문자 수
    heading_non_space = len(norm_heading.replace(' ', ''))
    orig_end = orig_start
    count = 0
    for i in range(orig_start, len(orig_line)):
        if orig_line[i] not in ' \t\r':
            count += 1
        if count >= heading_non_space:
            orig_end = i + 1
            break
    else:
        orig_end = len(orig_line)
    return orig_line[orig_start:orig_end].strip()


def _find_by_regex(full_text, pattern):
    """regex 패턴으로 모든 매칭 위치 반환."""
    results = []
    try:
        for m in re.finditer(pattern, full_text, re.IGNORECASE):
            results.append((m.start(), m.group()))
    except re.error:
        pass
    return results


def _find_core_in_lines(full_text, core_text, original_heading):
    """핵심어를 포함하는 줄을 찾되, 번호가 있으면 번호도 확인."""
    results = []
    # 원본 heading 에서 번호 추출
    num_match = re.match(r'[(\[]*(\d+)[)\].\s]*', normalize_for_match(original_heading))
    expected_num = num_match.group(1) if num_match else None

    lines = full_text.split('\n')
    char_pos = 0
    for line in lines:
        line_core = re.sub(r'\s+', '', normalize_for_match(line))
        if core_text in line_core:
            # 번호 확인: 원본에 번호가 있으면 줄에도 같은 번호 있는지 체크
            if expected_num:
                line_nums = re.findall(r'(\d+)', line[:30])
                if expected_num not in line_nums:
                    char_pos += len(line) + 1
                    continue
            # 줄 전체를 matched_text 로 (앞뒤 공백 제거)
            matched = line.strip()
            if matched:
                results.append((char_pos, matched))
        char_pos += len(line) + 1
    return results


# ---------------------------------------------------------------------------
# v0.7.9 — 마커/번호 자동 내어쓰기 (IndentAtCaret)
# ---------------------------------------------------------------------------

# 기호 마커 (○□■ 등)
_INDENT_MARKERS = set('○□■◆●•◦※➤▶▷►-*·')

# 번호 패턴 (1., 가., 1), (가), ①, (1) 등)
_NUMBER_PATTERNS = [
    re.compile(r'^(\s*\d+[.)]\s)'),           # "1. " or "1) "
    re.compile(r'^(\s*[가-힣][.)]\s)'),        # "가. " or "가) "
    re.compile(r'^(\s*\(\d+\)\s)'),            # "(1) "
    re.compile(r'^(\s*\([가-힣]\)\s)'),        # "(가) "
    re.compile(r'^(\s*[①-⑳]\s?)'),            # "① "
    re.compile(r'^(\s*[IVX]+[.)]\s)'),         # "I. " or "II) "
]


def _apply_indent_at_caret(hwp, text):
    """텍스트의 마커/번호 패턴 감지 → ParagraphShapeIndentAtCaret 자동 내어쓰기.

    동작: 마커/번호 뒤 위치로 커서 이동 → Shift+Tab 효과 (나머지줄 시작위치 설정).
    v0.6.5 복원 + v0.7.9 번호 패턴 확장.
    """
    raw = text.lstrip()
    if not raw:
        return

    # 마커 또는 번호 prefix 길이 계산
    prefix_len = 0

    # 1) 기호 마커 감지
    if raw[0] in _INDENT_MARKERS:
        skip = 0
        while skip < len(text) and text[skip] == ' ':
            skip += 1
        if skip < len(text) and text[skip] in _INDENT_MARKERS:
            skip += 1
        while skip < len(text) and text[skip] == ' ':
            skip += 1
        prefix_len = skip

    # 2) 번호 패턴 감지
    if prefix_len == 0:
        for pat in _NUMBER_PATTERNS:
            m = pat.match(text)
            if m:
                prefix_len = len(m.group(1))
                break

    if prefix_len == 0:
        return

    # v0.7.9-postfix4: GetPos 로 cursor 위치 저장 → SetPos 로 복원.
    # postfix1 (MovePos(3)=DocEnd): 본문이 문서 끝으로 흩어짐.
    # postfix2/3 (MoveLineEnd): visual line wrap 위치로 이동 → BreakPara 가
    #   paragraph 중간을 split → "산업 생태계" + "5,000억 달러" 가
    #   "지능형 산업" + "달러생태계" 로 wrong order.
    # postfix4 (★): cursor 위치를 정확한 (List, Para, Pos) 로 저장/복원.
    #   wrap 무관, 어떤 길이의 line 에도 안전.
    try:
        saved_pos = hwp.GetPos()
    except Exception:
        saved_pos = None

    try:
        hwp.HAction.Run("MovePrevParaBegin")
        for _ in range(prefix_len):
            hwp.HAction.Run("MoveRight")
        hwp.HAction.Run("ParagraphShapeIndentAtCaret")
        if saved_pos is not None:
            hwp.SetPos(*saved_pos)
        else:
            hwp.HAction.Run("MoveLineEnd")
    except Exception as e:
        print(f"[WARN] IndentAtCaret: {e}", file=sys.stderr)
        try:
            if saved_pos is not None:
                hwp.SetPos(*saved_pos)
            else:
                hwp.HAction.Run("MoveLineEnd")
        except Exception:
            pass


def _detect_heading_depth(heading_text):
    """제목 텍스트의 번호 패턴으로 뎁스(깊이) 감지.

    depth 1: "1.", "2.", "I.", "II." — 대제목
    depth 2: "가.", "나.", "다." — 중제목
    depth 3: "1)", "2)", "(1)", "(2)" — 소제목
    depth 4: "(가)", "(나)", "①", "②" — 세부
    depth 0: 감지 불가 → 기본 depth 1 취급

    v0.7.9 Phase 10 fix: ①-⑳ 는 NFKC 정규화 시 "1" 등으로 변환되므로
    원본 텍스트(정규화 전) 에서 먼저 감지.
    """
    # depth 4 (circled number): NFKC 가 ① → "1" 로 변환하므로 원본에서 먼저 검사
    raw = heading_text.lstrip()
    if raw and re.match(r'^[①-⑳]', raw):
        return 4

    text = normalize_for_match(heading_text)
    # depth 4: (가), (나)
    if re.match(r'^[\s]*\([가-힣]\)', text):
        return 4
    # depth 3: 1), 2), (1), (2)
    if re.match(r'^[\s]*\(?\d+\)', text):
        return 3
    # depth 2: 가., 나., 다.
    if re.match(r'^[\s]*[가-힣]\.', text):
        return 2
    # depth 1: 1., 2., I., II., 제1장
    if re.match(r'^[\s]*(?:\d+\.|[IVX]+\.|제\d)', text):
        return 1
    return 1  # 기본 depth

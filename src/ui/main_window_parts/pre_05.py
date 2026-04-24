# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def enum_windows_fast(filter_title_substr=None):
    if IS_MACOS:
        return enumerate_macos_windows(filter_title_substr=filter_title_substr)

    if isinstance(filter_title_substr, str):
        filters = [filter_title_substr.lower()]
    elif filter_title_substr:
        filters = [str(s).lower() for s in filter_title_substr]
    else:
        filters = None

    results = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def _enum_proc(hwnd, lparam):
        try:
            if not _user32.IsWindowVisible(hwnd):
                return True
            title = _win_get_window_text(hwnd)
            if not title:
                return True
            if filters and not any(f in title.lower() for f in filters):
                return True

            cls = _win_get_class_name(hwnd)
            pid = wintypes.DWORD()
            _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            results.append(
                {
                    "handle": int(hwnd),
                    "title": title,
                    "class_name": cls,
                    "pid": pid.value,
                }
            )
        except Exception:
            pass
        return True

    _user32.EnumWindows(_enum_proc, 0)
    return results


# ----------------- 0.3 리소스 경로 헬퍼 (PyInstaller 호환) -----------------
def resource_path(relative_path):
    """
    PyInstaller에서 묶인 리소스 파일을 찾는 경로를 반환합니다.
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = _get_app_base_path()

    return os.path.join(base_path, relative_path)


# ----------------- 1. 프로세스 실행 파일 경로 얻기 -----------------
def get_process_image_path(pid: int) -> Optional[str]:
    if not pid:
        return None
    if not IS_WINDOWS:
        return None

    now = time.monotonic()
    cached = _PROCESS_IMAGE_PATH_CACHE.get(pid)
    if cached and now < float(cached.get("expires_at", 0.0)):
        return cached.get("path")

    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

    # 64비트 안전: use_last_error로 WinAPI 에러 사용 가능
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    OpenProcess = kernel32.OpenProcess
    OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    OpenProcess.restype = wintypes.HANDLE

    QueryFullProcessImageNameW = kernel32.QueryFullProcessImageNameW
    QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE,
        wintypes.DWORD,
        wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD),
    ]
    QueryFullProcessImageNameW.restype = wintypes.BOOL

    CloseHandle = kernel32.CloseHandle
    CloseHandle.argtypes = [wintypes.HANDLE]
    CloseHandle.restype = wintypes.BOOL

    hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not hProcess:
        _PROCESS_IMAGE_PATH_CACHE[pid] = {
            "expires_at": now + _PROCESS_IMAGE_PATH_CACHE_TTL_SEC,
            "path": None,
        }
        return None
    try:
        # 1차 버퍼
        size = 512
        while True:
            buf_len = wintypes.DWORD(size)
            buf = ctypes.create_unicode_buffer(buf_len.value)
            ok = QueryFullProcessImageNameW(hProcess, 0, buf, ctypes.byref(buf_len))
            if ok:
                path = buf.value
                _PROCESS_IMAGE_PATH_CACHE[pid] = {
                    "expires_at": now + _PROCESS_IMAGE_PATH_CACHE_TTL_SEC,
                    "path": path,
                }
                return path
            # 버퍼 부족 시 한 번 정도 키워 봄
            err = ctypes.get_last_error()
            # ERROR_INSUFFICIENT_BUFFER = 122
            if err == 122 and size < 4096:
                size *= 2
                continue
            _PROCESS_IMAGE_PATH_CACHE[pid] = {
                "expires_at": now + _PROCESS_IMAGE_PATH_CACHE_TTL_SEC,
                "path": None,
            }
            return None
    finally:
        CloseHandle(hProcess)


# ----------------- 1.1 엄격한 OneNote 창 검증 헬퍼 -----------------
def is_strict_onenote_window(w: Dict[str, Any], my_pid: int) -> bool:
    """주어진 창 정보가 실제로 OneNote 앱 창인지 엄격하게 확인합니다."""
    if IS_MACOS:
        return is_macos_onenote_window_info(w, my_pid)

    if w.get("pid") == my_pid:
        return False

    title_lower = w.get("title", "").lower()
    cls = w.get("class_name", "")
    pid = w.get("pid")

    # 1. Classic Desktop (OMain*) - 레거시 OneNote
    if "omain" in (cls or "").lower():
        return True

    # 2. Modern App (ApplicationFrameWindow) + 타이틀 키워드
    if cls == ONENOTE_CLASS_NAME and (
        "onenote" in title_lower or "원노트" in title_lower
    ):
        return True

    # 3. Fallback: 제목에 키워드 + EXE 확인
    if "onenote" in title_lower or "원노트" in title_lower:
        exe_path = get_process_image_path(pid)
        if exe_path:
            exe_name = os.path.basename(exe_path).lower()
            if "onenote.exe" in exe_name or "onenoteim.exe" in exe_name:
                return True

    return False


# ----------------- 4. 짧은 폴링으로 Rect 안정화 대기 -----------------
def _wait_rect_settle(get_rect, timeout=0.3, interval=0.03):
    start = time.perf_counter()
    prev = get_rect()
    while time.perf_counter() - start < timeout:
        time.sleep(interval)
        cur = get_rect()
        if abs(cur.top - prev.top) < 2 and abs(cur.bottom - prev.bottom) < 2:
            break
        prev = cur


# ----------------- 5. 패턴 기반 수직 스크롤 시도 -----------------
def _scroll_vertical_via_pattern(
    container, direction: str, small=True, repeats=1
) -> bool:
    ensure_pywinauto()
    if not _pwa_ready:
        return False
    try:
        iface = getattr(container, "iface_scroll", None)
        if iface is None:
            return False

        from comtypes.gen.UIAutomationClient import (
            ScrollAmount_LargeIncrement,
            ScrollAmount_LargeDecrement,
            ScrollAmount_SmallIncrement,
            ScrollAmount_SmallDecrement,
            ScrollAmount_NoAmount,
        )

        v_inc = ScrollAmount_SmallIncrement if small else ScrollAmount_LargeIncrement
        v_dec = ScrollAmount_SmallDecrement if small else ScrollAmount_LargeDecrement
        v_amount = v_inc if direction == "down" else v_dec

        for _ in range(max(1, repeats)):
            iface.Scroll(ScrollAmount_NoAmount, v_amount)
        return True
    except Exception:
        return False


# ----------------- 6. 마우스 휠 기반 스크롤(폴백) -----------------
def _safe_wheel(scroll_container, steps: int):
    if steps == 0:
        return

    ensure_pywinauto()

    try:
        if hasattr(scroll_container, "wheel_scroll"):
            scroll_container.wheel_scroll(steps)
            return
    except Exception:
        pass

    try:
        if hasattr(scroll_container, "wheel_mouse_input"):
            scroll_container.wheel_mouse_input(wheel_dist=steps)
            return
    except Exception:
        pass

    try:
        rect = scroll_container.rectangle()
        center = rect.mid_point()
        try:
            mouse.scroll(coords=(center.x, center.y), wheel_dist=steps)
            return
        except Exception:
            pass
        try:
            mouse.wheel(coords=(center.x, center.y), wheel_dist=steps)
            return
        except Exception:
            pass
    except Exception:
        pass

    try:
        scroll_container.set_focus()
        if steps > 0:
            keyboard.send_keys("{UP %d}" % steps)
        else:
            keyboard.send_keys("{DOWN %d}" % abs(steps))
    except Exception:
        pass


# ----------------- 7. 선택 항목을 가장 빠르게 얻기 -----------------
def _wrapper_identity_key(ctrl):
    try:
        rect = ctrl.rectangle()
        return (
            _safe_window_text(ctrl),
            _safe_control_type(ctrl),
            rect.left,
            rect.top,
            rect.right,
            rect.bottom,
        )
    except Exception:
        return (id(ctrl),)


def _control_depth_within_tree(ctrl, tree_control) -> int:
    depth = 0
    current = ctrl
    for _ in range(20):
        current = _safe_parent(current)
        if current is None:
            break
        depth += 1
        if current == tree_control:
            break
    return depth


def _pick_best_tree_item_candidate(tree_control, candidates):
    best = None
    best_score = None
    for item in candidates:
        if item is None:
            continue
        try:
            focus = 1 if item.has_keyboard_focus() else 0
        except Exception:
            focus = 0
        try:
            selected = 1 if item.is_selected() else 0
        except Exception:
            selected = 0
        depth = _control_depth_within_tree(item, tree_control)
        try:
            rect = item.rectangle()
            height = max(1, rect.bottom - rect.top)
        except Exception:
            height = 9999
        score = (focus, depth, selected, -height)
        if best_score is None or score > best_score:
            best = item
            best_score = score
    return best

_publish_context(globals())

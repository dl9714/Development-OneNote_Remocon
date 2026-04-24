# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _ax_number_attribute(element: c_void_p, attr_name: str) -> int:
    value = _ax_copy_attribute(element, attr_name)
    if not value:
        return 0
    try:
        return _cf_number_to_int(value)
    finally:
        _cf_release(value)


def _ax_point_attribute(element: c_void_p, attr_name: str) -> Optional[_CGPoint]:
    value = _ax_copy_attribute(element, attr_name)
    if not value:
        return None
    try:
        point = _CGPoint()
        ok = bool(
            _APP_SERVICES.AXValueGetValue(
                value,
                _KAX_VALUE_CGPOINT_TYPE,
                ctypes.byref(point),
            )
        )
        return point if ok else None
    except Exception:
        return None
    finally:
        _cf_release(value)


def _ax_size_attribute(element: c_void_p, attr_name: str) -> Optional[_CGSize]:
    value = _ax_copy_attribute(element, attr_name)
    if not value:
        return None
    try:
        size = _CGSize()
        ok = bool(
            _APP_SERVICES.AXValueGetValue(
                value,
                _KAX_VALUE_CGSIZE_TYPE,
                ctypes.byref(size),
            )
        )
        return size if ok else None
    except Exception:
        return None
    finally:
        _cf_release(value)


def _cg_click_screen(x: float, y: float, *, click_count: int = 1) -> bool:
    if not (IS_MACOS and _APP_SERVICES):
        return False
    try:
        point = _CGPoint(float(x), float(y))
        count = max(1, int(click_count or 1))
        for _ in range(count):
            for event_type in (_KCG_EVENT_LEFT_MOUSE_DOWN, _KCG_EVENT_LEFT_MOUSE_UP):
                event_ref = _APP_SERVICES.CGEventCreateMouseEvent(
                    None,
                    event_type,
                    point,
                    _KCG_MOUSE_BUTTON_LEFT,
                )
                if not event_ref:
                    return False
                try:
                    _APP_SERVICES.CGEventPost(_KCG_HID_EVENT_TAP, event_ref)
                finally:
                    _cf_release(event_ref)
                time.sleep(0.04)
            time.sleep(0.08)
        return True
    except Exception:
        return False


def _ax_perform_action(element: c_void_p, action_name: str = "AXPress") -> bool:
    if not (IS_MACOS and _APP_SERVICES and element and action_name):
        return False
    action_ref = _cf_string(action_name)
    if not action_ref:
        return False
    try:
        err = int(_APP_SERVICES.AXUIElementPerformAction(element, action_ref))
        # OneNote sometimes returns kAXErrorCannotComplete after it has already
        # performed the action. Treat it as an attempted press and verify state
        # in the caller instead of falling back to coordinate clicks too early.
        return err in {0, -25205, -25206}
    except Exception:
        return False
    finally:
        _cf_release(action_ref)


def _ax_click_element_center(
    element: c_void_p,
    *,
    click_count: int = 1,
    preferred_x_offset: Optional[float] = None,
) -> bool:
    point = _ax_point_attribute(element, "AXPosition")
    size = _ax_size_attribute(element, "AXSize")
    if not point or not size:
        return False
    width = max(1.0, float(size.width))
    height = max(1.0, float(size.height))
    if preferred_x_offset is None:
        click_x = float(point.x) + (width / 2.0)
    else:
        click_x = float(point.x) + max(1.0, min(width - 1.0, float(preferred_x_offset)))
    click_y = float(point.y) + (height / 2.0)
    return _cg_click_screen(click_x, click_y, click_count=click_count)


def _ax_element_attribute(element: c_void_p, attr_name: str) -> Optional[c_void_p]:
    value = _ax_copy_attribute(element, attr_name)
    if not value:
        return None
    return value


def _ax_array_attribute(element: c_void_p, attr_name: str) -> List[c_void_p]:
    value = _ax_copy_attribute(element, attr_name)
    if not value:
        return []
    refs: List[c_void_p] = []
    try:
        if _cf_type_id(value) != int(_CF.CFArrayGetTypeID()):
            return []
        count = int(_CF.CFArrayGetCount(value))
        for index in range(max(0, count)):
            item = c_void_p(_CF.CFArrayGetValueAtIndex(value, index))
            if not item:
                continue
            try:
                refs.append(c_void_p(_CF.CFRetain(item)))
            except Exception:
                refs.append(item)
        return refs
    finally:
        _cf_release(value)


def _release_ax_refs(refs: Sequence[c_void_p]) -> None:
    for ref in refs:
        _cf_release(ref)


def macos_accessibility_is_trusted() -> bool:
    if not IS_MACOS:
        return True
    if not _APP_SERVICES:
        return False
    try:
        return bool(_APP_SERVICES.AXIsProcessTrusted())
    except Exception:
        return False


def macos_last_ax_notebook_debug() -> Dict[str, Any]:
    return dict(_MAC_LAST_AX_NOTEBOOK_DEBUG)


def _bundle_id_for_pid(pid: int) -> str:
    if pid in _MAC_BUNDLE_ID_CACHE:
        return _MAC_BUNDLE_ID_CACHE[pid]
    if not pid:
        return ""
    try:
        completed = subprocess.run(
            [MAC_LSAPPINFO_PATH, "info", "-only", "bundleid", str(int(pid))],
            text=True,
            capture_output=True,
            encoding="utf-8",
            errors="replace",
            timeout=5,
        )
        raw = (completed.stdout or "") + "\n" + (completed.stderr or "")
        marker = '"CFBundleIdentifier"="'
        if marker in raw:
            bundle_id = raw.split(marker, 1)[1].split('"', 1)[0].strip()
        else:
            bundle_id = ""
    except Exception:
        bundle_id = ""
    _MAC_BUNDLE_ID_CACHE[pid] = bundle_id
    return bundle_id


def _enumerate_macos_windows_via_coregraphics(filter_title_substr=None) -> List[Dict[str, Any]]:
    if not (_CF and _APP_SERVICES):
        return []

    if isinstance(filter_title_substr, str):
        filters = [_normalize_text(filter_title_substr)]
    elif filter_title_substr:
        filters = [_normalize_text(str(item)) for item in filter_title_substr]
    else:
        filters = []

    key_refs = {key: _cf_string(key) for key in _CG_WINDOW_INFO_KEYS}
    array_ref = c_void_p(
        _APP_SERVICES.CGWindowListCopyWindowInfo(
            _KCG_WINDOW_LIST_OPTION_ON_SCREEN_ONLY
            | _KCG_WINDOW_LIST_EXCLUDE_DESKTOP_ELEMENTS,
            0,
        )
    )
    if not array_ref:
        for ref in key_refs.values():
            if ref:
                _CF.CFRelease(ref)
        return []

    results: List[Dict[str, Any]] = []
    try:
        count = _CF.CFArrayGetCount(array_ref)
        for index in range(max(0, int(count))):
            entry = c_void_p(_CF.CFArrayGetValueAtIndex(array_ref, index))
            if not entry:
                continue
            layer = _cf_number_to_int(
                c_void_p(_CF.CFDictionaryGetValue(entry, key_refs["kCGWindowLayer"]))
            )
            if layer != 0:
                continue
            pid = _cf_number_to_int(
                c_void_p(_CF.CFDictionaryGetValue(entry, key_refs["kCGWindowOwnerPID"]))
            )
            window_number = _cf_number_to_int(
                c_void_p(_CF.CFDictionaryGetValue(entry, key_refs["kCGWindowNumber"]))
            )
            app_name = _clean_field(
                _cf_string_to_text(
                    c_void_p(_CF.CFDictionaryGetValue(entry, key_refs["kCGWindowOwnerName"]))
                )
            )
            title = _clean_field(
                _cf_string_to_text(
                    c_void_p(_CF.CFDictionaryGetValue(entry, key_refs["kCGWindowName"]))
                )
            )
            if not pid or not app_name:
                continue
            if filters and not any(token in _normalize_text(title) for token in filters):
                continue
            bundle_id = _bundle_id_for_pid(pid)
            handle = _synthetic_window_handle(
                pid,
                f"{window_number}:{title}",
                bundle_id or app_name,
            )
            results.append(
                {
                    "handle": handle,
                    "title": title,
                    "class_name": bundle_id or app_name,
                    "pid": pid,
                    "bundle_id": bundle_id,
                    "app_name": app_name,
                    "frontmost": False,
                    "window_number": window_number,
                }
            )
    finally:
        for ref in key_refs.values():
            if ref:
                _CF.CFRelease(ref)
        _CF.CFRelease(array_ref)

    return results


class MacRect:
    __slots__ = ("left", "top", "right", "bottom")

    def __init__(self, left: int, top: int, right: int, bottom: int):
        self.left = left
        self.top = top
        self.right = right
        self.bottom = bottom

    def mid_point(self):
        return type("Point", (), {"x": int((self.left + self.right) / 2), "y": int((self.top + self.bottom) / 2)})()


def _applescript_window_locator(pid: int, title: str) -> str:
    title_literal = _quote_applescript_text(title or "")
    return f"""
tell application "System Events"
    set targetProcess to first application process whose unix id is {int(pid)}
    set targetWindow to missing value
    repeat with w in windows of targetProcess
        try
            if {_quote_applescript_text(title or "")} is "" or (name of w as text) is {title_literal} then
                set targetWindow to w
                exit repeat
            end if
        end try
    end repeat
    if targetWindow is missing value then
        if (count of windows of targetProcess) is 0 then error "OneNote 창을 찾지 못했습니다."
        set targetWindow to window 1 of targetProcess
    end if
"""

_publish_context(globals())

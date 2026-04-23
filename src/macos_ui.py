# -*- coding: utf-8 -*-
"""
macOS용 OneNote UI 자동화 헬퍼.

OneNote for Mac의 화면 구조를 기준으로 System Events(접근성)와 osascript를 사용한다.
"""

from __future__ import annotations

import ctypes
import hashlib
import json
import os
import subprocess
import threading
import time
import unicodedata
from ctypes import (
    c_bool,
    c_char_p,
    c_int,
    c_long,
    c_longlong,
    c_uint32,
    c_void_p,
    create_string_buffer,
)
from ctypes.util import find_library
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
from urllib.parse import quote, unquote, urlparse

from src.platform_support import (
    IS_MACOS,
    MAC_OSASCRIPT_PATH,
    ONENOTE_MAC_APP_NAMES,
    ONENOTE_MAC_BUNDLE_ID,
    open_url_in_system,
)

MAC_LSAPPINFO_PATH = "/usr/bin/lsappinfo"
_KCF_STRING_ENCODING_UTF8 = 0x08000100
_KCF_NUMBER_LONGLONG_TYPE = 11
_KCG_WINDOW_LIST_OPTION_ON_SCREEN_ONLY = 1
_KCG_WINDOW_LIST_EXCLUDE_DESKTOP_ELEMENTS = 16
_CG_WINDOW_INFO_KEYS = (
    "kCGWindowOwnerPID",
    "kCGWindowOwnerName",
    "kCGWindowName",
    "kCGWindowNumber",
    "kCGWindowLayer",
)

if IS_MACOS:
    _CF = ctypes.CDLL(find_library("CoreFoundation"))
    _APP_SERVICES = ctypes.CDLL(find_library("ApplicationServices"))
    _APP_SERVICES.CGWindowListCopyWindowInfo.argtypes = [c_uint32, c_uint32]
    _APP_SERVICES.CGWindowListCopyWindowInfo.restype = c_void_p
    _CF.CFArrayGetCount.argtypes = [c_void_p]
    _CF.CFArrayGetCount.restype = c_long
    _CF.CFArrayGetValueAtIndex.argtypes = [c_void_p, c_long]
    _CF.CFArrayGetValueAtIndex.restype = c_void_p
    _CF.CFDictionaryGetValue.argtypes = [c_void_p, c_void_p]
    _CF.CFDictionaryGetValue.restype = c_void_p
    _CF.CFStringCreateWithCString.argtypes = [c_void_p, c_char_p, c_uint32]
    _CF.CFStringCreateWithCString.restype = c_void_p
    _CF.CFStringGetCString.argtypes = [c_void_p, c_char_p, c_long, c_uint32]
    _CF.CFStringGetCString.restype = c_bool
    _CF.CFNumberGetValue.argtypes = [c_void_p, c_int, c_void_p]
    _CF.CFNumberGetValue.restype = c_bool
    _CF.CFGetTypeID.argtypes = [c_void_p]
    _CF.CFGetTypeID.restype = c_long
    _CF.CFArrayGetTypeID.argtypes = []
    _CF.CFArrayGetTypeID.restype = c_long
    _CF.CFStringGetTypeID.argtypes = []
    _CF.CFStringGetTypeID.restype = c_long
    _CF.CFRetain.argtypes = [c_void_p]
    _CF.CFRetain.restype = c_void_p
    _CF.CFRelease.argtypes = [c_void_p]
    _CF.CFRelease.restype = None
    _APP_SERVICES.AXUIElementCreateApplication.argtypes = [c_int]
    _APP_SERVICES.AXUIElementCreateApplication.restype = c_void_p
    _APP_SERVICES.AXUIElementCopyAttributeValue.argtypes = [
        c_void_p,
        c_void_p,
        ctypes.POINTER(c_void_p),
    ]
    _APP_SERVICES.AXUIElementCopyAttributeValue.restype = c_int
    _APP_SERVICES.AXIsProcessTrusted.argtypes = []
    _APP_SERVICES.AXIsProcessTrusted.restype = c_bool
else:
    _CF = None
    _APP_SERVICES = None

_MAC_BUNDLE_ID_CACHE: Dict[int, str] = {}
_MAC_ONENOTE_NOTEBOOKS_PLIST = Path(
    os.path.expanduser(
        "~/Library/Group Containers/UBF8T346G9.Office/OneNote/ShareExtension/Notebooks.plist"
    )
)
_MAC_ONENOTE_RESOURCEINFOCACHE_JSON = Path(
    os.path.expanduser(
        "~/Library/Containers/com.microsoft.onenote.mac/Data/Library/Application Support/Microsoft/Office/16.0/ResourceInfoCache/data.json"
    )
)
_MAC_RECENT_CACHE_TIMED_OUT = False
_MAC_LAST_AX_NOTEBOOK_DEBUG: Dict[str, Any] = {}


class MacAutomationError(RuntimeError):
    """macOS 접근성 자동화 실패."""


def _quote_applescript_text(text: str) -> str:
    return '"' + (text or "").replace("\\", "\\\\").replace('"', '\\"') + '"'


def _run_osascript(script: str, timeout: int = 20) -> str:
    if not IS_MACOS:
        raise MacAutomationError("macOS 전용 기능입니다.")
    completed = subprocess.run(
        [MAC_OSASCRIPT_PATH, "-"],
        input=script,
        text=True,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        timeout=max(3, timeout),
    )
    if completed.returncode != 0:
        raise MacAutomationError((completed.stderr or completed.stdout or "").strip())
    return (completed.stdout or "").strip()


def _read_macos_clipboard_text() -> str:
    if not IS_MACOS:
        return ""
    try:
        completed = subprocess.run(
            ["/usr/bin/pbpaste"],
            capture_output=True,
            timeout=3,
        )
        return completed.stdout.decode("utf-8", "replace")
    except Exception:
        return ""


def _write_macos_clipboard_text(text: str) -> bool:
    if not IS_MACOS:
        return False
    try:
        subprocess.run(
            ["/usr/bin/pbcopy"],
            input=(text or "").encode("utf-8"),
            capture_output=True,
            timeout=3,
            check=True,
        )
        return True
    except Exception:
        return False


def _synthetic_window_handle(pid: int, title: str, bundle_id: str) -> int:
    raw = f"{pid}\0{title}\0{bundle_id}"
    return int(hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12], 16)


def _clean_field(value: str) -> str:
    text = unicodedata.normalize("NFC", str(value or ""))
    return text.replace("\t", " ").replace("\r", " ").replace("\n", " ").strip()


def _normalize_text(text: Optional[str]) -> str:
    return " ".join((unicodedata.normalize("NFC", text or "").strip().split())).casefold()


def _rect_key(x: int, y: int, w: int, h: int) -> Tuple[int, int, int, int]:
    return int(x), int(y), int(w), int(h)


def _cf_string(value: str) -> c_void_p:
    if not (_CF and value):
        return c_void_p()
    return c_void_p(
        _CF.CFStringCreateWithCString(
            None,
            value.encode("utf-8"),
            _KCF_STRING_ENCODING_UTF8,
        )
    )


def _cf_string_to_text(value: c_void_p) -> str:
    if not (_CF and value):
        return ""
    buf = create_string_buffer(4096)
    ok = _CF.CFStringGetCString(
        value,
        buf,
        len(buf),
        _KCF_STRING_ENCODING_UTF8,
    )
    if not ok:
        return ""
    return buf.value.decode("utf-8", "replace")


def _cf_number_to_int(value: c_void_p) -> int:
    if not (_CF and value):
        return 0
    out = c_longlong()
    ok = _CF.CFNumberGetValue(
        value,
        _KCF_NUMBER_LONGLONG_TYPE,
        ctypes.byref(out),
    )
    if not ok:
        return 0
    return int(out.value)


def _cf_type_id(value: c_void_p) -> int:
    if not (_CF and value):
        return 0
    try:
        return int(_CF.CFGetTypeID(value))
    except Exception:
        return 0


def _cf_release(value: c_void_p) -> None:
    if not (_CF and value):
        return
    try:
        _CF.CFRelease(value)
    except Exception:
        pass


def _ax_copy_attribute(element: c_void_p, attr_name: str) -> Optional[c_void_p]:
    if not (IS_MACOS and _APP_SERVICES and _CF and element and attr_name):
        return None
    attr_ref = _cf_string(attr_name)
    if not attr_ref:
        return None
    out = c_void_p()
    try:
        err = _APP_SERVICES.AXUIElementCopyAttributeValue(
            element,
            attr_ref,
            ctypes.byref(out),
        )
        if int(err) != 0 or not out:
            return None
        return c_void_p(out.value)
    except Exception:
        return None
    finally:
        _cf_release(attr_ref)


def _ax_text_attribute(element: c_void_p, attr_name: str) -> str:
    value = _ax_copy_attribute(element, attr_name)
    if not value:
        return ""
    try:
        if _cf_type_id(value) != int(_CF.CFStringGetTypeID()):
            return ""
        return _clean_field(_cf_string_to_text(value))
    finally:
        _cf_release(value)


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


@dataclass
class MacRect:
    left: int
    top: int
    right: int
    bottom: int

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


def _enumerate_macos_windows_via_system_events(filter_title_substr=None) -> List[Dict[str, Any]]:
    """보이는 macOS 앱 윈도우 목록."""
    if not IS_MACOS:
        return []

    script = r'''
tell application "System Events"
    set resultItems to {}
    repeat with p in (every application process whose background only is false)
        try
            set procName to name of p as text
            set procPid to unix id of p
            set procFront to frontmost of p as text
            try
                set procBundle to bundle identifier of p as text
            on error
                set procBundle to ""
            end try
            repeat with w in windows of p
                try
                    set winName to name of w as text
                on error
                    set winName to ""
                end try
                set winName to my cleanText(winName)
                set procName to my cleanText(procName)
                set procBundle to my cleanText(procBundle)
                set end of resultItems to procName & tab & procPid & tab & procFront & tab & procBundle & tab & winName
            end repeat
        end try
    end repeat
    set AppleScript's text item delimiters to linefeed
    return resultItems as text
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    raw = _run_osascript(script, timeout=20)
    if not raw:
        return []

    if isinstance(filter_title_substr, str):
        filters = [_normalize_text(filter_title_substr)]
    elif filter_title_substr:
        filters = [_normalize_text(str(item)) for item in filter_title_substr]
    else:
        filters = []

    results: List[Dict[str, Any]] = []
    for line in raw.splitlines():
        parts = line.split("\t")
        if len(parts) < 5:
            continue
        app_name, pid_text, front_text, bundle_id, title = parts[:5]
        try:
            pid = int(pid_text)
        except Exception:
            continue
        if filters and not any(token in _normalize_text(title) for token in filters):
            continue
        handle = _synthetic_window_handle(pid, title, bundle_id)
        results.append(
            {
                "handle": handle,
                "title": title,
                "class_name": bundle_id or app_name,
                "pid": pid,
                "bundle_id": bundle_id,
                "app_name": app_name,
                "frontmost": front_text == "true",
            }
        )
    return results


def _read_onenote_current_notebook_name(pid: int, title_hint: str = "") -> str:
    if not pid:
        return ""
    locator = _applescript_window_locator(pid, title_hint or "")
    script = locator + r'''
        set currentNotebook to ""
        repeat with e in (entire contents of targetWindow)
            try
                if role of e is "AXButton" then
                    set candidateText to ""
                    try
                        set candidateText to description of e as text
                    end try
                    if candidateText does not contain "(현재 전자필기장)" and candidateText does not contain "(현재 전자 필기장)" then
                        try
                            set candidateText to value of e as text
                        end try
                    end if
                    if candidateText does not contain "(현재 전자필기장)" and candidateText does not contain "(현재 전자 필기장)" then
                        try
                            set candidateText to title of e as text
                        end try
                    end if
                    set candidateText to my cleanText(candidateText)
                    if candidateText contains "(현재 전자필기장)" or candidateText contains "(현재 전자 필기장)" then
                        set currentNotebook to candidateText
                        exit repeat
                    end if
                end if
            end try
        end repeat
        return my cleanText(currentNotebook)

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    try:
        raw = _run_osascript(script, timeout=10)
    except Exception:
        return ""
    return _extract_current_notebook_name(raw)


def _hydrate_missing_macos_window_titles(
    results: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    if not results:
        return results

    missing = [
        info for info in results if not _clean_field(str(info.get("title") or ""))
    ]
    if not missing:
        return results

    fallback_results = _enumerate_macos_windows_via_system_events(
        filter_title_substr=None
    )
    fallback_titles: Dict[Tuple[int, str], List[str]] = {}
    for info in fallback_results:
        title = _clean_field(str(info.get("title") or ""))
        if not title:
            continue
        pid = int(info.get("pid") or 0)
        if not pid:
            continue
        bundle_id = str(info.get("bundle_id") or "")
        app_name = str(info.get("app_name") or "")
        for key in (
            (pid, bundle_id),
            (pid, app_name),
            (pid, ""),
        ):
            fallback_titles.setdefault(key, []).append(title)

    notebook_name_cache: Dict[int, str] = {}
    for info in results:
        if _clean_field(str(info.get("title") or "")):
            continue

        pid = int(info.get("pid") or 0)
        bundle_id = str(info.get("bundle_id") or "")
        app_name = str(info.get("app_name") or "")

        matched_titles = (
            fallback_titles.get((pid, bundle_id))
            or fallback_titles.get((pid, app_name))
            or fallback_titles.get((pid, ""))
            or []
        )
        if matched_titles:
            info["title"] = matched_titles[0]
            continue

        if bundle_id == ONENOTE_MAC_BUNDLE_ID or app_name in ONENOTE_MAC_APP_NAMES:
            if pid not in notebook_name_cache:
                notebook_name_cache[pid] = _read_onenote_current_notebook_name(pid)
            notebook_name = notebook_name_cache.get(pid) or ""
            if notebook_name:
                info["title"] = notebook_name

    return results


def enumerate_macos_windows(filter_title_substr=None) -> List[Dict[str, Any]]:
    results = _enumerate_macos_windows_via_coregraphics(
        filter_title_substr=filter_title_substr
    )
    if results:
        return _hydrate_missing_macos_window_titles(results)
    return _hydrate_missing_macos_window_titles(
        _enumerate_macos_windows_via_system_events(
        filter_title_substr=filter_title_substr
        )
    )


def enumerate_macos_windows_quick(filter_title_substr=None) -> List[Dict[str, Any]]:
    """Finder 실행 초기처럼 민감한 구간에서 쓰는 가벼운 창 열거."""
    return _enumerate_macos_windows_via_coregraphics(
        filter_title_substr=filter_title_substr
    )


class MacDesktop:
    """pywinauto Desktop 대체."""

    def __init__(self, backend: str = "uia"):
        self.backend = backend

    def window(self, handle=None):
        if handle is None:
            raise MacAutomationError("window handle이 필요합니다.")
        for info in enumerate_macos_windows():
            if int(info.get("handle") or 0) == int(handle):
                return MacWindow(info)
        raise MacAutomationError("해당 macOS 윈도우를 찾지 못했습니다.")


@dataclass
class MacWindow:
    info: Dict[str, Any]

    @property
    def handle(self) -> int:
        return int(self.info.get("handle") or 0)

    def window_text(self) -> str:
        title = _clean_field(str(self.info.get("title") or ""))
        if title:
            return title

        handle = self.handle
        pid = self.process_id()
        bundle_id = self.bundle_id()
        for info in enumerate_macos_windows():
            same_handle = handle and int(info.get("handle") or 0) == handle
            same_pid = pid and int(info.get("pid") or 0) == pid
            same_bundle = bundle_id and str(info.get("bundle_id") or "") == bundle_id
            if same_handle or (same_pid and same_bundle):
                self.info.update(info)
                title = _clean_field(str(info.get("title") or ""))
                if title:
                    return title

        if bundle_id == ONENOTE_MAC_BUNDLE_ID or self.app_name() in ONENOTE_MAC_APP_NAMES:
            title = _read_onenote_current_notebook_name(pid)
            if title:
                self.info["title"] = title
                return title

        return ""

    def class_name(self) -> str:
        return str(self.info.get("bundle_id") or self.info.get("class_name") or self.info.get("app_name") or "")

    def process_id(self) -> int:
        return int(self.info.get("pid") or 0)

    def bundle_id(self) -> str:
        return str(self.info.get("bundle_id") or "")

    def app_name(self) -> str:
        return str(self.info.get("app_name") or "")

    def is_visible(self) -> bool:
        handle = self.handle
        pid = self.process_id()
        title = self.window_text()
        bundle_id = self.bundle_id()
        for info in enumerate_macos_windows():
            if handle and int(info.get("handle") or 0) == handle:
                self.info.update(info)
                return True
            if (
                pid
                and bundle_id
                and int(info.get("pid") or 0) == pid
                and str(info.get("bundle_id") or "") == bundle_id
                and (
                    not title
                    or _clean_field(str(info.get("title") or "")) == title
                )
            ):
                self.info.update(info)
                return True
            if int(info.get("pid") or 0) == pid and str(info.get("title") or "") == title:
                self.info.update(info)
                return True
        return False

    def set_focus(self) -> None:
        bundle_id = self.bundle_id()
        app_name = self.app_name()
        if not bundle_id and not app_name:
            raise MacAutomationError("활성화할 앱 식별자를 찾지 못했습니다.")
        app_activate = (
            f'tell application id {_quote_applescript_text(bundle_id)} to activate'
            if bundle_id
            else f'tell application {_quote_applescript_text(app_name)} to activate'
        )
        script = f'''
{app_activate}
delay 0.1
tell application "System Events"
    tell first application process whose unix id is {self.process_id()}
        try
            perform action "AXRaise" of (first window whose name is {_quote_applescript_text(self.window_text())})
        end try
    end tell
end tell
'''
        _run_osascript(script, timeout=10)

    def child_window(self, control_type=None, found_index=0):
        if control_type in ("Tree", "List", None):
            return MacTreeControl(self)
        raise MacAutomationError(f"지원하지 않는 macOS child_window 요청: {control_type}")


@dataclass
class MacRow:
    window: MacWindow
    text: str
    selected: bool
    rect: MacRect
    scroll_rect: MacRect
    order: int
    detail: str = ""

    def window_text(self) -> str:
        return self.text

    def rectangle(self) -> MacRect:
        return self.rect

    def is_selected(self) -> bool:
        return bool(self.selected)

    def has_keyboard_focus(self) -> bool:
        return bool(self.selected)

    def select(self) -> None:
        select_row_by_text(self.window, self.text, preferred_scroll_left=self.scroll_rect.left)

    def click_input(self) -> None:
        self.select()


class MacTreeControl:
    """선택/검색/정렬 용도로 쓰는 단순 트리 래퍼."""

    def __init__(self, window: MacWindow):
        self.window = window

    def wrapper_object(self):
        return self

    def children(self) -> List[MacRow]:
        return collect_onenote_rows(self.window)

    def descendants(self, control_type=None) -> List[MacRow]:
        if control_type in (None, "TreeItem", "ListItem"):
            return self.children()
        return []

    def get_selection(self) -> List[MacRow]:
        return [row for row in self.children() if row.selected]

    def set_focus(self) -> None:
        self.window.set_focus()


def _extract_current_notebook_name(raw_text: str) -> str:
    text = _clean_field(raw_text)
    for marker in ("(현재 전자필기장)", "(현재 전자 필기장)"):
        if marker in text:
            text = text.split(marker, 1)[0].strip()
            break
    if "," in text:
        text = text.split(",", 1)[0].strip()
    return text


def collect_onenote_snapshot(window: MacWindow, timeout: int = 20) -> Dict[str, Any]:
    locator = _applescript_window_locator(window.process_id(), window.window_text())
    script = locator + r'''
        set resultItems to {}
        set currentNotebook to ""
        set allElems to entire contents of targetWindow

        repeat with e in allElems
            try
                if role of e is "AXButton" then
                    set candidateText to ""
                    try
                        set candidateText to description of e as text
                    end try
                    if candidateText does not contain "(현재 전자필기장)" and candidateText does not contain "(현재 전자 필기장)" then
                        try
                            set candidateText to value of e as text
                        end try
                    end if
                    if candidateText does not contain "(현재 전자필기장)" and candidateText does not contain "(현재 전자 필기장)" then
                        try
                            set candidateText to title of e as text
                        end try
                    end if
                    set candidateText to my cleanText(candidateText)
                    if candidateText contains "(현재 전자필기장)" or candidateText contains "(현재 전자 필기장)" then
                        set currentNotebook to candidateText
                    end if
                end if
            end try
        end repeat

        if currentNotebook is not "" then
            set end of resultItems to "NOTEBOOK" & tab & my cleanText(currentNotebook)
        end if

        repeat with e in allElems
            try
                if role of e is "AXScrollArea" then
                    set pp to position of e
                    set ss to size of e
                    set vx to ""
                    repeat with b in (every UI element of e)
                        try
                            if role of b is "AXScrollBar" then
                                try
                                    set bo to orientation of b as text
                                on error
                                    set bo to ""
                                end try
                                if bo is "AXVerticalOrientation" then
                                    try
                                        set vx to value of b as text
                                    on error
                                        set vx to ""
                                    end try
                                end if
                            end if
                        end try
                    end repeat
                    set end of resultItems to "SCROLL" & tab & (item 1 of pp) & tab & (item 2 of pp) & tab & (item 1 of ss) & tab & (item 2 of ss) & tab & vx
                end if
            end try
        end repeat

        set rowIndex to 0
        repeat with e in allElems
            try
                if role of e is "AXRow" then
                    set rowIndex to rowIndex + 1
                    set labelText to ""
                    set detailText to ""
                    set childElems to entire contents of e
                    repeat with c in childElems
                        try
                            if role of c is "AXStaticText" then
                                set labelText to my cleanText(value of c as text)
                                exit repeat
                            end if
                        end try
                    end repeat
                    if labelText is "" then
                        repeat with c in childElems
                            try
                                set maybeText to my cleanText(description of c as text)
                                if maybeText is not "" then
                                    set labelText to maybeText
                                    exit repeat
                                end if
                            end try
                            try
                                set maybeText to my cleanText(title of c as text)
                                if maybeText is not "" then
                                    set labelText to maybeText
                                    exit repeat
                                end if
                            end try
                        end repeat
                    end if

                    repeat with c in childElems
                        try
                            set maybeDetail to my cleanText(description of c as text)
                            if maybeDetail is not "" and maybeDetail is not labelText then
                                set detailText to maybeDetail
                                exit repeat
                            end if
                        end try
                        try
                            set maybeDetail to my cleanText(title of c as text)
                            if maybeDetail is not "" and maybeDetail is not labelText then
                                set detailText to maybeDetail
                                exit repeat
                            end if
                        end try
                    end repeat

                    try
                        set selectedState to value of attribute "AXSelected" of e as text
                    on error
                        set selectedState to "false"
                    end try
                    set rp to position of e
                    set rs to size of e
                    set scrollRect to ""
                    set parentElem to e
                    repeat 12 times
                        try
                            set parentElem to parent of parentElem
                            if parentElem is missing value then exit repeat
                            if role of parentElem is "AXScrollArea" then
                                set sp to position of parentElem
                                set ss to size of parentElem
                                set scrollRect to (item 1 of sp) & tab & (item 2 of sp) & tab & (item 1 of ss) & tab & (item 2 of ss)
                                exit repeat
                            end if
                        on error
                            exit repeat
                        end try
                    end repeat

                    set end of resultItems to "ROW" & tab & rowIndex & tab & my cleanText(labelText) & tab & my cleanText(detailText) & tab & selectedState & tab & (item 1 of rp) & tab & (item 2 of rp) & tab & (item 1 of rs) & tab & (item 2 of rs) & tab & scrollRect
                end if
            end try
        end repeat

        set AppleScript's text item delimiters to linefeed
        return resultItems as text
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    raw = _run_osascript(script, timeout=timeout)
    snapshot: Dict[str, Any] = {"current_notebook": "", "scrolls": [], "rows": []}
    if not raw:
        return snapshot

    for line in raw.splitlines():
        parts = line.split("\t")
        if not parts:
            continue
        kind = parts[0]
        if kind == "NOTEBOOK" and len(parts) >= 2:
            snapshot["current_notebook"] = _extract_current_notebook_name(parts[1])
            continue
        if kind == "SCROLL" and len(parts) >= 6:
            try:
                x, y, w, h = map(int, parts[1:5])
                value = float(parts[5]) if parts[5] not in ("", "missing value") else None
            except Exception:
                continue
            snapshot["scrolls"].append(
                {
                    "rect": _rect_key(x, y, w, h),
                    "value": value,
                }
            )
            continue
        if kind == "ROW" and len(parts) >= 9:
            try:
                order = int(parts[1])
                text = _clean_field(parts[2])
                detail = _clean_field(parts[3]) if len(parts) >= 4 else ""
                selected = (parts[4].strip().lower() == "true") if len(parts) >= 5 else False
                x, y, w, h = map(int, parts[5:9])
                if len(parts) >= 13 and all(parts[idx].strip() for idx in range(9, 13)):
                    sx, sy, sw, sh = map(int, parts[9:13])
                    scroll_rect = _rect_key(sx, sy, sw, sh)
                else:
                    scroll_rect = None
            except Exception:
                continue
            if not text:
                continue
            snapshot["rows"].append(
                {
                    "order": order,
                    "text": text,
                    "detail": detail,
                    "selected": selected,
                    "rect": _rect_key(x, y, w, h),
                    "scroll_rect": scroll_rect,
                }
            )
    scrolls = snapshot.get("scrolls", [])
    for row in snapshot.get("rows", []):
        if row.get("scroll_rect"):
            continue
        x, y, w, h = row["rect"]
        center_x = x + (w / 2.0)
        center_y = y + (h / 2.0)
        matches = []
        for scroll in scrolls:
            sx, sy, sw, sh = scroll["rect"]
            if sx <= center_x <= sx + sw and sy <= center_y <= sy + sh:
                matches.append(scroll["rect"])
        if matches:
            row["scroll_rect"] = sorted(matches, key=lambda rect: rect[2] * rect[3])[0]
    return snapshot


def collect_onenote_rows(window: MacWindow) -> List[MacRow]:
    snapshot = collect_onenote_snapshot(window)
    rows: List[MacRow] = []
    for item in snapshot.get("rows", []):
        x, y, w, h = item["rect"]
        scroll_rect = item.get("scroll_rect") or item["rect"]
        sx, sy, sw, sh = scroll_rect
        rows.append(
            MacRow(
                window=window,
                text=item["text"],
                detail=item.get("detail") or "",
                selected=bool(item.get("selected")),
                rect=MacRect(x, y, x + w, y + h),
                scroll_rect=MacRect(sx, sy, sx + sw, sy + sh),
                order=int(item["order"]),
            )
        )
    return rows


def pick_selected_row(window: MacWindow, prefer_leftmost: bool = True) -> Optional[MacRow]:
    selected = [row for row in collect_onenote_rows(window) if row.selected]
    if not selected:
        return None
    return sorted(
        selected,
        key=lambda row: (
            row.scroll_rect.left if prefer_leftmost else -row.scroll_rect.left,
            row.order,
        ),
    )[0]


def pick_selected_page_row(window: MacWindow) -> Optional[MacRow]:
    rows = collect_onenote_rows(window)
    selected = [row for row in rows if row.selected]
    if not selected:
        return None
    scroll_lefts = sorted({row.scroll_rect.left for row in rows})
    if len(scroll_lefts) < 2:
        return None
    page_scroll_left = scroll_lefts[-1]
    page_rows = [row for row in selected if row.scroll_rect.left == page_scroll_left]
    if not page_rows:
        return None
    return sorted(page_rows, key=lambda row: row.order)[0]


def select_row_by_text(
    window: MacWindow,
    text: str,
    preferred_scroll_left: Optional[int] = None,
) -> bool:
    return _select_outline_row_by_text(
        window,
        text,
        preferred_scroll_left=preferred_scroll_left,
    )


def _select_outline_row_by_text(
    window: MacWindow,
    text: str,
    preferred_scroll_left: Optional[int] = None,
    *,
    prefer_rightmost: bool = False,
) -> bool:
    wanted_text_nfc = _clean_field(text)
    wanted_text = unicodedata.normalize("NFD", wanted_text_nfc)
    if not wanted_text:
        return False

    locator = _applescript_window_locator(window.process_id(), window.window_text())
    wanted_left = -1 if preferred_scroll_left is None else int(preferred_scroll_left)
    rightmost_flag = "true" if prefer_rightmost else "false"
    script = locator + f'''
        set wantedText to {_quote_applescript_text(wanted_text)}
        set wantedTextNFC to {_quote_applescript_text(wanted_text_nfc)}
        set wantedScrollLeft to {wanted_left}
        set preferRightmost to {rightmost_flag}
        set bestRow to missing value
        set bestScore to 999999
        if preferRightmost then set bestScore to -999999

        try
            set targetSplitGroup to first UI element of targetWindow whose role is "AXSplitGroup"
            set leftGroup to first UI element of targetSplitGroup whose role is "AXGroup"
            set nestedSplitGroup to first UI element of leftGroup whose role is "AXSplitGroup"
            set groupsToScan to UI elements of nestedSplitGroup
        on error
            set groupsToScan to UI elements of targetWindow
        end try

        repeat with g in groupsToScan
            try
                set gRole to (role of g) as text
                set shouldScanGroup to false
                if gRole is "AXGroup" then
                    set shouldScanGroup to true
                end if
                if gRole is "AXScrollArea" then
                    set shouldScanGroup to true
                end if
                if shouldScanGroup then
                    set scrollAreas to {{}}
                    if gRole is "AXScrollArea" then
                        set end of scrollAreas to g
                    else
                        repeat with maybeScrollArea in (UI elements of g)
                            try
                                if ((role of maybeScrollArea) as text) is "AXScrollArea" then
                                    set end of scrollAreas to maybeScrollArea
                                end if
                            end try
                        end repeat
                    end if

                    repeat with sa in scrollAreas
                        try
                            set sp to position of sa
                            repeat with outlineElem in (UI elements of sa)
                                try
                                    set outlineRole to (role of outlineElem) as text
                                    set shouldScanOutline to false
                                    if outlineRole is "AXOutline" then
                                        set shouldScanOutline to true
                                    end if
                                    if outlineRole is "AXTable" then
                                        set shouldScanOutline to true
                                    end if
                                    if shouldScanOutline then
                                        set rowCount to count of rows of outlineElem
                                        repeat with rowIndex from 1 to rowCount
                                            try
                                                set targetRow to row rowIndex of outlineElem
                                                set labelText to my rowLabel(targetRow)
                                                set labelMatches to false
                                                if labelText is wantedText then
                                                    set labelMatches to true
                                                end if
                                                if labelText is wantedTextNFC then
                                                    set labelMatches to true
                                                end if
                                                if labelText contains wantedText then
                                                    set labelMatches to true
                                                end if
                                                if labelText contains wantedTextNFC then
                                                    set labelMatches to true
                                                end if
                                                if wantedText contains labelText then
                                                    set labelMatches to true
                                                end if
                                                if wantedTextNFC contains labelText then
                                                    set labelMatches to true
                                                end if
                                                if labelMatches then
                                                    if wantedScrollLeft is -1 then
                                                        if preferRightmost then
                                                            set rowLeft to item 1 of sp
                                                            if rowLeft > bestScore then
                                                                set bestScore to rowLeft
                                                                set bestRow to targetRow
                                                            end if
                                                        else
                                                            if my selectRow(targetRow) then return "OK"
                                                        end if
                                                    else
                                                        set rowLeft to item 1 of sp
                                                        set distanceValue to rowLeft - wantedScrollLeft
                                                        if distanceValue < 0 then set distanceValue to -distanceValue
                                                        if distanceValue < bestScore then
                                                            set bestScore to distanceValue
                                                            set bestRow to targetRow
                                                        end if
                                                    end if
                                                end if
                                            end try
                                        end repeat
                                    end if
                                end try
                            end repeat
                        end try
                    end repeat
                end if
            end try
        end repeat

        if bestRow is not missing value then
            if my selectRow(bestRow) then return "OK"
        end if
        return ""
end tell

on selectRow(targetRow)
    try
        set value of attribute "AXSelected" of targetRow to true
        return true
    end try
    try
        perform action "AXPress" of targetRow
        return true
    end try
    try
        click targetRow
        return true
    end try
    try
        set rowPosition to position of targetRow
        set rowSize to size of targetRow
        set clickX to (item 1 of rowPosition) + 20
        set clickY to (item 2 of rowPosition) + ((item 2 of rowSize) / 2)
        click at {{clickX, clickY}}
        return true
    end try
    return false
end selectRow

on rowLabel(targetRow)
    set labelText to ""
    try
        set firstCell to UI element 1 of targetRow
        try
            set labelText to my cleanText(value of attribute "AXDescription" of firstCell as text)
        end try
        if labelText is "" then
            try
                set labelText to my cleanText(name of firstCell as text)
            end try
        end if
        if labelText is "" then
            try
                set labelText to my cleanText(value of firstCell as text)
            end try
        end if
    end try
    if labelText is "" then
        repeat with c in (UI elements of targetRow)
            try
                if role of c is "AXStaticText" then
                    set labelText to my cleanText(value of c as text)
                    if labelText is not "" then exit repeat
                end if
            end try
            try
                if labelText is "" then
                    set labelText to my cleanText(value of attribute "AXDescription" of c as text)
                    if labelText is not "" then exit repeat
                end if
            end try
        end repeat
    end if
    return labelText
end rowLabel

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    return _run_osascript(script, timeout=8).strip() == "OK"


def select_page_row_by_text(window: MacWindow, text: str) -> bool:
    return _select_outline_row_by_text(
        window,
        text,
        prefer_rightmost=True,
    )


def current_outline_context(window: MacWindow) -> Dict[str, str]:
    locator = _applescript_window_locator(window.process_id(), window.window_text())
    script = locator + r'''
        set notebookName to ""
        set selectedRows to {}
        set allElems to entire contents of targetWindow

        repeat with e in allElems
            try
                if role of e is "AXButton" and notebookName is "" then
                    set candidateText to ""
                    try
                        set candidateText to description of e as text
                    end try
                    if candidateText does not contain "(현재 전자필기장)" and candidateText does not contain "(현재 전자 필기장)" then
                        try
                            set candidateText to value of e as text
                        end try
                    end if
                    if candidateText does not contain "(현재 전자필기장)" and candidateText does not contain "(현재 전자 필기장)" then
                        try
                            set candidateText to title of e as text
                        end try
                    end if
                    set candidateText to my cleanText(candidateText)
                    if candidateText contains "(현재 전자필기장)" or candidateText contains "(현재 전자 필기장)" then
                        set notebookName to candidateText
                    end if
                end if

                if role of e is "AXRow" then
                    try
                        set selectedState to value of attribute "AXSelected" of e as boolean
                    on error
                        set selectedState to false
                    end try
                    if selectedState is true then
                        set labelText to ""
                        set childElems to entire contents of e
                        repeat with c in childElems
                            try
                                if role of c is "AXStaticText" then
                                    set labelText to my cleanText(value of c as text)
                                    exit repeat
                                end if
                            end try
                        end repeat
                        if labelText is "" then
                            repeat with c in childElems
                                try
                                    set maybeText to my cleanText(description of c as text)
                                    if maybeText is not "" then
                                        set labelText to maybeText
                                        exit repeat
                                    end if
                                end try
                                try
                                    set maybeText to my cleanText(title of c as text)
                                    if maybeText is not "" then
                                        set labelText to maybeText
                                        exit repeat
                                    end if
                                end try
                            end repeat
                        end if
                        if labelText is not "" then
                            set rp to position of e
                            set end of selectedRows to ((item 1 of rp) as text) & tab & labelText
                        end if
                    end if
                end if
            end try
        end repeat

        set resultItems to {}
        if notebookName is not "" then
            set end of resultItems to "NOTEBOOK" & tab & my cleanText(notebookName)
        end if
        repeat with rowInfo in selectedRows
            set end of resultItems to "ROW" & tab & rowInfo
        end repeat

        set AppleScript's text item delimiters to linefeed
        return resultItems as text
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    raw = _run_osascript(script, timeout=12)
    notebook_name = str(window.window_text() or "").strip()
    selected_rows: List[Tuple[int, str]] = []

    for line in (raw or "").splitlines():
        parts = line.split("\t")
        if not parts:
            continue
        if parts[0] == "NOTEBOOK" and len(parts) >= 2:
            notebook_name = _extract_current_notebook_name(parts[1])
            continue
        if parts[0] == "ROW" and len(parts) >= 3:
            try:
                selected_rows.append((int(parts[1]), _clean_field(parts[2])))
            except Exception:
                continue

    selected_rows = [(x, text) for x, text in selected_rows if text]
    selected_rows.sort(key=lambda item: item[0])
    section_name = selected_rows[0][1] if selected_rows else ""
    page_name = selected_rows[-1][1] if len(selected_rows) >= 2 else ""

    return {
        "notebook": str(notebook_name or "").strip(),
        "section": str(section_name or "").strip(),
        "page": str(page_name or "").strip(),
    }


def center_selected_row(
    window: MacWindow,
    prefer_leftmost: bool = True,
    target_text: str = "",
) -> Tuple[bool, Optional[str]]:
    for _ in range(3):
        snapshot = collect_onenote_snapshot(window)
        rows = snapshot.get("rows", [])
        selected_rows = [row for row in rows if row.get("selected")]
        if not selected_rows:
            return False, None

        wanted_key = _normalize_text(target_text)
        if wanted_key:
            target_rows = [
                row
                for row in selected_rows
                if _normalize_text(row.get("text") or "") == wanted_key
            ]
            if target_rows:
                selected_rows = target_rows

        target = sorted(
            selected_rows,
            key=lambda row: (
                row["scroll_rect"][0] if prefer_leftmost else -row["scroll_rect"][0],
                row["order"],
            ),
        )[0]

        sx, sy, sw, sh = target["scroll_rect"]
        x, y, w, h = target["rect"]
        scroll = next(
            (
                item for item in snapshot.get("scrolls", [])
                if tuple(item.get("rect") or ()) == (sx, sy, sw, sh)
            ),
            None,
        )
        row_center = y + (h / 2.0)
        scroll_center = sy + (sh / 2.0)
        delta = row_center - scroll_center
        if abs(delta) <= 12:
            return True, target["text"]

        if not scroll or scroll.get("value") is None:
            return True, target["text"]

        group_rows = [row for row in rows if tuple(row.get("scroll_rect") or ()) == (sx, sy, sw, sh)]
        if not group_rows:
            return False, target["text"]

        min_top = min(row["rect"][1] for row in group_rows)
        max_bottom = max((row["rect"][1] + row["rect"][3]) for row in group_rows)
        content_height = max_bottom - min_top
        scrollable_height = max(1.0, float(content_height - sh))
        if scrollable_height <= 1.0:
            return True, target["text"]

        current_value = float(scroll["value"])
        target_value = max(0.0, min(1.0, current_value + (delta / scrollable_height)))
        if abs(target_value - current_value) < 0.002:
            return True, target["text"]

        locator = _applescript_window_locator(window.process_id(), window.window_text())
        script = locator + f'''
            set targetScrollLeft to {sx}
            set targetScrollTop to {sy}
            set targetValue to {target_value}
            set allElems to entire contents of targetWindow
            repeat with e in allElems
                try
                    if role of e is "AXScrollArea" then
                        set pp to position of e
                        if (item 1 of pp) is targetScrollLeft and (item 2 of pp) is targetScrollTop then
                            repeat with b in (every UI element of e)
                                try
                                    if role of b is "AXScrollBar" then
                                        try
                                            set bo to orientation of b as text
                                        on error
                                            set bo to ""
                                        end try
                                        if bo is "AXVerticalOrientation" then
                                            set value of b to targetValue
                                            return "OK"
                                        end if
                                    end if
                                end try
                            end repeat
                        end if
                    end if
                end try
            end repeat
            return ""
end tell
'''
        if _run_osascript(script, timeout=15).strip() != "OK":
            return False, target["text"]
        time.sleep(0.12)

    final_row = pick_selected_row(window, prefer_leftmost=prefer_leftmost)
    return (final_row is not None), (final_row.text if final_row else None)


def list_current_notebook_targets(window: MacWindow) -> List[Dict[str, str]]:
    snapshot = collect_onenote_snapshot(window)
    notebook_name = snapshot.get("current_notebook") or window.window_text()
    rows = snapshot.get("rows", [])
    if not rows:
        return [
            {
                "kind": "notebook",
                "name": f"전자필기장 - {notebook_name}",
                "path": notebook_name,
                "notebook": notebook_name,
                "section_group": "",
                "section": "",
                "section_group_id": "",
                "section_id": "",
            }
        ]

    leftmost_scroll = min(row["scroll_rect"][0] for row in rows)
    sections = []
    seen = set()
    for row in rows:
        if row["scroll_rect"][0] != leftmost_scroll:
            continue
        section_name = str(row.get("text") or "").strip()
        if not section_name:
            continue
        key = _normalize_text(section_name)
        if key in seen:
            continue
        seen.add(key)
        sections.append(
            {
                "kind": "section",
                "name": f"섹션 - {section_name}",
                "path": f"{notebook_name} > {section_name}",
                "notebook": notebook_name,
                "section_group": "",
                "section": section_name,
                "section_group_id": "",
                "section_id": "",
            }
        )

    base = [
        {
            "kind": "notebook",
            "name": f"전자필기장 - {notebook_name}",
            "path": notebook_name,
            "notebook": notebook_name,
            "section_group": "",
            "section": "",
            "section_group_id": "",
            "section_id": "",
        }
    ]
    return base + sections


def _read_open_notebook_names_from_plist() -> List[str]:
    if not _MAC_ONENOTE_NOTEBOOKS_PLIST.is_file():
        return []
    try:
        import plistlib

        with _MAC_ONENOTE_NOTEBOOKS_PLIST.open("rb") as f:
            data = plistlib.load(f)
    except Exception:
        return []

    if not isinstance(data, list):
        return []

    names: List[str] = []
    seen = set()
    for item in data:
        if not isinstance(item, dict):
            continue
        if int(item.get("Type") or 0) != 1:
            continue
        name = _clean_field(str(item.get("Name") or ""))
        key = _normalize_text(name)
        if not key or key in seen:
            continue
        seen.add(key)
        names.append(name)
    return names


def _read_open_notebook_names_from_plist_with_timeout(
    timeout_sec: float = 1.5,
) -> Optional[List[str]]:
    box: Dict[str, Any] = {}
    done = threading.Event()

    def _runner() -> None:
        try:
            box["value"] = _read_open_notebook_names_from_plist()
        except Exception as exc:
            box["error"] = exc
        finally:
            done.set()

    worker = threading.Thread(target=_runner, daemon=True)
    worker.start()
    if not done.wait(max(0.2, float(timeout_sec or 0.0))):
        return None
    if "error" in box:
        return []
    value = box.get("value")
    return value if isinstance(value, list) else []


def _ax_window_roots_for_onenote(window: Optional[MacWindow]) -> List[c_void_p]:
    if not (IS_MACOS and _APP_SERVICES and window is not None):
        return []
    pid = window.process_id()
    if not pid:
        return []

    app_ref = c_void_p(_APP_SERVICES.AXUIElementCreateApplication(int(pid)))
    if not app_ref:
        return []

    roots: List[c_void_p] = []
    try:
        for attr_name in ("AXFocusedWindow", "AXMainWindow"):
            ref = _ax_element_attribute(app_ref, attr_name)
            if ref:
                roots.append(ref)
        roots.extend(_ax_array_attribute(app_ref, "AXWindows"))
    finally:
        _cf_release(app_ref)

    if not roots:
        return []

    title_hint = _clean_field(str(window.info.get("title") or ""))
    if not title_hint:
        try:
            title_hint = _clean_field(window.window_text())
        except Exception:
            title_hint = ""
    title_key = _normalize_text(title_hint)

    unique: List[c_void_p] = []
    seen_ptrs = set()
    for root in roots:
        ptr = int(root.value or 0)
        if not ptr:
            continue
        if ptr in seen_ptrs:
            _cf_release(root)
            continue
        seen_ptrs.add(ptr)
        unique.append(root)

    def _rank(root: c_void_p) -> Tuple[int, int, str]:
        ax_title = _ax_text_attribute(root, "AXTitle")
        key = _normalize_text(ax_title)
        if title_key and key == title_key:
            return (0, 0, key)
        if title_key and (title_key in key or key in title_key):
            return (0, 1, key)
        return (1, 0, key)

    return sorted(unique, key=_rank)


def _notebook_name_from_ax_label(raw_label: str) -> str:
    label = _clean_field(raw_label)
    if not label:
        return ""
    name = _extract_current_notebook_name(label)
    name = _clean_field(name)
    key = _normalize_text(name)
    if not key:
        return ""
    blocked_exact = {
        "전자 필기장",
        "전자필기장",
        "notebook",
        "notebooks",
        "open notebook",
        "close notebook",
        "열기",
        "닫기",
        "검색",
        "추가",
    }
    if key in blocked_exact:
        return ""
    blocked_fragments = (
        "sectiontab",
        "pagetab",
        "outline",
        "ax",
        "button",
        "scroll",
    )
    if any(fragment in key for fragment in blocked_fragments):
        return ""
    if len(name) > 120:
        return ""
    return name


def _ax_candidate_labels(element: c_void_p, depth: int = 0) -> List[str]:
    if not element or depth > 4:
        return []

    labels: List[str] = []
    for attr_name in ("AXDescription", "AXTitle", "AXValue"):
        label = _ax_text_attribute(element, attr_name)
        if label:
            labels.append(label)

    children = _ax_array_attribute(element, "AXChildren")
    try:
        preferred: List[str] = []
        fallback: List[str] = []
        for child in children:
            role = _ax_text_attribute(child, "AXRole")
            child_labels = _ax_candidate_labels(child, depth + 1)
            if role in {"AXStaticText", "AXTextField", "AXCell"}:
                preferred.extend(child_labels)
            else:
                fallback.extend(child_labels)
        labels.extend(preferred)
        labels.extend(fallback)
    finally:
        _release_ax_refs(children)

    deduped: List[str] = []
    seen = set()
    for label in labels:
        key = _normalize_text(label)
        if not key or key in seen:
            continue
        seen.add(key)
        deduped.append(label)
    return deduped


def _ax_row_notebook_name(row: c_void_p) -> str:
    for label in _ax_candidate_labels(row):
        name = _notebook_name_from_ax_label(label)
        if name:
            return name
    return ""


def _ax_outline_notebook_names(outline: c_void_p) -> List[str]:
    rows = _ax_array_attribute(outline, "AXRows")
    if not rows:
        rows = _ax_array_attribute(outline, "AXChildren")

    names: List[str] = []
    seen = set()
    try:
        for row in rows:
            name = _ax_row_notebook_name(row)
            key = _normalize_text(name)
            if not key or key in seen:
                continue
            seen.add(key)
            names.append(name)
    finally:
        _release_ax_refs(rows)
    return names


def _collect_ax_outline_name_groups(root: c_void_p) -> List[Dict[str, Any]]:
    groups: List[Dict[str, Any]] = []
    node_count = 0
    deadline = time.monotonic() + 4.0

    def _visit(element: c_void_p, depth: int) -> None:
        nonlocal node_count
        if not element or depth > 18 or node_count >= 1800:
            return
        if time.monotonic() > deadline:
            return
        if not int(element.value or 0):
            return
        node_count += 1

        role = _ax_text_attribute(element, "AXRole")
        if role == "AXOutline":
            names = _ax_outline_notebook_names(element)
            if names:
                groups.append(
                    {
                        "names": names,
                        "order": len(groups),
                        "depth": depth,
                    }
                )

        children = _ax_array_attribute(element, "AXChildren")
        try:
            for child in children:
                _visit(child, depth + 1)
        finally:
            _release_ax_refs(children)

    _visit(root, 0)
    return groups


def _read_open_notebook_names_from_ax(window: Optional[MacWindow]) -> List[str]:
    global _MAC_LAST_AX_NOTEBOOK_DEBUG
    debug: Dict[str, Any] = {
        "trusted": macos_accessibility_is_trusted(),
        "pid": 0,
        "title": "",
        "roots": 0,
        "groups": 0,
        "best_count": 0,
        "reason": "",
    }
    if not (IS_MACOS and window is not None):
        debug["reason"] = "not_macos_or_no_window"
        _MAC_LAST_AX_NOTEBOOK_DEBUG = debug
        return []

    current_title = _clean_field(str(window.info.get("title") or ""))
    debug["pid"] = window.process_id()
    if not current_title:
        try:
            current_title = _clean_field(window.window_text())
        except Exception:
            current_title = ""
    debug["title"] = current_title
    current_key = _normalize_text(current_title)

    roots = _ax_window_roots_for_onenote(window)
    debug["roots"] = len(roots)
    if not roots:
        debug["reason"] = "no_ax_roots"
        _MAC_LAST_AX_NOTEBOOK_DEBUG = debug
        return []

    groups: List[Dict[str, Any]] = []
    try:
        for root in roots[:3]:
            groups.extend(_collect_ax_outline_name_groups(root))
    finally:
        _release_ax_refs(roots)

    debug["groups"] = len(groups)
    if not groups:
        debug["reason"] = "no_ax_outline_groups"
        _MAC_LAST_AX_NOTEBOOK_DEBUG = debug
        return []

    def _score(group: Dict[str, Any]) -> Tuple[int, int, int]:
        names = [str(name or "") for name in group.get("names") or []]
        keys = {_normalize_text(name) for name in names}
        score = min(len(names), 80)
        if current_key and current_key in keys:
            score += 300
        if len(names) >= 3:
            score += 20
        score -= int(group.get("order") or 0) * 3
        score -= int(group.get("depth") or 0)
        return (score, len(names), -int(group.get("order") or 0))

    best = max(groups, key=_score)
    best_names = [str(name or "").strip() for name in best.get("names") or []]
    debug["best_count"] = len(best_names)
    if current_key and current_key not in {_normalize_text(name) for name in best_names}:
        debug["reason"] = "best_group_missing_current_title"
        _MAC_LAST_AX_NOTEBOOK_DEBUG = debug
        return []
    if len(best_names) < 2 and not current_key:
        debug["reason"] = "best_group_too_small"
        _MAC_LAST_AX_NOTEBOOK_DEBUG = debug
        return []
    debug["reason"] = "ok"
    _MAC_LAST_AX_NOTEBOOK_DEBUG = debug
    return best_names


def _read_open_notebook_names_from_sidebar(window: Optional[MacWindow]) -> List[str]:
    if window is None:
        return []
    try:
        sidebar_ready, opened_sidebar = _ensure_notebook_sidebar(window)
    except Exception:
        return []
    if not sidebar_ready:
        return []

    script = _applescript_window_locator(window.process_id(), window.window_text()) + r'''
        set resultItems to {}
        set targetSplitGroup to first UI element of targetWindow whose role is "AXSplitGroup"
        set leftGroup to first UI element of targetSplitGroup whose role is "AXGroup"
        set nestedSplitGroup to first UI element of leftGroup whose role is "AXSplitGroup"
        set notebookGroup to first UI element of nestedSplitGroup whose role is "AXGroup"
        set notebookScrollArea to first UI element of notebookGroup whose role is "AXScrollArea"
        set targetOutline to first UI element of notebookScrollArea whose role is "AXOutline"

        set rowCount to count of rows of targetOutline
        repeat with rowIndex from 1 to rowCount
            try
                set targetRow to row rowIndex of targetOutline
                set firstCell to UI element 1 of targetRow
                set labelText to ""
                try
                    set labelText to my cleanText(value of attribute "AXDescription" of firstCell as text)
                end try
                if labelText is "" then
                    try
                        set labelText to my cleanText(name of firstCell as text)
                    end try
                end if
                if labelText is "" then
                    try
                        set labelText to my cleanText(value of firstCell as text)
                    end try
                end if
                if labelText is not "" then
                    set end of resultItems to labelText
                end if
            end try
        end repeat
        set AppleScript's text item delimiters to linefeed
        return resultItems as text
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    try:
        raw = _run_osascript(script, timeout=6)
    except Exception:
        raw = ""
    finally:
        _restore_notebook_sidebar(window, opened_sidebar)

    names: List[str] = []
    seen = set()
    for line in (raw or "").splitlines():
        name = _clean_field(line)
        key = _normalize_text(name)
        if not key or key in seen:
            continue
        seen.add(key)
        names.append(name)
    return names


def _recent_notebook_records_from_cache() -> List[Dict[str, Any]]:
    if not _MAC_ONENOTE_RESOURCEINFOCACHE_JSON.is_file():
        return []
    try:
        payload = json.loads(_MAC_ONENOTE_RESOURCEINFOCACHE_JSON.read_text(encoding="utf-8"))
    except Exception:
        return []

    entries = payload.get("ResourceInfoCache") or []
    if not isinstance(entries, list):
        return []

    records: List[Dict[str, Any]] = []
    seen = set()
    for item in sorted(
        (entry for entry in entries if isinstance(entry, dict)),
        key=lambda entry: int(entry.get("LastAccessedAt") or 0),
        reverse=True,
    ):
        raw_url = str(item.get("Url") or "").strip()
        if not raw_url:
            continue
        parsed = urlparse(raw_url)
        if parsed.scheme.lower() not in {"http", "https"}:
            continue
        if parsed.netloc.lower() != "d.docs.live.net":
            continue
        notebook_name = _clean_field(str(item.get("Title") or item.get("title") or ""))
        if not notebook_name:
            notebook_name = _clean_field(unquote(parsed.path.rstrip("/").split("/")[-1]))
        key = _normalize_text(notebook_name)
        if not key or key in seen:
            continue
        seen.add(key)
        records.append(
            {
                "name": notebook_name,
                "url": raw_url,
                "last_accessed_at": int(item.get("LastAccessedAt") or 0),
            }
        )
    return records


def _recent_notebook_records_from_cache_with_timeout(
    timeout_sec: float = 0.8,
) -> Optional[List[Dict[str, Any]]]:
    global _MAC_RECENT_CACHE_TIMED_OUT
    if _MAC_RECENT_CACHE_TIMED_OUT:
        return None

    box: Dict[str, Any] = {}
    done = threading.Event()

    def _runner() -> None:
        try:
            box["value"] = _recent_notebook_records_from_cache()
        except Exception as exc:
            box["error"] = exc
        finally:
            done.set()

    threading.Thread(
        target=_runner,
        name="onenote-mac-recent-cache",
        daemon=True,
    ).start()
    if not done.wait(timeout_sec):
        _MAC_RECENT_CACHE_TIMED_OUT = True
        return None
    if "error" in box:
        raise box["error"]
    return list(box.get("value") or [])


def _onenote_protocol_url_from_web_url(web_url: str) -> str:
    parsed = urlparse(str(web_url or "").strip())
    if not parsed.scheme or not parsed.netloc:
        return ""
    path = quote(unquote(parsed.path or ""), safe="/-_.~")
    query = f"?{parsed.query}" if parsed.query else ""
    fragment = f"#{parsed.fragment}" if parsed.fragment else ""
    return f"onenote:{parsed.scheme}://{parsed.netloc}{path}{query}{fragment}"


def recent_notebook_records(window: Optional[MacWindow] = None) -> List[Dict[str, Any]]:
    try:
        records = _recent_notebook_records_from_cache_with_timeout()
    except Exception:
        records = []
    if records is None:
        print("[DBG][MAC][RECENT_CACHE] timeout; fallback=dialog")
        records = []
    if records:
        return [dict(record) for record in records]

    if window is None:
        return []

    names = recent_notebook_names(window)
    return [
        {"name": name, "url": "", "last_accessed_at": 0}
        for name in names
        if str(name or "").strip()
    ]


def open_recent_notebook_record(
    window: Optional[MacWindow],
    record: Dict[str, Any],
    wait_for_visible: bool = True,
) -> bool:
    name = _clean_field(str((record or {}).get("name") or ""))
    web_url = str((record or {}).get("url") or "").strip()
    protocol_url = _onenote_protocol_url_from_web_url(web_url)

    if protocol_url:
        try:
            open_url_in_system(protocol_url)
            return True
        except Exception:
            pass

    if window is None or not name:
        return False
    return open_recent_notebook_by_name(
        window,
        name,
        wait_for_visible=wait_for_visible,
    )


def _is_notebook_sidebar_snapshot(snapshot: Dict[str, Any]) -> bool:
    rows = snapshot.get("rows") or []
    if not rows:
        return False
    scroll_rects = {
        tuple(row.get("scroll_rect") or ())
        for row in rows
        if row.get("scroll_rect")
    }
    return len(scroll_rects) == 1


def _press_window_button(
    window: MacWindow,
    *,
    help_contains: Optional[List[str]] = None,
    value_equals: Optional[List[str]] = None,
    timeout: int = 15,
) -> bool:
    locator = _applescript_window_locator(window.process_id(), window.window_text())
    help_matchers = [str(item or "").strip() for item in (help_contains or []) if str(item or "").strip()]
    value_matchers = [str(item or "").strip() for item in (value_equals or []) if str(item or "").strip()]
    help_tokens = ", ".join(_quote_applescript_text(item) for item in help_matchers) or ""
    value_tokens = ", ".join(_quote_applescript_text(item) for item in value_matchers) or ""
    script = locator + f'''
        set helpMatchers to {{{help_tokens}}}
        set valueMatchers to {{{value_tokens}}}
        set allElems to entire contents of targetWindow
        repeat with e in allElems
            try
                set roleText to role of e as text
            on error
                set roleText to ""
            end try
            if roleText is "AXButton" or roleText is "AXMenuButton" then
                set titleText to ""
                set nameText to ""
                set descText to ""
                set valueText to ""
                set helpText to ""
                try
                    set titleText to my cleanText(title of e as text)
                end try
                try
                    set nameText to my cleanText(name of e as text)
                end try
                try
                    set descText to my cleanText(description of e as text)
                end try
                try
                    set valueText to my cleanText(value of e as text)
                end try
                try
                    set helpText to my cleanText(help of e as text)
                end try

                set shouldPress to false
                repeat with wantedHelp in helpMatchers
                    if wantedHelp is not "" and helpText contains wantedHelp then
                        set shouldPress to true
                        exit repeat
                    end if
                end repeat

                if shouldPress is false then
                    repeat with wantedValue in valueMatchers
                        if wantedValue is not "" and (titleText is wantedValue or nameText is wantedValue or descText is wantedValue or valueText is wantedValue) then
                            set shouldPress to true
                            exit repeat
                        end if
                    end repeat
                end if

                if shouldPress then
                    try
                        perform action "AXPress" of e
                    on error
                        click e
                    end try
                    return "OK"
                end if
            end if
        end repeat
        return ""
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    return _run_osascript(script, timeout=timeout).strip() == "OK"


def _has_notebook_sidebar_ui(window: MacWindow) -> bool:
    script = _applescript_window_locator(window.process_id(), window.window_text()) + r'''
        try
            set targetSplitGroup to first UI element of targetWindow whose role is "AXSplitGroup"
            set leftGroup to first UI element of targetSplitGroup whose role is "AXGroup"
            set nestedSplitGroup to first UI element of leftGroup whose role is "AXSplitGroup"
            set notebookGroup to first UI element of nestedSplitGroup whose role is "AXGroup"
            set notebookScrollArea to first UI element of notebookGroup whose role is "AXScrollArea"
            set targetOutline to first UI element of notebookScrollArea whose role is "AXOutline"
            if targetOutline is not missing value then return "OK"
        end try
        return ""
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    try:
        return _run_osascript(script, timeout=6).strip() == "OK"
    except Exception:
        return False


def _ensure_notebook_sidebar(window: MacWindow) -> Tuple[bool, bool]:
    if _has_notebook_sidebar_ui(window):
        return True, False
    if not _press_window_button(
        window,
        help_contains=["전자 필기장 보기, 만들기 또는 열기"],
        timeout=15,
    ):
        return False, False
    for _ in range(8):
        time.sleep(0.12)
        if _has_notebook_sidebar_ui(window):
            return True, True
    return False, True


def _restore_notebook_sidebar(window: MacWindow, opened_by_us: bool) -> None:
    if not opened_by_us:
        return
    try:
        _press_window_button(
            window,
            help_contains=["전자 필기장 목록 숨기기"],
            value_equals=["전자 필기장"],
            timeout=10,
        )
    except Exception:
        pass


def _recent_notebook_dialog_locator(window: MacWindow) -> str:
    pid = int(window.process_id())
    window_number = int(window.info.get("window_number") or 0)
    return f'''
tell application "System Events"
    set targetProcess to first application process whose unix id is {pid}
    set targetWindow to missing value
    set targetWindowNumber to {window_number}
    if targetWindowNumber > 0 then
        repeat with w in windows of targetProcess
            try
                set candidateNumber to value of attribute "AXWindowNumber" of w
                if (candidateNumber as integer) is targetWindowNumber then
                    set targetWindow to w
                    exit repeat
                end if
            end try
        end repeat
    end if
    if targetWindow is missing value then
        repeat with w in windows of targetProcess
            try
                set winName to name of w as text
            on error
                set winName to ""
            end try
            if winName contains "최근 전자 필기장" or winName contains "새 전자 필기장" then
                set targetWindow to w
                exit repeat
            end if
        end repeat
    end if
    if targetWindow is missing value then
        repeat with w in windows of targetProcess
            try
                if exists (first UI element of w whose role is "AXTextField") then
                    try
                        if exists (first button of w whose name is "열기") then
                            set targetWindow to w
                            exit repeat
                        end if
                    end try
                    try
                        if exists (first button of w whose name is "Open") then
                            set targetWindow to w
                            exit repeat
                        end if
                    end try
                    try
                        if exists (first radio button of w whose name is "최근") then
                            set targetWindow to w
                            exit repeat
                        end if
                    end try
                    try
                        if exists (first radio button of w whose name is "Recent") then
                            set targetWindow to w
                            exit repeat
                        end if
                    end try
                end if
            end try
        end repeat
    end if
    if targetWindow is missing value then
        repeat with w in windows of targetProcess
            try
                if role of w is "AXWindow" or role of w is "AXStandardWindow" then
                    set targetWindow to w
                    exit repeat
                end if
            end try
        end repeat
    end if
    if targetWindow is missing value then error "최근 전자필기장 창을 찾지 못했습니다."
'''


def _open_recent_notebooks_dialog(window: MacWindow) -> bool:
    ready, _ = _open_recent_notebooks_dialog_with_state(window)
    return ready


def _is_recent_notebooks_dialog_open(window: MacWindow) -> bool:
    try:
        _run_osascript(
            _recent_notebook_dialog_locator(window) + '\nreturn "OK"\nend tell\n',
            timeout=5,
        )
        return True
    except Exception:
        return False


def _open_recent_notebooks_dialog_with_state(window: MacWindow) -> Tuple[bool, bool]:
    if _recent_notebook_dialog_row_count(window) > 0:
        return True, False
    sidebar_ready, opened_sidebar = _ensure_notebook_sidebar(window)
    if not sidebar_ready:
        return False, False
    try:
        if not _press_window_button(
            window,
            help_contains=["전자 필기장 만들기 또는 열기"],
            value_equals=["전자 필기장 추가"],
            timeout=15,
        ):
            _restore_notebook_sidebar(window, opened_sidebar)
            return False, False
        deadline = time.monotonic() + 4.0
        while time.monotonic() < deadline:
            time.sleep(0.12)
            try:
                _run_osascript(
                    _recent_notebook_dialog_locator(window) + '\nreturn "OK"\nend tell\n',
                    timeout=5,
                )
                return True, opened_sidebar
            except Exception:
                continue
        _restore_notebook_sidebar(window, opened_sidebar)
        return False, False
    except Exception:
        _restore_notebook_sidebar(window, opened_sidebar)
        raise


def dismiss_recent_notebooks_dialog(window: MacWindow) -> bool:
    script = _recent_notebook_dialog_locator(window) + r'''
    try
        set frontmost of targetProcess to true
    end try
    delay 0.05
    repeat with e in UI elements of targetWindow
        try
            if role of e is "AXButton" then
                set buttonText to ""
                try
                    set buttonText to my cleanText(title of e as text)
                end try
                if buttonText is "" then
                    try
                        set buttonText to my cleanText(name of e as text)
                    end try
                end if
                if buttonText is "" then
                    try
                        set buttonText to my cleanText(description of e as text)
                    end try
                end if
                if buttonText is "취소" then
                    try
                        perform action "AXPress" of e
                    on error
                        click e
                    end try
                    return "OK"
                end if
            end if
        end try
    end repeat
    key code 53
    return "OK"
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    try:
        ok = _run_osascript(script, timeout=5).strip() == "OK"
    except Exception:
        return False
    if ok:
        deadline = time.monotonic() + 2.0
        while time.monotonic() < deadline:
            time.sleep(0.12)
            if not _is_recent_notebooks_dialog_open(window):
                return True
    return not _is_recent_notebooks_dialog_open(window)


def _clear_recent_notebooks_dialog_search(window: MacWindow) -> bool:
    script = _recent_notebook_dialog_locator(window) + r'''
    set didClear to false
    repeat with e in UI elements of targetWindow
        try
            if role of e is "AXTextField" then
                set currentText to ""
                try
                    set currentText to my cleanText(value of e as text)
                end try
                if currentText is "" or currentText is "최근 전자 필기장 검색" then
                    return ""
                end if

                try
                    repeat with childElem in UI elements of e
                        try
                            if role of childElem is "AXButton" then
                                set buttonText to ""
                                try
                                    set buttonText to my cleanText(name of childElem as text)
                                end try
                                if buttonText is "" then
                                    try
                                        set buttonText to my cleanText(title of childElem as text)
                                    end try
                                end if
                                if buttonText is "" then
                                    try
                                        set buttonText to my cleanText(description of childElem as text)
                                    end try
                                end if
                                if buttonText is "취소" then
                                    try
                                        perform action "AXPress" of childElem
                                    on error
                                        click childElem
                                    end try
                                    return "OK"
                                end if
                            end if
                        end try
                    end repeat
                end try

                try
                    set value of attribute "AXValue" of e to ""
                    set didClear to true
                    exit repeat
                end try
            end if
        end try
    end repeat
    if didClear then
        delay 0.2
        return "OK"
    end if
    return ""
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    try:
        ok = _run_osascript(script, timeout=10).strip() == "OK"
    except Exception:
        return False
    if ok:
        time.sleep(0.25)
    return ok


def _set_recent_notebooks_dialog_search(window: MacWindow, search_text: str) -> bool:
    wanted_text = _clean_field(search_text)
    if not wanted_text:
        return _clear_recent_notebooks_dialog_search(window)
    previous_clipboard = _read_macos_clipboard_text()
    clipboard_written = _write_macos_clipboard_text(wanted_text)
    script = _recent_notebook_dialog_locator(window) + f'''
    set wantedText to {_quote_applescript_text(wanted_text)}
    try
        set frontmost of targetProcess to true
    end try
    try
        perform action "AXRaise" of targetWindow
    end try
    repeat with e in UI elements of targetWindow
        try
            if role of e is "AXTextField" then
                try
                    set value of attribute "AXFocused" of e to true
                end try
                try
                    set value of attribute "AXValue" of e to wantedText
                    delay 0.25
                    return "OK"
                end try
                try
                    set fieldPos to position of e
                    set fieldSize to size of e
                    set clickX to (item 1 of fieldPos) + ((item 1 of fieldSize) / 2)
                    set clickY to (item 2 of fieldPos) + ((item 2 of fieldSize) / 2)
                    click at {{clickX, clickY}}
                    delay 0.1
                end try
                try
                    perform action "AXPress" of e
                end try
                delay 0.1
                try
                    keystroke "a" using command down
                    delay 0.05
                end try
                try
                    key code 51
                    delay 0.05
                end try
                try
                    if {str(bool(clipboard_written)).lower()} then
                        keystroke "v" using command down
                    else
                        keystroke wantedText
                    end if
                end try
                delay 0.35
                return "OK"
            end if
        end try
    end repeat
    return ""
end tell
'''
    try:
        ok = _run_osascript(script, timeout=10).strip() == "OK"
    except Exception:
        ok = False
    finally:
        if clipboard_written:
            _write_macos_clipboard_text(previous_clipboard)
    if ok:
        time.sleep(0.25)
    return ok


def _recent_notebook_dialog_names(window: MacWindow) -> List[str]:
    script = _recent_notebook_dialog_locator(window) + r'''
    set resultItems to {}
    set targetGroup to first UI element of targetWindow whose role is "AXGroup"
    set targetScrollArea to first UI element of targetGroup whose role is "AXScrollArea"
    set targetTable to first UI element of targetScrollArea whose role is "AXTable"
    set rowCount to count of rows of targetTable
    repeat with rowIndex from 1 to rowCount
        try
            set targetRow to row rowIndex of targetTable
            set firstCell to UI element 1 of targetRow
            set labelText to ""
            try
                set labelText to my cleanText(value of attribute "AXDescription" of firstCell as text)
            end try
            if labelText is "" then
                try
                    set labelText to my cleanText(name of firstCell as text)
                end try
            end if
            if labelText is "" then
                try
                    set labelText to my cleanText(value of firstCell as text)
                end try
            end if
            if labelText is not "" then
                set end of resultItems to labelText
            end if
        end try
    end repeat
    set AppleScript's text item delimiters to linefeed
    return resultItems as text
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    raw = _run_osascript(script, timeout=35)
    names: List[str] = []
    seen = set()
    for line in (raw or "").splitlines():
        name = _clean_field(line)
        key = _normalize_text(name)
        if not key or key in seen:
            continue
        seen.add(key)
        names.append(name)
    return names


def _recent_notebook_dialog_row_count(window: MacWindow) -> int:
    script = _recent_notebook_dialog_locator(window) + r'''
    set targetGroup to first UI element of targetWindow whose role is "AXGroup"
    set targetScrollArea to first UI element of targetGroup whose role is "AXScrollArea"
    set targetTable to first UI element of targetScrollArea whose role is "AXTable"
    try
        return count of rows of targetTable
    on error
        return 0
    end try
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    try:
        raw = _run_osascript(script, timeout=15)
    except Exception:
        return 0
    try:
        return int((raw or "0").strip() or "0")
    except Exception:
        return 0


def _wait_for_recent_notebook_rows(window: MacWindow, timeout_sec: float = 4.0) -> bool:
    deadline = time.monotonic() + max(0.5, timeout_sec)
    while time.monotonic() < deadline:
        if _recent_notebook_dialog_row_count(window) > 0:
            return True
        time.sleep(0.15)
    return _recent_notebook_dialog_row_count(window) > 0


def _recent_notebook_dialog_rows(window: MacWindow) -> List[Dict[str, Any]]:
    script = _recent_notebook_dialog_locator(window) + r'''
    set resultItems to {}
    set targetGroup to first UI element of targetWindow whose role is "AXGroup"
    set targetScrollArea to first UI element of targetGroup whose role is "AXScrollArea"
    set targetTable to first UI element of targetScrollArea whose role is "AXTable"
    set rowCount to count of rows of targetTable
    repeat with rowIndex from 1 to rowCount
        try
            set targetRow to row rowIndex of targetTable
            set labelText to ""
            set detailText to ""
            set firstCell to UI element 1 of targetRow
            try
                set labelText to my cleanText(value of attribute "AXDescription" of firstCell as text)
            end try
            if labelText is "" then
                try
                    set labelText to my cleanText(name of firstCell as text)
                end try
            end if
            if labelText is "" then
                try
                    set labelText to my cleanText(value of firstCell as text)
                end try
            end if

            try
                set secondCell to UI element 2 of targetRow
                set detailText to my cleanText(value of attribute "AXDescription" of secondCell as text)
            end try
            if detailText is "" then
                try
                    set thirdCell to UI element 3 of targetRow
                    set detailText to my cleanText(name of thirdCell as text)
                end try
            end if

            try
                set selectedState to value of attribute "AXSelected" of targetRow as text
            on error
                set selectedState to "false"
            end try

            if labelText is not "" then
                set end of resultItems to "ROW" & tab & rowIndex & tab & my cleanText(labelText) & tab & my cleanText(detailText) & tab & selectedState
            end if
        end try
    end repeat
    set AppleScript's text item delimiters to linefeed
    return resultItems as text
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    raw = _run_osascript(script, timeout=45)
    rows: List[Dict[str, Any]] = []
    for line in (raw or "").splitlines():
        parts = line.split("\t")
        if len(parts) < 5 or parts[0] != "ROW":
            continue
        try:
            order = int(parts[1])
        except Exception:
            continue
        label = _clean_field(parts[2])
        detail = _clean_field(parts[3])
        if not label:
            continue
        rows.append(
            {
                "order": order,
                "text": label,
                "detail": detail,
                "selected": parts[4].strip().lower() == "true",
            }
        )
    return rows


def _dismiss_onenote_open_warning_dialog(window: MacWindow) -> bool:
    script = f'''
tell application "System Events"
    set targetProcess to first application process whose unix id is {int(window.process_id())}
    repeat with w in windows of targetProcess
        try
            set dialogText to ""
            repeat with t in static texts of w
                try
                    set dialogText to dialogText & " " & (value of t as text)
                end try
            end repeat
            if dialogText contains "열지 못했습니다" or dialogText contains "권한이 없는" or dialogText contains "couldn't open" or dialogText contains "permission" then
                try
                    perform action "AXPress" of (first button of w whose name is "확인")
                    return "OK"
                end try
                try
                    perform action "AXPress" of (first button of w whose name is "OK")
                    return "OK"
                end try
            end if
        end try
    end repeat
    return ""
end tell
'''
    try:
        ok = _run_osascript(script, timeout=5).strip() == "OK"
    except Exception:
        return False
    if ok:
        time.sleep(0.2)
    return ok


def _drain_onenote_open_warning_dialogs(
    window: MacWindow,
    *,
    timeout_sec: float = 0.6,
    poll_sec: float = 0.12,
) -> bool:
    deadline = time.monotonic() + max(timeout_sec, poll_sec)
    dismissed = False
    misses = 0
    max_misses = 5
    while time.monotonic() < deadline:
        if _dismiss_onenote_open_warning_dialog(window):
            dismissed = True
            misses = 0
            continue
        misses += 1
        if misses >= max_misses:
            break
        time.sleep(max(poll_sec, 0.05))
    return dismissed


def _press_recent_notebook_open(window: MacWindow, notebook_name: str) -> bool:
    script = _recent_notebook_dialog_locator(window) + f'''
    set wantedText to {_quote_applescript_text(notebook_name)}
    set targetGroup to first UI element of targetWindow whose role is "AXGroup"
    set targetScrollArea to first UI element of targetGroup whose role is "AXScrollArea"
    set targetTable to first UI element of targetScrollArea whose role is "AXTable"
    set rowCount to count of rows of targetTable
    set currentIndex to 1
    set targetIndex to 0
    repeat with rowIndex from 1 to rowCount
        try
            set targetRow to row rowIndex of targetTable
            set firstCell to UI element 1 of targetRow
            set labelText to ""
            try
                set labelText to my cleanText(value of attribute "AXDescription" of firstCell as text)
            end try
            if labelText is "" then
                try
                    set labelText to my cleanText(name of firstCell as text)
                end try
            end if
            if labelText is "" then
                try
                    set labelText to my cleanText(value of firstCell as text)
                end try
            end if
            try
                set selectedState to value of attribute "AXSelected" of targetRow as text
            on error
                set selectedState to "false"
            end try
            if selectedState is "true" then
                set currentIndex to rowIndex
            end if
            if labelText is wantedText then
                set targetIndex to rowIndex
            end if
        end try
    end repeat
    if targetIndex is 0 then return ""

    set moveDelta to targetIndex - currentIndex
    try
        set value of attribute "AXFocused" of targetTable to true
    end try
    delay 0.1
    if moveDelta > 0 then
        repeat moveDelta times
            key code 125
            delay 0.04
        end repeat
    else if moveDelta < 0 then
        repeat (moveDelta * -1) times
            key code 126
            delay 0.04
        end repeat
    end if

    delay 0.1
    repeat with e in UI elements of targetWindow
        try
            if role of e is "AXButton" then
                set titleText to ""
                try
                    set titleText to my cleanText(title of e as text)
                end try
                if titleText is "" then
                    try
                        set titleText to my cleanText(name of e as text)
                    end try
                end if
                if titleText is "" then
                    try
                        set titleText to my cleanText(description of e as text)
                    end try
                end if
                if titleText is "열기" then
                    try
                        click e
                    on error
                        perform action "AXPress" of e
                    end try
                    return "OK"
                end if
            end if
        end try
    end repeat
    return ""
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    return _run_osascript(script, timeout=15).strip() == "OK"


def recent_notebook_names(window: MacWindow) -> List[str]:
    opened_sidebar = False
    if _recent_notebook_dialog_row_count(window) <= 0:
        ready, opened_sidebar = _open_recent_notebooks_dialog_with_state(window)
        if not ready and _recent_notebook_dialog_row_count(window) <= 0:
            return []
    try:
        _clear_recent_notebooks_dialog_search(window)
        _wait_for_recent_notebook_rows(window, timeout_sec=6.0)
        names = _recent_notebook_dialog_names(window)
    except Exception:
        names = []
    finally:
        dismiss_recent_notebooks_dialog(window)
        _restore_notebook_sidebar(window, opened_sidebar)
    return names


def _read_onenote_current_notebook_name_quick(window: MacWindow) -> str:
    """Read the visible current-notebook button without scanning the whole window."""
    if window is None or not window.process_id():
        return ""
    script = _applescript_window_locator(window.process_id(), window.window_text()) + r'''
        try
            set targetSplitGroup to first UI element of targetWindow whose role is "AXSplitGroup"
            set leftGroup to first UI element of targetSplitGroup whose role is "AXGroup"
            set nestedSplitGroup to first UI element of leftGroup whose role is "AXSplitGroup"
            repeat with targetButton in (buttons of nestedSplitGroup)
                set candidateText to ""
                try
                    set candidateText to my cleanText(description of targetButton as text)
                end try
                if candidateText contains "(현재 전자필기장)" or candidateText contains "(현재 전자 필기장)" then return candidateText
                try
                    set candidateText to my cleanText(title of targetButton as text)
                end try
                if candidateText contains "(현재 전자필기장)" or candidateText contains "(현재 전자 필기장)" then return candidateText
                try
                    set candidateText to my cleanText(value of targetButton as text)
                end try
                if candidateText contains "(현재 전자필기장)" or candidateText contains "(현재 전자 필기장)" then return candidateText
            end repeat
        end try
        return ""
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    try:
        return _extract_current_notebook_name(_run_osascript(script, timeout=4))
    except Exception:
        return ""


def current_notebook_name(window: Optional[MacWindow]) -> str:
    if window is None:
        return ""
    quick_name = _read_onenote_current_notebook_name_quick(window)
    if quick_name:
        return quick_name
    try:
        return _extract_current_notebook_name(
            _read_onenote_current_notebook_name(
                window.process_id(),
                window.window_text(),
            )
        )
    except Exception:
        return ""


def _is_target_notebook_visible(window: MacWindow, notebook_name: str) -> bool:
    wanted_key = _normalize_text(notebook_name)
    if not wanted_key:
        return False
    quick_name = _read_onenote_current_notebook_name_quick(window)
    if _normalize_text(quick_name) == wanted_key:
        return True
    if quick_name:
        return False
    try:
        current_name = _read_onenote_current_notebook_name(
            window.process_id(), window.window_text()
        )
    except Exception:
        current_name = ""
    if _normalize_text(current_name) == wanted_key:
        return True
    try:
        for info in enumerate_macos_windows():
            if int(info.get("pid") or 0) != window.process_id():
                continue
            if _normalize_text(str(info.get("title") or "")) == wanted_key:
                return True
    except Exception:
        return False
    return False


def _select_notebook_sidebar_row_by_name(window: MacWindow, notebook_name: str) -> bool:
    wanted_name_nfc = _clean_field(notebook_name)
    wanted_name = unicodedata.normalize("NFD", wanted_name_nfc)
    if not wanted_name:
        return False

    script = _applescript_window_locator(window.process_id(), window.window_text()) + f'''
        set wantedText to {_quote_applescript_text(wanted_name)}
        set wantedTextNFC to {_quote_applescript_text(wanted_name_nfc)}
        set targetSplitGroup to first UI element of targetWindow whose role is "AXSplitGroup"
        set leftGroup to first UI element of targetSplitGroup whose role is "AXGroup"
        set nestedSplitGroup to first UI element of leftGroup whose role is "AXSplitGroup"
        set notebookGroup to first UI element of nestedSplitGroup whose role is "AXGroup"
        set notebookScrollArea to first UI element of notebookGroup whose role is "AXScrollArea"
        set targetOutline to first UI element of notebookScrollArea whose role is "AXOutline"

        set rowCount to count of rows of targetOutline
        repeat with rowIndex from 1 to rowCount
            try
                set targetRow to row rowIndex of targetOutline
                set firstCell to UI element 1 of targetRow
                set labelText to ""
                try
                    set labelText to my cleanText(value of attribute "AXDescription" of firstCell as text)
                end try
                if labelText is "" then
                    try
                        set labelText to my cleanText(name of firstCell as text)
                    end try
                end if
                if labelText is "" then
                    try
                        set labelText to my cleanText(value of firstCell as text)
                    end try
                end if
                set labelMatches to false
                if labelText is wantedText then
                    set labelMatches to true
                end if
                if labelText is wantedTextNFC then
                    set labelMatches to true
                end if
                if labelText contains wantedText then
                    set labelMatches to true
                end if
                if labelText contains wantedTextNFC then
                    set labelMatches to true
                end if
                if wantedText contains labelText then
                    set labelMatches to true
                end if
                if wantedTextNFC contains labelText then
                    set labelMatches to true
                end if
                if labelMatches then
                    try
                        set value of attribute "AXSelected" of targetRow to true
                        return "OK"
                    end try
                    try
                        perform action "AXPress" of targetRow
                        return "OK"
                    end try
                    try
                        perform action "AXPress" of firstCell
                        return "OK"
                    end try
                    try
                        click targetRow
                        return "OK"
                    end try
                    try
                        click firstCell
                        return "OK"
                    end try
                    try
                        set rowPosition to position of targetRow
                        set rowSize to size of targetRow
                        set clickX to (item 1 of rowPosition) + 20
                        set clickY to (item 2 of rowPosition) + ((item 2 of rowSize) / 2)
                        click at {{clickX, clickY}}
                        return "OK"
                    end try
                end if
            end try
        end repeat
        return ""
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    return _run_osascript(script, timeout=6).strip() == "OK"


def _select_notebook_sidebar_row_by_index(window: MacWindow, row_index: int) -> bool:
    try:
        target_index = max(1, int(row_index))
    except Exception:
        return False

    script = _applescript_window_locator(window.process_id(), window.window_text()) + f'''
        set targetRowIndex to {target_index}
        set targetSplitGroup to first UI element of targetWindow whose role is "AXSplitGroup"
        set leftGroup to first UI element of targetSplitGroup whose role is "AXGroup"
        set nestedSplitGroup to first UI element of leftGroup whose role is "AXSplitGroup"
        set notebookGroup to first UI element of nestedSplitGroup whose role is "AXGroup"
        set notebookScrollArea to first UI element of notebookGroup whose role is "AXScrollArea"
        set targetOutline to first UI element of notebookScrollArea whose role is "AXOutline"
        set targetRow to row targetRowIndex of targetOutline
        try
            set value of attribute "AXSelected" of targetRow to true
            return "OK"
        end try
        try
            perform action "AXPress" of targetRow
            return "OK"
        end try
        try
            click targetRow
            return "OK"
        end try
        try
            set firstCell to UI element 1 of targetRow
            click firstCell
            return "OK"
        end try
        try
            set rowPosition to position of targetRow
            set rowSize to size of targetRow
            set clickX to (item 1 of rowPosition) + 20
            set clickY to (item 2 of rowPosition) + ((item 2 of rowSize) / 2)
            click at {{clickX, clickY}}
            return "OK"
        end try
        return ""
end tell
'''
    return _run_osascript(script, timeout=6).strip() == "OK"


def _activate_selected_notebook_sidebar_row(window: MacWindow) -> bool:
    script = _applescript_window_locator(window.process_id(), window.window_text()) + r'''
        key code 49
        return "OK"
end tell
'''
    try:
        return _run_osascript(script, timeout=4).strip() == "OK"
    except Exception:
        return False


def select_open_notebook_by_name(
    window: MacWindow,
    notebook_name: str,
    *,
    wait_for_visible: bool = True,
) -> bool:
    wanted_name = _clean_field(notebook_name)
    if not wanted_name:
        return False
    try:
        window.set_focus()
    except Exception:
        pass

    opened_sidebar = False
    try:
        ready, opened_sidebar = _ensure_notebook_sidebar(window)
        if not ready:
            return False

        sidebar_names = _read_open_notebook_names_from_sidebar(window)
        wanted_key = _normalize_text(wanted_name)
        row_index = 0
        for index, name in enumerate(sidebar_names, start=1):
            name_key = _normalize_text(name)
            if name_key == wanted_key or wanted_key in name_key or name_key in wanted_key:
                row_index = index
                break
        if not row_index:
            return False

        if not _select_notebook_sidebar_row_by_index(window, row_index):
            return False
        _activate_selected_notebook_sidebar_row(window)

        if not wait_for_visible:
            _drain_onenote_open_warning_dialogs(window, timeout_sec=0.35, poll_sec=0.1)
            return True

        deadline = time.monotonic() + 6.0
        while time.monotonic() < deadline:
            if _is_target_notebook_visible(window, wanted_name):
                return True
            time.sleep(0.2)
        return _is_target_notebook_visible(window, wanted_name)
    finally:
        if opened_sidebar:
            try:
                _restore_notebook_sidebar(window, opened_sidebar)
            except Exception:
                pass


def open_recent_notebook_by_name(
    window: MacWindow,
    notebook_name: str,
    *,
    wait_for_visible: bool = True,
) -> bool:
    wanted_name = _clean_field(notebook_name)
    if not wanted_name:
        return False
    try:
        window.set_focus()
    except Exception:
        pass
    opened_sidebar = False
    if _recent_notebook_dialog_row_count(window) <= 0:
        ready, opened_sidebar = _open_recent_notebooks_dialog_with_state(window)
        if not ready and _recent_notebook_dialog_row_count(window) <= 0:
            return False
    try:
        _clear_recent_notebooks_dialog_search(window)
        _wait_for_recent_notebook_rows(window, timeout_sec=6.0)
        opened = False
        dismissed_warning = False

        # Recent-notebook search is noticeably faster and more reliable on macOS
        # than walking long tables via repeated arrow-key moves.
        if _set_recent_notebooks_dialog_search(window, wanted_name):
            _wait_for_recent_notebook_rows(window, timeout_sec=2.5)
            opened = _press_recent_notebook_open(window, wanted_name)

        if not opened:
            _clear_recent_notebooks_dialog_search(window)
            _wait_for_recent_notebook_rows(window, timeout_sec=2.5)
            opened = _press_recent_notebook_open(window, wanted_name)
        if opened and not wait_for_visible:
            dismissed_warning = _drain_onenote_open_warning_dialogs(
                window,
                timeout_sec=0.7,
                poll_sec=0.12,
            )
            try:
                _restore_notebook_sidebar(window, opened_sidebar)
            except Exception:
                pass
            return not dismissed_warning
        deadline = time.monotonic() + 12.0
        while time.monotonic() < deadline:
            if _is_target_notebook_visible(window, wanted_name):
                opened = True
                break
            if _dismiss_onenote_open_warning_dialog(window):
                dismissed_warning = True
                break
            time.sleep(0.25)
        if opened:
            try:
                _restore_notebook_sidebar(window, opened_sidebar)
            except Exception:
                pass
            return True
        if dismissed_warning:
            try:
                _restore_notebook_sidebar(window, opened_sidebar)
            except Exception:
                pass
            return False
        dismiss_recent_notebooks_dialog(window)
        _restore_notebook_sidebar(window, opened_sidebar)
        return False
    except Exception:
        try:
            dismiss_recent_notebooks_dialog(window)
        except Exception:
            pass
        _restore_notebook_sidebar(window, opened_sidebar)
        return False


def is_onenote_window_info(info: Dict[str, Any], my_pid: int) -> bool:
    if int(info.get("pid") or 0) == int(my_pid):
        return False
    bundle_id = str(info.get("bundle_id") or "")
    app_name = str(info.get("app_name") or "")
    title = _normalize_text(str(info.get("title") or ""))
    if bundle_id == ONENOTE_MAC_BUNDLE_ID:
        return True
    if app_name in ONENOTE_MAC_APP_NAMES and "onenote" in title:
        return True
    return False


def current_open_notebook_names(window: Optional[MacWindow]) -> List[str]:
    names: List[str] = []
    seen = set()

    def _append_name(raw_name: str) -> None:
        name = _clean_field(str(raw_name or ""))
        key = _normalize_text(name)
        if not key or key in seen:
            return
        seen.add(key)
        names.append(name)

    if window is not None:
        try:
            current_title = _clean_field(window.window_text())
        except Exception:
            current_title = ""
        _append_name(current_title)

    ax_names = _read_open_notebook_names_from_ax(window)
    if ax_names:
        for ax_name in ax_names:
            _append_name(ax_name)
        if names:
            return names

    plist_names = _read_open_notebook_names_from_plist_with_timeout(timeout_sec=1.5)
    if plist_names:
        for plist_name in plist_names:
            _append_name(plist_name)
        return names

    sidebar_names = _read_open_notebook_names_from_sidebar(window)
    if sidebar_names:
        for sidebar_name in sidebar_names:
            _append_name(sidebar_name)
        if names:
            return names

    if names:
        return names

    if window is None:
        return []
    try:
        sidebar_ready, opened_sidebar = _ensure_notebook_sidebar(window)
    except Exception:
        return names
    if not sidebar_ready:
        targets = list_current_notebook_targets(window)
        names = []
        for item in targets:
            if item.get("kind") == "notebook":
                name = str(item.get("notebook") or "").strip()
                if name:
                    names.append(name)
        return names

    try:
        snapshot = collect_onenote_snapshot(window)
        rows = snapshot.get("rows") or []
        names = []
        seen = set()
        for row in sorted(rows, key=lambda item: (not bool(item.get("selected")), int(item.get("order") or 0))):
            name = _clean_field(str(row.get("text") or ""))
            key = _normalize_text(name)
            if not key or key in seen:
                continue
            seen.add(key)
            names.append(name)
        return names
    finally:
        _restore_notebook_sidebar(window, opened_sidebar)


def macos_lookup_targets_json(window: MacWindow) -> str:
    payload = {
        "generated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "count": 0,
        "targets": [],
    }
    targets = list_current_notebook_targets(window)
    payload["targets"] = targets
    payload["count"] = len(targets)
    return json.dumps(payload, ensure_ascii=False)

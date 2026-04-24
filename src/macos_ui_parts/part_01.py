# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

# -*- coding: utf-8 -*-
"""
macOS용 OneNote UI 자동화 헬퍼.

OneNote for Mac의 화면 구조를 기준으로 System Events(접근성)와 osascript를 사용한다.
"""

import ctypes
import os
import sys
import threading
import time
from ctypes import (
    c_bool,
    c_char_p,
    c_double,
    c_int,
    c_long,
    c_longlong,
    c_uint32,
    c_void_p,
    create_string_buffer,
)

from src.lazy_import import LazyModule, LazyPath
from src.platform_support import (
    IS_MACOS,
    MAC_OSASCRIPT_PATH,
    ONENOTE_MAC_APP_NAMES,
    ONENOTE_MAC_BUNDLE_ID,
    open_url_in_system,
)

json = LazyModule("json")
hashlib = LazyModule("hashlib")
subprocess = LazyModule("subprocess")
unicodedata = LazyModule("unicodedata")
_urllib_parse = LazyModule("urllib.parse")


def quote(*args, **kwargs): return _urllib_parse.quote(*args, **kwargs)
def unquote(*args, **kwargs): return _urllib_parse.unquote(*args, **kwargs)
def urlparse(*args, **kwargs): return _urllib_parse.urlparse(*args, **kwargs)

MAC_LSAPPINFO_PATH = "/usr/bin/lsappinfo"
_KCF_STRING_ENCODING_UTF8 = 0x08000100
_KCF_NUMBER_LONGLONG_TYPE = 11
_KAX_VALUE_CGPOINT_TYPE = 1
_KAX_VALUE_CGSIZE_TYPE = 2
_KCG_HID_EVENT_TAP = 0
_KCG_EVENT_LEFT_MOUSE_DOWN = 1
_KCG_EVENT_LEFT_MOUSE_UP = 2
_KCG_MOUSE_BUTTON_LEFT = 0
_KCG_WINDOW_LIST_OPTION_ON_SCREEN_ONLY = 1
_KCG_WINDOW_LIST_EXCLUDE_DESKTOP_ELEMENTS = 16
_CG_WINDOW_INFO_KEYS = (
    "kCGWindowOwnerPID",
    "kCGWindowOwnerName",
    "kCGWindowName",
    "kCGWindowNumber",
    "kCGWindowLayer",
)

class _CGPoint(ctypes.Structure):
    _fields_ = [("x", c_double), ("y", c_double)]


class _CGSize(ctypes.Structure):
    _fields_ = [("width", c_double), ("height", c_double)]


_CF_LIB = None
_APP_SERVICES_LIB = None


def _load_macos_libraries():
    global _CF_LIB, _APP_SERVICES_LIB
    if not IS_MACOS:
        return None, None
    if _CF_LIB is not None and _APP_SERVICES_LIB is not None:
        return _CF_LIB, _APP_SERVICES_LIB

    from ctypes.util import find_library

    cf = ctypes.CDLL(find_library("CoreFoundation"))
    app_services = ctypes.CDLL(find_library("ApplicationServices"))
    app_services.CGWindowListCopyWindowInfo.argtypes = [c_uint32, c_uint32]
    app_services.CGWindowListCopyWindowInfo.restype = c_void_p
    cf.CFArrayGetCount.argtypes = [c_void_p]
    cf.CFArrayGetCount.restype = c_long
    cf.CFArrayGetValueAtIndex.argtypes = [c_void_p, c_long]
    cf.CFArrayGetValueAtIndex.restype = c_void_p
    cf.CFDictionaryGetValue.argtypes = [c_void_p, c_void_p]
    cf.CFDictionaryGetValue.restype = c_void_p
    cf.CFStringCreateWithCString.argtypes = [c_void_p, c_char_p, c_uint32]
    cf.CFStringCreateWithCString.restype = c_void_p
    cf.CFStringGetCString.argtypes = [c_void_p, c_char_p, c_long, c_uint32]
    cf.CFStringGetCString.restype = c_bool
    cf.CFNumberGetValue.argtypes = [c_void_p, c_int, c_void_p]
    cf.CFNumberGetValue.restype = c_bool
    cf.CFGetTypeID.argtypes = [c_void_p]
    cf.CFGetTypeID.restype = c_long
    cf.CFArrayGetTypeID.argtypes = []
    cf.CFArrayGetTypeID.restype = c_long
    cf.CFStringGetTypeID.argtypes = []
    cf.CFStringGetTypeID.restype = c_long
    cf.CFRetain.argtypes = [c_void_p]
    cf.CFRetain.restype = c_void_p
    cf.CFRelease.argtypes = [c_void_p]
    cf.CFRelease.restype = None
    app_services.AXUIElementCreateApplication.argtypes = [c_int]
    app_services.AXUIElementCreateApplication.restype = c_void_p
    app_services.AXUIElementCopyAttributeValue.argtypes = [
        c_void_p,
        c_void_p,
        ctypes.POINTER(c_void_p),
    ]
    app_services.AXUIElementCopyAttributeValue.restype = c_int
    app_services.AXUIElementPerformAction.argtypes = [c_void_p, c_void_p]
    app_services.AXUIElementPerformAction.restype = c_int
    app_services.AXValueGetValue.argtypes = [c_void_p, c_int, c_void_p]
    app_services.AXValueGetValue.restype = c_bool
    app_services.AXIsProcessTrusted.argtypes = []
    app_services.AXIsProcessTrusted.restype = c_bool
    app_services.CGEventCreateMouseEvent.argtypes = [
        c_void_p,
        c_uint32,
        _CGPoint,
        c_uint32,
    ]
    app_services.CGEventCreateMouseEvent.restype = c_void_p
    app_services.CGEventPost.argtypes = [c_uint32, c_void_p]
    app_services.CGEventPost.restype = None
    _CF_LIB = cf
    _APP_SERVICES_LIB = app_services
    return _CF_LIB, _APP_SERVICES_LIB


class _LazyMacOSLibrary:
    def __init__(self, index: int):
        self._index = index

    def __bool__(self) -> bool:
        return bool(IS_MACOS)

    def __getattr__(self, name: str):
        library = _load_macos_libraries()[self._index]
        if library is None:
            raise AttributeError(name)
        return getattr(library, name)


_CF = _LazyMacOSLibrary(0)
_APP_SERVICES = _LazyMacOSLibrary(1)

_MAC_BUNDLE_ID_CACHE: Dict[int, str] = {}
_MAC_ONENOTE_NOTEBOOKS_PLIST = LazyPath(
    os.path.expanduser(
        "~/Library/Group Containers/UBF8T346G9.Office/OneNote/ShareExtension/Notebooks.plist"
    )
)
_MAC_ONENOTE_RESOURCEINFOCACHE_JSON = LazyPath(
    os.path.expanduser(
        "~/Library/Containers/com.microsoft.onenote.mac/Data/Library/Application Support/Microsoft/Office/16.0/ResourceInfoCache/data.json"
    )
)
_MAC_DEBUG_LOG_PATH = LazyPath(
    os.path.expanduser("~/Library/Logs/OneNote_Remocon/macos_ui_debug.log")
)
_MAC_RECENT_CACHE_TIMED_OUT = False
_MAC_RECENT_CACHE_READER_PATHS = ("/usr/bin/python3", sys.executable)
_MAC_LAST_AX_NOTEBOOK_DEBUG: Dict[str, Any] = {}
_MAC_OPEN_TAB_RECORDS_CACHE_TTL_SEC = 300.0
_MAC_OPEN_TAB_RECORDS_CACHE: Dict[str, Any] = {
    "timestamp": 0.0,
    "records": [],
}
_MAC_OPEN_TAB_RECORDS_CACHE_LOCK = threading.Lock()
_MAC_RECENT_NOTEBOOK_DIALOG_TOKENS = (
    "최근 전자 필기장",
    "최근 전자필기장",
    "새 전자 필기장",
    "새 전자필기장",
    "새 전자 필기장 및 최근 전자 필기장 열기",
    "recent notebook",
    "new and recent notebook",
    "new notebook",
)


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


def _append_macos_debug_log(line: str) -> None:
    try:
        _MAC_DEBUG_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with _MAC_DEBUG_LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(
                f"{time.strftime('%Y-%m-%d %H:%M:%S')} {str(line or '').rstrip()}\n"
            )
    except Exception:
        pass


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


def _is_recent_notebook_dialog_title(value: str) -> bool:
    title = _clean_field(value).casefold()
    if not title:
        return False
    return any(token.casefold() in title for token in _MAC_RECENT_NOTEBOOK_DIALOG_TOKENS)


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


def _cf_retain(value: c_void_p) -> Optional[c_void_p]:
    if not (_CF and value):
        return None
    try:
        retained = _CF.CFRetain(value)
        return c_void_p(retained) if retained else None
    except Exception:
        return None


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

_publish_context(globals())

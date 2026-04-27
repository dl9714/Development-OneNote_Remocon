# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


_QUICK_PLIST_CACHE: Dict[str, Any] = {
    "key": None,
    "timestamp": 0.0,
    "names": [],
}


def _xml_unescape_basic(value: str) -> str:
    if "&" not in value:
        return value
    return (
        value.replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
    )


def _read_open_notebook_names_from_plist_fast_xml() -> Optional[List[str]]:
    names: List[str] = []
    seen = set()
    dict_depth = 0
    pending_key = ""
    current_name = ""
    current_type: Optional[int] = None
    in_data = False
    saw_xml = False

    def _append_name(raw_name: str) -> None:
        name = _clean_field(str(raw_name or ""))
        key = _normalize_text(name)
        if not key or key in seen:
            return
        seen.add(key)
        names.append(name)

    with _MAC_ONENOTE_NOTEBOOKS_PLIST.open(
        "r",
        encoding="utf-8",
        errors="ignore",
    ) as f:
        for raw_line in f:
            if in_data:
                if "</data>" in raw_line:
                    in_data = False
                continue
            line = raw_line.strip()
            if line.startswith("<?xml"):
                saw_xml = True
            if line == "<data>":
                in_data = True
                continue
            if line == "<dict>":
                dict_depth += 1
                if dict_depth == 1:
                    pending_key = ""
                    current_name = ""
                    current_type = None
                continue
            if line == "</dict>":
                if dict_depth == 1 and current_type == 1:
                    _append_name(current_name)
                dict_depth = max(0, dict_depth - 1)
                pending_key = ""
                continue
            if dict_depth != 1:
                continue
            if line.startswith("<key>") and line.endswith("</key>"):
                pending_key = line[5:-6]
                continue
            if pending_key == "Name":
                if line.startswith("<string>") and line.endswith("</string>"):
                    current_name = _xml_unescape_basic(line[8:-9])
                pending_key = ""
            elif pending_key == "Type":
                if line.startswith("<integer>") and line.endswith("</integer>"):
                    try:
                        current_type = int(line[9:-10].strip())
                    except Exception:
                        current_type = None
                pending_key = ""

    return names if saw_xml else None


def _quick_plist_names(timeout_sec: float) -> Tuple[List[str], bool]:
    try:
        stat = _MAC_ONENOTE_NOTEBOOKS_PLIST.stat()
        key = (int(stat.st_mtime_ns), int(stat.st_size))
    except Exception:
        key = None
    now = time.monotonic()
    if key is not None and _QUICK_PLIST_CACHE.get("key") == key:
        return list(_QUICK_PLIST_CACHE.get("names") or []), False
    names = _read_open_notebook_names_from_plist()
    if key is not None:
        _QUICK_PLIST_CACHE.update(
            {"key": key, "timestamp": now, "names": list(names)}
        )
    return list(names), False



def current_open_notebook_names_quick(
    window: Optional[MacWindow],
    *,
    ax_timeout_sec: float = 0.8,
    plist_timeout_sec: float = 1.2,
    sidebar_timeout_sec: float = 0.6,
    min_names_before_sidebar: int = 8,
) -> Dict[str, Any]:
    names: List[str] = []
    seen = set()
    debug: Dict[str, Any] = {
        "title_count": 0,
        "ax_count": 0,
        "ax_error": "",
        "ax_timed_out": False,
        "plist_count": 0,
        "plist_cached": False,
        "plist_timed_out": False,
        "sidebar_count": 0,
        "sidebar_error": "",
        "sidebar_timed_out": False,
    }

    def _append_name(raw_name: str) -> None:
        name = _clean_field(str(raw_name or ""))
        if _is_recent_notebook_dialog_title(name):
            return
        key = _normalize_text(name)
        if not key or key in seen:
            return
        seen.add(key)
        names.append(name)

    if window is not None:
        info_title = _clean_field(str(getattr(window, "info", {}).get("title") or ""))
        try:
            current_title = _clean_field(window.window_text())
        except Exception:
            current_title = ""
        for title_candidate in (info_title, current_title):
            if title_candidate and not _is_recent_notebook_dialog_title(title_candidate):
                _append_name(title_candidate)
                debug["title_count"] = 1
                break

    plist_cache_before = _QUICK_PLIST_CACHE.get("timestamp")
    plist_names, plist_timed_out = _quick_plist_names(plist_timeout_sec)
    debug["plist_timed_out"] = plist_timed_out
    debug["plist_cached"] = (
        bool(plist_cache_before)
        and plist_cache_before == _QUICK_PLIST_CACHE.get("timestamp")
    )
    for plist_name in plist_names:
        _append_name(plist_name)
    debug["plist_count"] = len(plist_names)

    should_probe_ax = (
        window is not None
        and float(ax_timeout_sec or 0.0) > 0.0
        and len(names) < max(0, int(min_names_before_sidebar))
    )
    if should_probe_ax:
        ax_names, ax_error, ax_timed_out = _read_notebook_names_with_timeout(
            lambda: _read_open_notebook_names_from_ax(window),
            ax_timeout_sec,
        )
        debug["ax_timed_out"] = ax_timed_out
        debug["ax_error"] = ax_error
        for ax_name in ax_names or []:
            _append_name(ax_name)
        debug["ax_count"] = len(ax_names or [])

    should_probe_sidebar = (
        window is not None
        and float(sidebar_timeout_sec or 0.0) > 0.0
        and len(names) < max(0, int(min_names_before_sidebar))
    )
    if should_probe_sidebar:
        sidebar_names, sidebar_error, sidebar_timed_out = _read_notebook_names_with_timeout(
            lambda: _read_open_notebook_names_from_sidebar(window),
            sidebar_timeout_sec,
        )
        debug["sidebar_timed_out"] = sidebar_timed_out
        debug["sidebar_error"] = sidebar_error
        for sidebar_name in sidebar_names or []:
            _append_name(sidebar_name)
        debug["sidebar_count"] = len(sidebar_names or [])

    return {
        "names": names,
        "debug": debug,
    }


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

_publish_context(globals())

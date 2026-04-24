# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



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

        ax_names, ax_error, ax_timed_out = _read_notebook_names_with_timeout(
            lambda: _read_open_notebook_names_from_ax(window),
            ax_timeout_sec,
        )
        debug["ax_timed_out"] = ax_timed_out
        debug["ax_error"] = ax_error
        for ax_name in ax_names or []:
            _append_name(ax_name)
        debug["ax_count"] = len(ax_names or [])

    plist_names = _read_open_notebook_names_from_plist_with_timeout(
        timeout_sec=plist_timeout_sec
    )
    if plist_names is None:
        debug["plist_timed_out"] = True
        plist_names = []
    for plist_name in plist_names:
        _append_name(plist_name)
    debug["plist_count"] = len(plist_names)

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

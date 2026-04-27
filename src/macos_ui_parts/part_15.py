# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



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
    if wait_for_visible:
        try:
            window.set_focus()
        except Exception:
            pass

    if _select_open_notebook_by_name_ax(
        window,
        wanted_name,
        wait_for_visible=wait_for_visible,
    ):
        return True

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
    fast: bool = False,
) -> bool:
    wanted_name = _clean_field(notebook_name)
    if not wanted_name:
        return False
    if wait_for_visible:
        try:
            window.set_focus()
        except Exception:
            pass
    opened_sidebar = False
    row_count_timeout = 2 if fast else 15
    clear_timeout = 3 if fast else 10
    search_timeout = 4 if fast else 10
    press_timeout = 5 if fast else 15
    if _recent_notebook_dialog_row_count(window, timeout=row_count_timeout) <= 0:
        ready, opened_sidebar = _open_recent_notebooks_dialog_with_state(
            window,
            fast=fast,
        )
        if (
            not ready
            and _recent_notebook_dialog_row_count(window, timeout=row_count_timeout) <= 0
        ):
            return False
    try:
        initial_rows_timeout = (
            1.2 if fast else (6.0 if wait_for_visible else 1.2)
        )
        search_rows_timeout = (
            0.5 if fast else (2.5 if wait_for_visible else 0.55)
        )
        search_settle_sec = 0.08 if fast else (0.25 if wait_for_visible else 0.09)
        _clear_recent_notebooks_dialog_search(
            window,
            settle_sec=search_settle_sec,
            timeout=clear_timeout,
        )
        _wait_for_recent_notebook_rows(
            window,
            timeout_sec=initial_rows_timeout,
            row_count_timeout=row_count_timeout,
        )
        opened = False
        dismissed_warning = False

        # Recent-notebook search is noticeably faster and more reliable on macOS
        # than walking long tables via repeated arrow-key moves. For the fast
        # "open all" worker, keep search settle waits short so missing results
        # do not stall every candidate for multiple seconds.
        if _set_recent_notebooks_dialog_search(
            window,
            wanted_name,
            settle_sec=search_settle_sec,
            timeout=search_timeout,
        ):
            _wait_for_recent_notebook_rows(
                window,
                timeout_sec=search_rows_timeout,
                row_count_timeout=row_count_timeout,
            )
            opened = _press_recent_notebook_open(
                window,
                wanted_name,
                timeout=press_timeout,
            )

        if not opened:
            _clear_recent_notebooks_dialog_search(
                window,
                settle_sec=search_settle_sec,
                timeout=clear_timeout,
            )
            _wait_for_recent_notebook_rows(
                window,
                timeout_sec=search_rows_timeout,
                row_count_timeout=row_count_timeout,
            )
            opened = _press_recent_notebook_open(
                window,
                wanted_name,
                timeout=press_timeout,
            )
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

    quick_plist_names = globals().get("_quick_plist_names")
    if callable(quick_plist_names):
        try:
            plist_names, _plist_timed_out = quick_plist_names(1.5)
        except Exception:
            plist_names = []
    else:
        plist_names = _read_open_notebook_names_from_plist_with_timeout(timeout_sec=1.5)
    if plist_names:
        for plist_name in plist_names:
            _append_name(plist_name)
        if len(names) >= 12:
            return names

    ax_names = _read_open_notebook_names_from_ax(window)
    if ax_names:
        for ax_name in ax_names:
            _append_name(ax_name)
        if len(names) >= 12:
            return names

    # AX/plist sources are fast but sometimes incomplete on macOS.
    # When they return only a small subset, merge the live sidebar too so
    # callers such as "open all notebooks" can skip notebooks that are
    # already open instead of reprocessing them.
    if len(names) < 8:
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

_publish_context(globals())

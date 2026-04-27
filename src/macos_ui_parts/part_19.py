# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


def open_tab_notebook_by_name(
    window: Optional[MacWindow],
    notebook_name: str,
    *,
    wait_for_visible: bool = True,
    fast: bool = False,
) -> bool:
    wanted_name = _clean_field(notebook_name)
    if window is None or not wanted_name:
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
    rows_timeout = 0.7 if fast else 2.5
    search_settle_sec = 0.08 if fast else 0.25

    try:
        ready, opened_sidebar = _ensure_open_tab_notebooks_dialog(
            window,
            fast=fast,
        )
        if not ready and _recent_notebook_dialog_row_count(
            window,
            timeout=row_count_timeout,
        ) <= 0:
            return False

        opened = False
        _clear_recent_notebooks_dialog_search(
            window,
            settle_sec=search_settle_sec,
            timeout=clear_timeout,
        )
        _wait_for_recent_notebook_rows(
            window,
            timeout_sec=rows_timeout,
            row_count_timeout=row_count_timeout,
        )
        if _set_recent_notebooks_dialog_search(
            window,
            wanted_name,
            settle_sec=search_settle_sec,
            timeout=search_timeout,
        ):
            _wait_for_recent_notebook_rows(
                window,
                timeout_sec=rows_timeout,
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
                timeout_sec=rows_timeout,
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
            return not dismissed_warning

        deadline = time.monotonic() + 12.0
        while opened and time.monotonic() < deadline:
            if _is_target_notebook_visible(window, wanted_name):
                return True
            if _dismiss_onenote_open_warning_dialog(window):
                return False
            time.sleep(0.25)
        return bool(opened and _is_target_notebook_visible(window, wanted_name))
    except Exception:
        return False
    finally:
        try:
            dismiss_recent_notebooks_dialog(window)
        except Exception:
            pass
        try:
            _restore_notebook_sidebar(window, opened_sidebar)
        except Exception:
            pass


_publish_context(globals())

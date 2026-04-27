# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _open_recent_notebooks_dialog_with_state(
    window: MacWindow,
    *,
    fast: bool = False,
) -> Tuple[bool, bool]:
    row_count_timeout = 2 if fast else 15
    press_timeout = 4 if fast else 15
    locator_timeout = 2 if fast else 5
    open_wait_sec = 2.0 if fast else 4.0
    if _recent_notebook_dialog_row_count(window, timeout=row_count_timeout) > 0:
        return True, False
    sidebar_ready, opened_sidebar = _ensure_notebook_sidebar(window)
    if not sidebar_ready:
        return False, False
    try:
        add_timeout = 4 if fast else 8
        if not (
            _press_add_notebook_button_in_sidebar(window, timeout=add_timeout)
            or _press_window_button(
                window,
                help_contains=["전자 필기장 만들기 또는 열기"],
                value_equals=["전자 필기장 추가"],
                timeout=press_timeout,
            )
        ):
            _restore_notebook_sidebar(window, opened_sidebar)
            return False, False
        deadline = time.monotonic() + open_wait_sec
        while time.monotonic() < deadline:
            time.sleep(0.12)
            try:
                _run_osascript(
                    _recent_notebook_dialog_locator(window) + '\nreturn "OK"\nend tell\n',
                    timeout=locator_timeout,
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


def _clear_recent_notebooks_dialog_search(
    window: MacWindow,
    *,
    settle_sec: float = 0.25,
    timeout: int = 10,
) -> bool:
    settle_delay = max(0.05, float(settle_sec))
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
        delay ''' + f"{settle_delay:.2f}" + r'''
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
        ok = _run_osascript(script, timeout=timeout).strip() == "OK"
    except Exception:
        return False
    return ok


def _set_recent_notebooks_dialog_search(
    window: MacWindow,
    search_text: str,
    *,
    settle_sec: float = 0.25,
    timeout: int = 10,
) -> bool:
    wanted_text = _clean_field(search_text)
    if not wanted_text:
        return _clear_recent_notebooks_dialog_search(
            window,
            settle_sec=settle_sec,
            timeout=timeout,
        )
    settle_delay = max(0.08, float(settle_sec))
    focus_delay = max(0.04, min(0.12, settle_delay / 2.0))
    key_delay = max(0.03, min(0.06, settle_delay / 3.0))
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
                    delay {settle_delay:.2f}
                    return "OK"
                end try
                try
                    set fieldPos to position of e
                    set fieldSize to size of e
                    set clickX to (item 1 of fieldPos) + ((item 1 of fieldSize) / 2)
                    set clickY to (item 2 of fieldPos) + ((item 2 of fieldSize) / 2)
                    click at {{clickX, clickY}}
                    delay {focus_delay:.2f}
                end try
                try
                    perform action "AXPress" of e
                end try
                delay {focus_delay:.2f}
                try
                    keystroke "a" using command down
                    delay {key_delay:.2f}
                end try
                try
                    key code 51
                    delay {key_delay:.2f}
                end try
                try
                    if {str(bool(clipboard_written)).lower()} then
                        keystroke "v" using command down
                    else
                        keystroke wantedText
                    end if
                end try
                delay {settle_delay:.2f}
                return "OK"
            end if
        end try
    end repeat
    return ""
end tell
'''
    try:
        ok = _run_osascript(script, timeout=timeout).strip() == "OK"
    except Exception:
        ok = False
    finally:
        if clipboard_written:
            _write_macos_clipboard_text(previous_clipboard)
    return ok

_publish_context(globals())

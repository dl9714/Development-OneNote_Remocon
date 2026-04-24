# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _press_recent_notebook_open(
    window: MacWindow,
    notebook_name: str,
    *,
    timeout: int = 15,
) -> bool:
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
    return _run_osascript(script, timeout=timeout).strip() == "OK"


def recent_notebook_names(window: MacWindow, *, fast: bool = False) -> List[str]:
    opened_sidebar = False
    row_count_timeout = 2 if fast else 15
    clear_timeout = 3 if fast else 10
    rows_timeout = 1.2 if fast else 6.0
    names_timeout = 6 if fast else 35
    if _recent_notebook_dialog_row_count(window, timeout=row_count_timeout) <= 0:
        ready, opened_sidebar = _open_recent_notebooks_dialog_with_state(
            window,
            fast=fast,
        )
        if (
            not ready
            and _recent_notebook_dialog_row_count(window, timeout=row_count_timeout) <= 0
        ):
            return []
    try:
        _clear_recent_notebooks_dialog_search(window, timeout=clear_timeout)
        _wait_for_recent_notebook_rows(
            window,
            timeout_sec=rows_timeout,
            row_count_timeout=row_count_timeout,
        )
        names = _recent_notebook_dialog_names(window, timeout=names_timeout)
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

_publish_context(globals())

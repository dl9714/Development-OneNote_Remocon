# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _recent_notebook_dialog_names(
    window: MacWindow,
    *,
    timeout: int = 35,
) -> List[str]:
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
    raw = _run_osascript(script, timeout=timeout)
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


def _recent_notebook_dialog_row_count(
    window: MacWindow,
    *,
    timeout: int = 15,
) -> int:
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
        raw = _run_osascript(script, timeout=timeout)
    except Exception:
        return 0
    try:
        return int((raw or "0").strip() or "0")
    except Exception:
        return 0


def _wait_for_recent_notebook_rows(
    window: MacWindow,
    timeout_sec: float = 4.0,
    *,
    row_count_timeout: int = 15,
) -> bool:
    deadline = time.monotonic() + max(0.5, timeout_sec)
    while time.monotonic() < deadline:
        if _recent_notebook_dialog_row_count(window, timeout=row_count_timeout) > 0:
            return True
        time.sleep(0.15)
    return _recent_notebook_dialog_row_count(window, timeout=row_count_timeout) > 0


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

_publish_context(globals())

# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


def _select_open_notebook_by_name_direct(
    window: MacWindow,
    notebook_name: str,
    *,
    wait_for_visible: bool = True,
) -> bool:
    wanted_name = _clean_field(notebook_name)
    wanted_compact = _normalize_text(wanted_name)
    wanted_flat = str(wanted_compact or "").replace("-", "")
    if not (window is not None and wanted_name and wanted_compact):
        return False
    try:
        window.set_focus()
        time.sleep(0.06)
    except Exception:
        pass

    script = _applescript_window_locator(window.process_id(), window.window_text()) + f'''
        set wantedName to {_quote_applescript_text(wanted_name)}
        set wantedCompact to {_quote_applescript_text(wanted_compact)}
        set wantedFlat to {_quote_applescript_text(wanted_flat)}
        set targetOutline to missing value
        if my notebookSidebarOpen(targetWindow) is false then
            my pressNotebookButton(targetWindow)
            delay 0.14
        end if

        repeat with attemptIndex from 1 to 8
            try
                set targetSplitGroup to first UI element of targetWindow whose role is "AXSplitGroup"
                set leftGroup to first UI element of targetSplitGroup whose role is "AXGroup"
                set nestedSplitGroup to first UI element of leftGroup whose role is "AXSplitGroup"
                set notebookGroup to first UI element of nestedSplitGroup whose role is "AXGroup"
                set notebookScrollArea to first UI element of notebookGroup whose role is "AXScrollArea"
                set targetOutline to first UI element of notebookScrollArea whose role is "AXOutline"
                exit repeat
            on error
                my pressNotebookButton(targetWindow)
                delay 0.12
            end try
        end repeat

        if targetOutline is missing value then return ""
        repeat with targetRow in rows of targetOutline
            set labelText to my rowLabel(targetRow)
            set labelCompact to my compactText(labelText)
            if labelText is wantedName or labelText contains wantedName or wantedName contains labelText or labelCompact is wantedCompact or labelCompact is wantedFlat then
                try
                    set value of attribute "AXSelected" of targetRow to true
                end try
                delay 0.03
                key code 49
                return "OK" & tab & labelText
            end if
        end repeat
        return ""
end tell

on notebookSidebarOpen(targetWindow)
    tell application "System Events"
    try
        set targetSplitGroup to first UI element of targetWindow whose role is "AXSplitGroup"
        set leftGroup to first UI element of targetSplitGroup whose role is "AXGroup"
        set nestedSplitGroup to first UI element of leftGroup whose role is "AXSplitGroup"
        set scanItems to UI elements of nestedSplitGroup
    on error
        set scanItems to entire contents of targetWindow
    end try
    repeat with e in scanItems
        try
            set roleText to role of e as text
        on error
            set roleText to ""
        end try
        if roleText is "AXButton" or roleText is "AXMenuButton" then
            set helpText to ""
            try
                set helpText to my cleanText(help of e as text)
            end try
            if helpText contains "전자 필기장 목록 숨기기" or helpText contains "Hide notebook list" then
                return true
            end if
        end if
    end repeat
    return false
    end tell
end notebookSidebarOpen

on pressNotebookButton(targetWindow)
    tell application "System Events"
    try
        try
            set targetSplitGroup to first UI element of targetWindow whose role is "AXSplitGroup"
            set leftGroup to first UI element of targetSplitGroup whose role is "AXGroup"
            set nestedSplitGroup to first UI element of leftGroup whose role is "AXSplitGroup"
            set scanItems to UI elements of nestedSplitGroup
        on error
            set scanItems to entire contents of targetWindow
        end try
        repeat with e in scanItems
            try
                set roleText to role of e as text
            on error
                set roleText to ""
            end try
            if roleText is "AXButton" or roleText is "AXMenuButton" then
                set helpText to ""
                try
                    set helpText to my cleanText(help of e as text)
                end try
                if helpText contains "전자 필기장 보기" or helpText contains "View, create" or helpText contains "View notebooks" then
                    try
                        perform action "AXPress" of e
                    on error
                        click e
                    end try
                    return true
                end if
            end if
        end repeat
    end try
    return false
    end tell
end pressNotebookButton

on rowLabel(targetRow)
    tell application "System Events"
    set labelText to ""
    try
        set firstCell to UI element 1 of targetRow
        try
            set labelText to my cleanText((value of attribute "AXDescription" of firstCell) as text)
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
        try
            set labelText to my cleanText(value of targetRow as text)
        end try
    end if
    if labelText contains ", 동기화할 수 없습니다" then
        set AppleScript's text item delimiters to ", 동기화할 수 없습니다"
        set labelText to item 1 of text items of labelText
        set AppleScript's text item delimiters to ""
    end if
    return labelText
    end tell
end rowLabel

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    repeat while t contains "  "
        set t to my replaceText("  ", " ", t)
    end repeat
    return t
end cleanText

on compactText(v)
    set t to my cleanText(v)
    set t to my replaceText(" ", "", t)
    set t to my replaceText("-", "", t)
    return t
end compactText

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
        raw = _run_osascript(script, timeout=5).strip()
    except Exception:
        return False
    if not raw.startswith("OK"):
        return False
    if not wait_for_visible:
        return True
    deadline = time.monotonic() + 3.0
    while time.monotonic() < deadline:
        if _is_target_notebook_visible(window, wanted_name):
            return True
        time.sleep(0.15)
    return True


_publish_context(globals())

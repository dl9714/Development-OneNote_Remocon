# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _press_add_notebook_button_in_sidebar(
    window: MacWindow,
    *,
    timeout: int = 4,
) -> bool:
    """Press OneNote's Add Notebook button from the left sidebar only."""
    locator = _applescript_window_locator(window.process_id(), window.window_text())
    help_matchers = [
        "전자 필기장 만들기 또는 열기",
        "Create or open a notebook",
        "Create or open notebooks",
    ]
    value_matchers = [
        "전자 필기장 추가",
        "Add Notebook",
        "Add notebook",
    ]
    help_tokens = ", ".join(_quote_applescript_text(item) for item in help_matchers)
    value_tokens = ", ".join(_quote_applescript_text(item) for item in value_matchers)
    script = locator + f'''
        set helpMatchers to {{{help_tokens}}}
        set valueMatchers to {{{value_tokens}}}
        set scanItems to {{}}
        try
            set targetSplitGroup to first UI element of targetWindow whose role is "AXSplitGroup"
            set leftGroup to first UI element of targetSplitGroup whose role is "AXGroup"
            set nestedSplitGroup to first UI element of leftGroup whose role is "AXSplitGroup"
            try
                set scanItems to scanItems & (buttons of nestedSplitGroup)
            end try
            repeat with rootElem in UI elements of nestedSplitGroup
                try
                    set scanItems to scanItems & (buttons of rootElem)
                end try
            end repeat
            if (count of scanItems) is 0 then
                set scanItems to entire contents of nestedSplitGroup
            end if
        on error
            set scanItems to {{}}
        end try

        repeat with e in scanItems
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
                repeat with wantedValue in valueMatchers
                    if wantedValue is not "" and (titleText is wantedValue or nameText is wantedValue or descText is wantedValue or valueText is wantedValue) then
                        set shouldPress to true
                        exit repeat
                    end if
                end repeat
                if shouldPress is false then
                    repeat with wantedHelp in helpMatchers
                        if wantedHelp is not "" and helpText contains wantedHelp then
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
    try:
        return _run_osascript(script, timeout=timeout).strip() == "OK"
    except Exception:
        return False


def _has_notebook_sidebar_ui(window: MacWindow) -> bool:
    script = _applescript_window_locator(window.process_id(), window.window_text()) + r'''
        try
            set targetSplitGroup to first UI element of targetWindow whose role is "AXSplitGroup"
            set leftGroup to first UI element of targetSplitGroup whose role is "AXGroup"
            set nestedSplitGroup to first UI element of leftGroup whose role is "AXSplitGroup"
            set scanItems to UI elements of nestedSplitGroup
        on error
            set scanItems to {}
        end try
        repeat with e in scanItems
            try
                set roleText to role of e as text
                if roleText is "AXButton" or roleText is "AXMenuButton" then
                    set helpText to ""
                    set valueText to ""
                    set titleText to ""
                    try
                        set helpText to my cleanText(help of e as text)
                    end try
                    try
                        set valueText to my cleanText(value of e as text)
                    end try
                    try
                        set titleText to my cleanText(title of e as text)
                    end try
                    if helpText contains "전자 필기장 목록 숨기기" then
                        return "OK"
                    end if
                    if helpText contains "Hide notebook list" then
                        return "OK"
                    end if
                    if (valueText is "전자 필기장" or titleText is "전자 필기장") and helpText contains "숨기기" then
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
    try:
        return _run_osascript(script, timeout=6).strip() == "OK"
    except Exception:
        return False


def _ensure_notebook_sidebar(window: MacWindow) -> Tuple[bool, bool]:
    try:
        window.set_focus()
        time.sleep(0.08)
    except Exception:
        pass
    if _has_notebook_sidebar_ui(window):
        return True, False
    if not (
        _press_notebook_sidebar_button(window, open_sidebar=True, timeout=4)
        or _press_window_button(
            window,
            help_contains=["전자 필기장 보기, 만들기 또는 열기"],
            timeout=6,
        )
    ):
        return False, False
    for _ in range(12):
        time.sleep(0.15)
        if _has_notebook_sidebar_ui(window):
            return True, True
    return False, True


def _restore_notebook_sidebar(window: MacWindow, opened_by_us: bool) -> None:
    if not opened_by_us:
        return
    try:
        if _press_notebook_sidebar_button(window, open_sidebar=False, timeout=4):
            return
        _press_window_button(
            window,
            help_contains=["전자 필기장 목록 숨기기"],
            value_equals=["전자 필기장"],
            timeout=6,
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
    set titleTokens to {{"최근 전자 필기장", "최근 전자필기장", "새 전자 필기장", "새 전자필기장", "새 전자 필기장 및 최근 전자 필기장 열기", "Recent Notebooks", "New Notebook", "New and Recent Notebooks"}}

    repeat with w in windows of targetProcess
        try
            set winName to name of w as text
        on error
            set winName to ""
        end try
        repeat with titleToken in titleTokens
            if winName contains (titleToken as text) then
                set targetWindow to w
                exit repeat
            end if
        end repeat
        if targetWindow is not missing value then exit repeat
    end repeat

    if targetWindow is missing value then
        repeat with w in windows of targetProcess
            try
                set hasTextField to false
                set hasOpenButton to false
                set hasRecentRadio to false
                try
                    if exists (first UI element of w whose role is "AXTextField") then set hasTextField to true
                end try
                try
                    if exists (first button of w whose name is "열기") then set hasOpenButton to true
                end try
                try
                    if exists (first button of w whose name is "Open") then set hasOpenButton to true
                end try
                try
                    if exists (first radio button of w whose name is "최근") then set hasRecentRadio to true
                end try
                try
                    if exists (first radio button of w whose name is "Recent") then set hasRecentRadio to true
                end try
                if hasTextField and (hasOpenButton or hasRecentRadio) then
                    set targetWindow to w
                    exit repeat
                end if
            end try
        end repeat
    end if

    if targetWindow is missing value then
        repeat with w in windows of targetProcess
            try
                set candidateNumber to value of attribute "AXWindowNumber" of w
                if targetWindowNumber > 0 and (candidateNumber as integer) is targetWindowNumber then
                    set targetWindow to w
                    exit repeat
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
    roots = _ax_onenote_notebook_dialog_roots(window)
    try:
        return bool(roots)
    finally:
        _release_ax_refs(roots)

_publish_context(globals())

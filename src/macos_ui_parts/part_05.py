# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def select_row_by_text(
    window: MacWindow,
    text: str,
    preferred_scroll_left: Optional[int] = None,
) -> bool:
    return _select_outline_row_by_text(
        window,
        text,
        preferred_scroll_left=preferred_scroll_left,
    )


def _select_outline_row_by_text(
    window: MacWindow,
    text: str,
    preferred_scroll_left: Optional[int] = None,
    *,
    prefer_rightmost: bool = False,
) -> bool:
    wanted_text_nfc = _clean_field(text)
    wanted_text = unicodedata.normalize("NFD", wanted_text_nfc)
    if not wanted_text:
        return False

    locator = _applescript_window_locator(window.process_id(), window.window_text())
    wanted_left = -1 if preferred_scroll_left is None else int(preferred_scroll_left)
    rightmost_flag = "true" if prefer_rightmost else "false"
    script = locator + f'''
        set wantedText to {_quote_applescript_text(wanted_text)}
        set wantedTextNFC to {_quote_applescript_text(wanted_text_nfc)}
        set wantedScrollLeft to {wanted_left}
        set preferRightmost to {rightmost_flag}
        set bestRow to missing value
        set bestScore to 999999
        if preferRightmost then set bestScore to -999999

        try
            set targetSplitGroup to first UI element of targetWindow whose role is "AXSplitGroup"
            set leftGroup to first UI element of targetSplitGroup whose role is "AXGroup"
            set nestedSplitGroup to first UI element of leftGroup whose role is "AXSplitGroup"
            set groupsToScan to UI elements of nestedSplitGroup
        on error
            set groupsToScan to UI elements of targetWindow
        end try

        repeat with g in groupsToScan
            try
                set gRole to (role of g) as text
                set shouldScanGroup to false
                if gRole is "AXGroup" then
                    set shouldScanGroup to true
                end if
                if gRole is "AXScrollArea" then
                    set shouldScanGroup to true
                end if
                if shouldScanGroup then
                    set scrollAreas to {{}}
                    if gRole is "AXScrollArea" then
                        set end of scrollAreas to g
                    else
                        repeat with maybeScrollArea in (UI elements of g)
                            try
                                if ((role of maybeScrollArea) as text) is "AXScrollArea" then
                                    set end of scrollAreas to maybeScrollArea
                                end if
                            end try
                        end repeat
                    end if

                    repeat with sa in scrollAreas
                        try
                            set sp to position of sa
                            repeat with outlineElem in (UI elements of sa)
                                try
                                    set outlineRole to (role of outlineElem) as text
                                    set shouldScanOutline to false
                                    if outlineRole is "AXOutline" then
                                        set shouldScanOutline to true
                                    end if
                                    if outlineRole is "AXTable" then
                                        set shouldScanOutline to true
                                    end if
                                    if shouldScanOutline then
                                        set rowCount to count of rows of outlineElem
                                        repeat with rowIndex from 1 to rowCount
                                            try
                                                set targetRow to row rowIndex of outlineElem
                                                set labelText to my rowLabel(targetRow)
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
                                                    if wantedScrollLeft is -1 then
                                                        if preferRightmost then
                                                            set rowLeft to item 1 of sp
                                                            if rowLeft > bestScore then
                                                                set bestScore to rowLeft
                                                                set bestRow to targetRow
                                                            end if
                                                        else
                                                            if my selectRow(targetRow) then return "OK"
                                                        end if
                                                    else
                                                        set rowLeft to item 1 of sp
                                                        set distanceValue to rowLeft - wantedScrollLeft
                                                        if distanceValue < 0 then set distanceValue to -distanceValue
                                                        if distanceValue < bestScore then
                                                            set bestScore to distanceValue
                                                            set bestRow to targetRow
                                                        end if
                                                    end if
                                                end if
                                            end try
                                        end repeat
                                    end if
                                end try
                            end repeat
                        end try
                    end repeat
                end if
            end try
        end repeat

        if bestRow is not missing value then
            if my selectRow(bestRow) then return "OK"
        end if
        return ""
end tell

on selectRow(targetRow)
    try
        set value of attribute "AXSelected" of targetRow to true
        return true
    end try
    try
        perform action "AXPress" of targetRow
        return true
    end try
    try
        click targetRow
        return true
    end try
    try
        set rowPosition to position of targetRow
        set rowSize to size of targetRow
        set clickX to (item 1 of rowPosition) + 20
        set clickY to (item 2 of rowPosition) + ((item 2 of rowSize) / 2)
        click at {{clickX, clickY}}
        return true
    end try
    return false
end selectRow

on rowLabel(targetRow)
    set labelText to ""
    try
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
    end try
    if labelText is "" then
        repeat with c in (UI elements of targetRow)
            try
                if role of c is "AXStaticText" then
                    set labelText to my cleanText(value of c as text)
                    if labelText is not "" then exit repeat
                end if
            end try
            try
                if labelText is "" then
                    set labelText to my cleanText(value of attribute "AXDescription" of c as text)
                    if labelText is not "" then exit repeat
                end if
            end try
        end repeat
    end if
    return labelText
end rowLabel

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
    return _run_osascript(script, timeout=8).strip() == "OK"


def select_page_row_by_text(window: MacWindow, text: str) -> bool:
    return _select_outline_row_by_text(
        window,
        text,
        prefer_rightmost=True,
    )

_publish_context(globals())

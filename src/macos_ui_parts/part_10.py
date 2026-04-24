# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _open_tab_web_url_from_description(name: str, description: str) -> str:
    clean_name = _clean_field(name)
    clean_desc = _clean_field(description)
    if not clean_name or "d.docs.live.net" not in clean_desc.casefold():
        return ""
    parts = [_clean_field(part) for part in clean_desc.split("»")]
    if len(parts) < 2:
        return ""
    account_id = ""
    for part in parts:
        token = part.strip().strip(",")
        if token and all(ch in "0123456789abcdefABCDEF" for ch in token) and len(token) >= 8:
            account_id = token.lower()
            break
    if not account_id:
        return ""
    folder_parts = [part for part in parts[2:] if part]
    if not folder_parts:
        folder_parts = ["문서"]
    path = "/".join(quote(part, safe="-_.~") for part in [*folder_parts, clean_name])
    return f"https://d.docs.live.net/{account_id}/{path}"


def _open_tab_notebook_entries_from_ax(window: Optional[MacWindow]) -> List[Dict[str, str]]:
    roots = _ax_onenote_notebook_dialog_roots(window)
    if not roots:
        return []
    entries: List[Dict[str, str]] = []
    seen = set()
    deadline = time.monotonic() + 5.0
    node_count = 0

    def _append_static_text(element: c_void_p) -> None:
        value = _clean_field(_ax_text_attribute(element, "AXValue"))
        desc = _clean_field(_ax_text_attribute(element, "AXDescription"))
        key = _normalize_text(value)
        if not key or key in seen:
            return
        if key in {_normalize_text("문서"), _normalize_text("Documents"), _normalize_text("최근 폴더")}:
            return
        web_url = _open_tab_web_url_from_description(value, desc)
        if not web_url:
            return
        seen.add(key)
        entries.append(
            {
                "name": value,
                "url": web_url,
                "path": desc,
            }
        )

    def _visit(element: c_void_p, depth: int) -> None:
        nonlocal node_count
        if not element or depth > 16 or node_count >= 2600 or time.monotonic() > deadline:
            return
        node_count += 1
        role = _ax_text_attribute(element, "AXRole")
        if role == "AXStaticText":
            _append_static_text(element)

        children = _ax_array_attribute(element, "AXChildren")
        try:
            for child in children:
                _visit(child, depth + 1)
        finally:
            _release_ax_refs(children)

    try:
        for root in roots:
            _visit(root, 0)
            if entries:
                break
    finally:
        _release_ax_refs(roots)
    return entries


def _ensure_open_tab_notebooks_dialog(
    window: MacWindow,
    *,
    fast: bool = False,
) -> Tuple[bool, bool]:
    opened_sidebar = False
    dialog_roots = _ax_onenote_notebook_dialog_roots(window)
    dialog_already_open = bool(dialog_roots)
    _release_ax_refs(dialog_roots)
    if not dialog_already_open:
        ready, opened_sidebar = _open_recent_notebooks_dialog_with_state(
            window,
            fast=fast,
        )
        if not ready:
            return False, opened_sidebar

    if not _ax_press_dialog_radio(
        window,
        ("열기", "Open"),
        timeout_sec=3.0 if fast else 5.0,
    ):
        return False, opened_sidebar

    deadline = time.monotonic() + (6.0 if fast else 10.0)
    clicked_documents = False
    while time.monotonic() < deadline:
        entries = _open_tab_notebook_entries_from_ax(window)
        if len(entries) >= 5:
            return True, opened_sidebar
        if not clicked_documents:
            clicked_documents = _ax_click_open_tab_documents_folder(window)
        time.sleep(0.25)
    return bool(_open_tab_notebook_entries_from_ax(window)), opened_sidebar


def open_tab_notebook_records(
    window: Optional[MacWindow],
    *,
    fast: bool = False,
) -> List[Dict[str, Any]]:
    if window is None:
        return []
    ready, opened_sidebar = _ensure_open_tab_notebooks_dialog(window, fast=fast)
    if not ready:
        return []
    try:
        return [
            {
                "name": entry["name"],
                "url": entry["url"],
                "path": entry.get("path", ""),
                "last_accessed_at": 0,
                "source": "MAC_OPEN_TAB",
            }
            for entry in _open_tab_notebook_entries_from_ax(window)
            if str(entry.get("name") or "").strip() and str(entry.get("url") or "").strip()
        ]
    finally:
        dismiss_recent_notebooks_dialog(window)
        _restore_notebook_sidebar(window, opened_sidebar)


def _is_notebook_sidebar_snapshot(snapshot: Dict[str, Any]) -> bool:
    rows = snapshot.get("rows") or []
    if not rows:
        return False
    scroll_rects = {
        tuple(row.get("scroll_rect") or ())
        for row in rows
        if row.get("scroll_rect")
    }
    return len(scroll_rects) == 1


def _press_window_button(
    window: MacWindow,
    *,
    help_contains: Optional[List[str]] = None,
    value_equals: Optional[List[str]] = None,
    timeout: int = 15,
) -> bool:
    locator = _applescript_window_locator(window.process_id(), window.window_text())
    help_matchers = [str(item or "").strip() for item in (help_contains or []) if str(item or "").strip()]
    value_matchers = [str(item or "").strip() for item in (value_equals or []) if str(item or "").strip()]
    help_tokens = ", ".join(_quote_applescript_text(item) for item in help_matchers) or ""
    value_tokens = ", ".join(_quote_applescript_text(item) for item in value_matchers) or ""
    script = locator + f'''
        set helpMatchers to {{{help_tokens}}}
        set valueMatchers to {{{value_tokens}}}
        set allElems to entire contents of targetWindow
        repeat with e in allElems
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
                repeat with wantedHelp in helpMatchers
                    if wantedHelp is not "" and helpText contains wantedHelp then
                        set shouldPress to true
                        exit repeat
                    end if
                end repeat

                if shouldPress is false then
                    repeat with wantedValue in valueMatchers
                        if wantedValue is not "" and (titleText is wantedValue or nameText is wantedValue or descText is wantedValue or valueText is wantedValue) then
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
    return _run_osascript(script, timeout=timeout).strip() == "OK"


def _press_notebook_sidebar_button(
    window: MacWindow,
    *,
    open_sidebar: bool,
    timeout: int = 4,
) -> bool:
    """Press the macOS OneNote notebook-list toggle without scanning the full UI tree."""
    locator = _applescript_window_locator(window.process_id(), window.window_text())
    if open_sidebar:
        help_matchers = [
            "전자 필기장 보기, 만들기 또는 열기",
            "View, create, or open notebooks",
            "View, create or open notebooks",
        ]
    else:
        help_matchers = [
            "전자 필기장 목록 숨기기",
            "Hide notebook list",
        ]
    help_tokens = ", ".join(_quote_applescript_text(item) for item in help_matchers)
    script = locator + f'''
        set helpMatchers to {{{help_tokens}}}
        try
            set targetSplitGroup to first UI element of targetWindow whose role is "AXSplitGroup"
            set leftGroup to first UI element of targetSplitGroup whose role is "AXGroup"
            set nestedSplitGroup to first UI element of leftGroup whose role is "AXSplitGroup"
            set scanItems to UI elements of nestedSplitGroup
        on error
            set scanItems to UI elements of targetWindow
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
                repeat with wantedHelp in helpMatchers
                    if wantedHelp is not "" and helpText contains wantedHelp then
                        try
                            perform action "AXPress" of e
                        on error
                            click e
                        end try
                        return "OK"
                    end if
                end repeat
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

_publish_context(globals())

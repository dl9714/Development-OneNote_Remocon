# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



class MacRow:
    __slots__ = ("window", "text", "selected", "rect", "scroll_rect", "order", "detail")

    def __init__(
        self,
        window: MacWindow,
        text: str,
        selected: bool,
        rect: MacRect,
        scroll_rect: MacRect,
        order: int,
        detail: str = "",
    ):
        self.window = window
        self.text = text
        self.selected = selected
        self.rect = rect
        self.scroll_rect = scroll_rect
        self.order = order
        self.detail = detail

    def window_text(self) -> str:
        return self.text

    def rectangle(self) -> MacRect:
        return self.rect

    def is_selected(self) -> bool:
        return bool(self.selected)

    def has_keyboard_focus(self) -> bool:
        return bool(self.selected)

    def select(self) -> None:
        select_row_by_text(self.window, self.text, preferred_scroll_left=self.scroll_rect.left)

    def click_input(self) -> None:
        self.select()


class MacTreeControl:
    """선택/검색/정렬 용도로 쓰는 단순 트리 래퍼."""

    def __init__(self, window: MacWindow):
        self.window = window

    def wrapper_object(self):
        return self

    def children(self) -> List[MacRow]:
        return collect_onenote_rows(self.window)

    def descendants(self, control_type=None) -> List[MacRow]:
        if control_type in (None, "TreeItem", "ListItem"):
            return self.children()
        return []

    def get_selection(self) -> List[MacRow]:
        return [row for row in self.children() if row.selected]

    def set_focus(self) -> None:
        self.window.set_focus()


def _extract_current_notebook_name(raw_text: str) -> str:
    text = _clean_field(raw_text)
    for marker in ("(현재 전자필기장)", "(현재 전자 필기장)"):
        if marker in text:
            text = text.split(marker, 1)[0].strip()
            break
    if "," in text:
        text = text.split(",", 1)[0].strip()
    return text


def collect_onenote_snapshot(window: MacWindow, timeout: int = 20) -> Dict[str, Any]:
    locator = _applescript_window_locator(window.process_id(), window.window_text())
    script = locator + r'''
        set resultItems to {}
        set currentNotebook to ""
        set allElems to entire contents of targetWindow

        repeat with e in allElems
            try
                if role of e is "AXButton" then
                    set candidateText to ""
                    try
                        set candidateText to description of e as text
                    end try
                    if candidateText does not contain "(현재 전자필기장)" and candidateText does not contain "(현재 전자 필기장)" then
                        try
                            set candidateText to value of e as text
                        end try
                    end if
                    if candidateText does not contain "(현재 전자필기장)" and candidateText does not contain "(현재 전자 필기장)" then
                        try
                            set candidateText to title of e as text
                        end try
                    end if
                    set candidateText to my cleanText(candidateText)
                    if candidateText contains "(현재 전자필기장)" or candidateText contains "(현재 전자 필기장)" then
                        set currentNotebook to candidateText
                    end if
                end if
            end try
        end repeat

        if currentNotebook is not "" then
            set end of resultItems to "NOTEBOOK" & tab & my cleanText(currentNotebook)
        end if

        repeat with e in allElems
            try
                if role of e is "AXScrollArea" then
                    set pp to position of e
                    set ss to size of e
                    set vx to ""
                    repeat with b in (every UI element of e)
                        try
                            if role of b is "AXScrollBar" then
                                try
                                    set bo to orientation of b as text
                                on error
                                    set bo to ""
                                end try
                                if bo is "AXVerticalOrientation" then
                                    try
                                        set vx to value of b as text
                                    on error
                                        set vx to ""
                                    end try
                                end if
                            end if
                        end try
                    end repeat
                    set end of resultItems to "SCROLL" & tab & (item 1 of pp) & tab & (item 2 of pp) & tab & (item 1 of ss) & tab & (item 2 of ss) & tab & vx
                end if
            end try
        end repeat

        set rowIndex to 0
        repeat with e in allElems
            try
                if role of e is "AXRow" then
                    set rowIndex to rowIndex + 1
                    set labelText to ""
                    set detailText to ""
                    set childElems to entire contents of e
                    repeat with c in childElems
                        try
                            if role of c is "AXStaticText" then
                                set labelText to my cleanText(value of c as text)
                                exit repeat
                            end if
                        end try
                    end repeat
                    if labelText is "" then
                        repeat with c in childElems
                            try
                                set maybeText to my cleanText(description of c as text)
                                if maybeText is not "" then
                                    set labelText to maybeText
                                    exit repeat
                                end if
                            end try
                            try
                                set maybeText to my cleanText(title of c as text)
                                if maybeText is not "" then
                                    set labelText to maybeText
                                    exit repeat
                                end if
                            end try
                        end repeat
                    end if

                    repeat with c in childElems
                        try
                            set maybeDetail to my cleanText(description of c as text)
                            if maybeDetail is not "" and maybeDetail is not labelText then
                                set detailText to maybeDetail
                                exit repeat
                            end if
                        end try
                        try
                            set maybeDetail to my cleanText(title of c as text)
                            if maybeDetail is not "" and maybeDetail is not labelText then
                                set detailText to maybeDetail
                                exit repeat
                            end if
                        end try
                    end repeat

                    try
                        set selectedState to value of attribute "AXSelected" of e as text
                    on error
                        set selectedState to "false"
                    end try
                    set rp to position of e
                    set rs to size of e
                    set scrollRect to ""
                    set parentElem to e
                    repeat 12 times
                        try
                            set parentElem to parent of parentElem
                            if parentElem is missing value then exit repeat
                            if role of parentElem is "AXScrollArea" then
                                set sp to position of parentElem
                                set ss to size of parentElem
                                set scrollRect to (item 1 of sp) & tab & (item 2 of sp) & tab & (item 1 of ss) & tab & (item 2 of ss)
                                exit repeat
                            end if
                        on error
                            exit repeat
                        end try
                    end repeat

                    set end of resultItems to "ROW" & tab & rowIndex & tab & my cleanText(labelText) & tab & my cleanText(detailText) & tab & selectedState & tab & (item 1 of rp) & tab & (item 2 of rp) & tab & (item 1 of rs) & tab & (item 2 of rs) & tab & scrollRect
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
    snapshot: Dict[str, Any] = {"current_notebook": "", "scrolls": [], "rows": []}
    if not raw:
        return snapshot

    for line in raw.splitlines():
        parts = line.split("\t")
        if not parts:
            continue
        kind = parts[0]
        if kind == "NOTEBOOK" and len(parts) >= 2:
            snapshot["current_notebook"] = _extract_current_notebook_name(parts[1])
            continue
        if kind == "SCROLL" and len(parts) >= 6:
            try:
                x, y, w, h = map(int, parts[1:5])
                value = float(parts[5]) if parts[5] not in ("", "missing value") else None
            except Exception:
                continue
            snapshot["scrolls"].append(
                {
                    "rect": _rect_key(x, y, w, h),
                    "value": value,
                }
            )
            continue
        if kind == "ROW" and len(parts) >= 9:
            try:
                order = int(parts[1])
                text = _clean_field(parts[2])
                detail = _clean_field(parts[3]) if len(parts) >= 4 else ""
                selected = (parts[4].strip().lower() == "true") if len(parts) >= 5 else False
                x, y, w, h = map(int, parts[5:9])
                if len(parts) >= 13 and all(parts[idx].strip() for idx in range(9, 13)):
                    sx, sy, sw, sh = map(int, parts[9:13])
                    scroll_rect = _rect_key(sx, sy, sw, sh)
                else:
                    scroll_rect = None
            except Exception:
                continue
            if not text:
                continue
            snapshot["rows"].append(
                {
                    "order": order,
                    "text": text,
                    "detail": detail,
                    "selected": selected,
                    "rect": _rect_key(x, y, w, h),
                    "scroll_rect": scroll_rect,
                }
            )
    scrolls = snapshot.get("scrolls", [])
    for row in snapshot.get("rows", []):
        if row.get("scroll_rect"):
            continue
        x, y, w, h = row["rect"]
        center_x = x + (w / 2.0)
        center_y = y + (h / 2.0)
        matches = []
        for scroll in scrolls:
            sx, sy, sw, sh = scroll["rect"]
            if sx <= center_x <= sx + sw and sy <= center_y <= sy + sh:
                matches.append(scroll["rect"])
        if matches:
            row["scroll_rect"] = sorted(matches, key=lambda rect: rect[2] * rect[3])[0]
    return snapshot


def collect_onenote_rows(window: MacWindow) -> List[MacRow]:
    snapshot = collect_onenote_snapshot(window)
    rows: List[MacRow] = []
    for item in snapshot.get("rows", []):
        x, y, w, h = item["rect"]
        scroll_rect = item.get("scroll_rect") or item["rect"]
        sx, sy, sw, sh = scroll_rect
        rows.append(
            MacRow(
                window=window,
                text=item["text"],
                detail=item.get("detail") or "",
                selected=bool(item.get("selected")),
                rect=MacRect(x, y, x + w, y + h),
                scroll_rect=MacRect(sx, sy, sx + sw, sy + sh),
                order=int(item["order"]),
            )
        )
    return rows


def pick_selected_row(window: MacWindow, prefer_leftmost: bool = True) -> Optional[MacRow]:
    selected = [row for row in collect_onenote_rows(window) if row.selected]
    if not selected:
        return None
    return sorted(
        selected,
        key=lambda row: (
            row.scroll_rect.left if prefer_leftmost else -row.scroll_rect.left,
            row.order,
        ),
    )[0]


def pick_selected_page_row(window: MacWindow) -> Optional[MacRow]:
    rows = collect_onenote_rows(window)
    selected = [row for row in rows if row.selected]
    if not selected:
        return None
    scroll_lefts = sorted({row.scroll_rect.left for row in rows})
    if len(scroll_lefts) < 2:
        return None
    page_scroll_left = scroll_lefts[-1]
    page_rows = [row for row in selected if row.scroll_rect.left == page_scroll_left]
    if not page_rows:
        return None
    return sorted(page_rows, key=lambda row: row.order)[0]

_publish_context(globals())

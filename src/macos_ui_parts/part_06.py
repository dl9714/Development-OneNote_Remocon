# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def current_outline_context(window: MacWindow) -> Dict[str, str]:
    locator = _applescript_window_locator(window.process_id(), window.window_text())
    script = locator + r'''
        set notebookName to ""
        set selectedRows to {}
        set allElems to entire contents of targetWindow

        repeat with e in allElems
            try
                if role of e is "AXButton" and notebookName is "" then
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
                        set notebookName to candidateText
                    end if
                end if

                if role of e is "AXRow" then
                    try
                        set selectedState to value of attribute "AXSelected" of e as boolean
                    on error
                        set selectedState to false
                    end try
                    if selectedState is true then
                        set labelText to ""
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
                        if labelText is not "" then
                            set rp to position of e
                            set end of selectedRows to ((item 1 of rp) as text) & tab & labelText
                        end if
                    end if
                end if
            end try
        end repeat

        set resultItems to {}
        if notebookName is not "" then
            set end of resultItems to "NOTEBOOK" & tab & my cleanText(notebookName)
        end if
        repeat with rowInfo in selectedRows
            set end of resultItems to "ROW" & tab & rowInfo
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
    raw = _run_osascript(script, timeout=12)
    notebook_name = str(window.window_text() or "").strip()
    selected_rows: List[Tuple[int, str]] = []

    for line in (raw or "").splitlines():
        parts = line.split("\t")
        if not parts:
            continue
        if parts[0] == "NOTEBOOK" and len(parts) >= 2:
            notebook_name = _extract_current_notebook_name(parts[1])
            continue
        if parts[0] == "ROW" and len(parts) >= 3:
            try:
                selected_rows.append((int(parts[1]), _clean_field(parts[2])))
            except Exception:
                continue

    selected_rows = [(x, text) for x, text in selected_rows if text]
    selected_rows.sort(key=lambda item: item[0])
    section_name = selected_rows[0][1] if selected_rows else ""
    page_name = selected_rows[-1][1] if len(selected_rows) >= 2 else ""

    return {
        "notebook": str(notebook_name or "").strip(),
        "section": str(section_name or "").strip(),
        "page": str(page_name or "").strip(),
    }


def center_selected_row(
    window: MacWindow,
    prefer_leftmost: bool = True,
    target_text: str = "",
) -> Tuple[bool, Optional[str]]:
    for _ in range(3):
        snapshot = collect_onenote_snapshot(window)
        rows = snapshot.get("rows", [])
        selected_rows = [row for row in rows if row.get("selected")]
        if not selected_rows:
            return False, None

        wanted_key = _normalize_text(target_text)
        if wanted_key:
            target_rows = [
                row
                for row in selected_rows
                if _normalize_text(row.get("text") or "") == wanted_key
            ]
            if target_rows:
                selected_rows = target_rows

        target = sorted(
            selected_rows,
            key=lambda row: (
                row["scroll_rect"][0] if prefer_leftmost else -row["scroll_rect"][0],
                row["order"],
            ),
        )[0]

        sx, sy, sw, sh = target["scroll_rect"]
        x, y, w, h = target["rect"]
        scroll = next(
            (
                item for item in snapshot.get("scrolls", [])
                if tuple(item.get("rect") or ()) == (sx, sy, sw, sh)
            ),
            None,
        )
        row_center = y + (h / 2.0)
        scroll_center = sy + (sh / 2.0)
        delta = row_center - scroll_center
        if abs(delta) <= 12:
            return True, target["text"]

        if not scroll or scroll.get("value") is None:
            return True, target["text"]

        group_rows = [row for row in rows if tuple(row.get("scroll_rect") or ()) == (sx, sy, sw, sh)]
        if not group_rows:
            return False, target["text"]

        min_top = min(row["rect"][1] for row in group_rows)
        max_bottom = max((row["rect"][1] + row["rect"][3]) for row in group_rows)
        content_height = max_bottom - min_top
        scrollable_height = max(1.0, float(content_height - sh))
        if scrollable_height <= 1.0:
            return True, target["text"]

        current_value = float(scroll["value"])
        target_value = max(0.0, min(1.0, current_value + (delta / scrollable_height)))
        if abs(target_value - current_value) < 0.002:
            return True, target["text"]

        locator = _applescript_window_locator(window.process_id(), window.window_text())
        script = locator + f'''
            set targetScrollLeft to {sx}
            set targetScrollTop to {sy}
            set targetValue to {target_value}
            set allElems to entire contents of targetWindow
            repeat with e in allElems
                try
                    if role of e is "AXScrollArea" then
                        set pp to position of e
                        if (item 1 of pp) is targetScrollLeft and (item 2 of pp) is targetScrollTop then
                            repeat with b in (every UI element of e)
                                try
                                    if role of b is "AXScrollBar" then
                                        try
                                            set bo to orientation of b as text
                                        on error
                                            set bo to ""
                                        end try
                                        if bo is "AXVerticalOrientation" then
                                            set value of b to targetValue
                                            return "OK"
                                        end if
                                    end if
                                end try
                            end repeat
                        end if
                    end if
                end try
            end repeat
            return ""
end tell
'''
        if _run_osascript(script, timeout=15).strip() != "OK":
            return False, target["text"]
        time.sleep(0.12)

    final_row = pick_selected_row(window, prefer_leftmost=prefer_leftmost)
    return (final_row is not None), (final_row.text if final_row else None)


def list_current_notebook_targets(window: MacWindow) -> List[Dict[str, str]]:
    snapshot = collect_onenote_snapshot(window)
    notebook_name = snapshot.get("current_notebook") or window.window_text()
    rows = snapshot.get("rows", [])
    if not rows:
        return [
            {
                "kind": "notebook",
                "name": f"전자필기장 - {notebook_name}",
                "path": notebook_name,
                "notebook": notebook_name,
                "section_group": "",
                "section": "",
                "section_group_id": "",
                "section_id": "",
            }
        ]

    leftmost_scroll = min(row["scroll_rect"][0] for row in rows)
    sections = []
    seen = set()
    for row in rows:
        if row["scroll_rect"][0] != leftmost_scroll:
            continue
        section_name = str(row.get("text") or "").strip()
        if not section_name:
            continue
        key = _normalize_text(section_name)
        if key in seen:
            continue
        seen.add(key)
        sections.append(
            {
                "kind": "section",
                "name": f"섹션 - {section_name}",
                "path": f"{notebook_name} > {section_name}",
                "notebook": notebook_name,
                "section_group": "",
                "section": section_name,
                "section_group_id": "",
                "section_id": "",
            }
        )

    base = [
        {
            "kind": "notebook",
            "name": f"전자필기장 - {notebook_name}",
            "path": notebook_name,
            "notebook": notebook_name,
            "section_group": "",
            "section": "",
            "section_group_id": "",
            "section_id": "",
        }
    ]
    return base + sections


def _read_open_notebook_names_from_plist() -> List[str]:
    if not _MAC_ONENOTE_NOTEBOOKS_PLIST.is_file():
        return []
    try:
        import plistlib

        with _MAC_ONENOTE_NOTEBOOKS_PLIST.open("rb") as f:
            data = plistlib.load(f)
    except Exception:
        return []

    if not isinstance(data, list):
        return []

    names: List[str] = []
    seen = set()
    for item in data:
        if not isinstance(item, dict):
            continue
        if int(item.get("Type") or 0) != 1:
            continue
        name = _clean_field(str(item.get("Name") or ""))
        key = _normalize_text(name)
        if not key or key in seen:
            continue
        seen.add(key)
        names.append(name)
    return names


def _read_open_notebook_names_from_plist_with_timeout(
    timeout_sec: float = 1.5,
) -> Optional[List[str]]:
    box: Dict[str, Any] = {}
    done = threading.Event()

    def _runner() -> None:
        try:
            box["value"] = _read_open_notebook_names_from_plist()
        except Exception as exc:
            box["error"] = exc
        finally:
            done.set()

    worker = threading.Thread(target=_runner, daemon=True)
    worker.start()
    if not done.wait(max(0.2, float(timeout_sec or 0.0))):
        return None
    if "error" in box:
        return []
    value = box.get("value")
    return value if isinstance(value, list) else []

_publish_context(globals())

# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


def _quick_open_notebook_key_set(window: Optional[MacWindow]) -> set:
    reader = globals().get("current_open_notebook_names_quick")
    if not callable(reader):
        return set()
    try:
        result = reader(
            window, ax_timeout_sec=0.0, plist_timeout_sec=0.18,
            sidebar_timeout_sec=0.0, min_names_before_sidebar=1,
        )
    except Exception:
        return set()
    return {
        _normalize_text(name)
        for name in (result.get("names") if isinstance(result, dict) else []) or []
        if _normalize_text(name)
    }


def _ax_click_notebook_sidebar_row(
    window: Optional[MacWindow],
    notebook_name: str,
    known_notebook_keys=None,
) -> bool:
    wanted_key = _normalize_text(notebook_name)
    if not (IS_MACOS and _APP_SERVICES and window is not None and wanted_key):
        return False
    known_keys = {str(key or "") for key in (known_notebook_keys or set()) if str(key or "")}

    roots = _ax_window_roots_for_onenote(window)
    if not roots:
        return False

    current_title = _clean_field(str(getattr(window, "info", {}).get("title") or ""))
    if not current_title:
        try:
            current_title = _clean_field(window.window_text())
        except Exception:
            current_title = ""
    current_key = _normalize_text(current_title)
    best = {"score": -1, "row": None}
    state = {"nodes": 0, "order": 0}
    deadline = time.monotonic() + 2.2

    def _maybe_keep_row(row: c_void_p, score: int) -> None:
        if score <= int(best.get("score") or -1):
            return
        retained = _cf_retain(row)
        if not retained:
            return
        old_row = best.get("row")
        if old_row:
            _cf_release(old_row)
        best["score"] = int(score)
        best["row"] = retained

    def _visit(element: c_void_p, depth: int) -> None:
        if not element or depth > 18 or state["nodes"] >= 1800 or time.monotonic() > deadline:
            return
        state["nodes"] += 1
        role = _ax_text_attribute(element, "AXRole")
        if role == "AXOutline":
            state["order"] += 1
            order = int(state["order"])
            rows = _ax_array_attribute(element, "AXRows")
            if not rows:
                rows = _ax_array_attribute(element, "AXChildren")
            try:
                names = []
                matched_row = None
                exact_match = False
                for row in rows:
                    name = _ax_row_notebook_name(row)
                    name_key = _normalize_text(name)
                    if not name_key:
                        continue
                    names.append(name_key)
                    if (
                        matched_row is None
                        and (
                            name_key == wanted_key
                            or wanted_key in name_key
                            or name_key in wanted_key
                        )
                    ):
                        matched_row = row
                        exact_match = name_key == wanted_key
                if matched_row is not None:
                    keys = set(names)
                    overlap = len(keys.intersection(known_keys)) if known_keys else 0
                    plausible = (
                        current_key in keys
                        or overlap >= 2
                        or (not known_keys and len(keys) >= 8)
                    )
                    if plausible:
                        score = min(len(keys), 80) - (order * 3) - depth
                        score += 300 if current_key and current_key in keys else 0
                        score += min(overlap, 20) * 15
                        score += 80 if exact_match else 35
                        _maybe_keep_row(matched_row, score)
            finally:
                _release_ax_refs(rows)

        children = _ax_array_attribute(element, "AXChildren")
        try:
            for child in children:
                _visit(child, depth + 1)
        finally:
            _release_ax_refs(children)

    try:
        for root in roots[:3]:
            _visit(root, 0)
    finally:
        _release_ax_refs(roots)
    row = best.get("row")
    try:
        if not row:
            return False
        # A double click is more reliable on OneNote's macOS notebook list:
        # first click selects the row, second click activates it.
        return _ax_click_element_center(row, click_count=2, preferred_x_offset=36.0)
    finally:
        if row:
            _cf_release(row)


def _select_open_notebook_by_name_ax(
    window: Optional[MacWindow],
    notebook_name: str,
    *,
    wait_for_visible: bool = True,
) -> bool:
    wanted_name = _clean_field(notebook_name)
    wanted_key = _normalize_text(wanted_name)
    if not wanted_key or window is None:
        return False

    if wait_for_visible:
        try:
            window.set_focus()
        except Exception:
            try:
                subprocess.run(
                    ["/usr/bin/open", "-b", ONENOTE_MAC_BUNDLE_ID],
                    capture_output=True,
                    timeout=3,
                )
                time.sleep(0.25)
            except Exception:
                pass

    known_keys = _quick_open_notebook_key_set(window)
    if wanted_key in known_keys and _ax_click_notebook_sidebar_row(
        window,
        wanted_name,
        known_keys,
    ):
        if not wait_for_visible:
            return True
        deadline = time.monotonic() + 4.0
        while time.monotonic() < deadline:
            if _is_target_notebook_visible(window, wanted_name):
                return True
            time.sleep(0.2)
        return True

    names = _read_open_notebook_names_from_ax(window)
    name_keys = {_normalize_text(name) for name in names}
    if wanted_key not in name_keys:
        _ax_press_notebook_sidebar_button(window)
        deadline = time.monotonic() + 2.5
        while time.monotonic() < deadline:
            names = _read_open_notebook_names_from_ax(window)
            name_keys = {_normalize_text(name) for name in names}
            if wanted_key in name_keys:
                break
            time.sleep(0.15)

    if wanted_key not in name_keys:
        _append_macos_debug_log(
            "[AX_SELECT_NOTEBOOK] target_not_visible "
            f"target={wanted_name!r} count={len(names)} "
            f"debug={macos_last_ax_notebook_debug()!r}"
        )
        return False

    if not _ax_click_notebook_sidebar_row(window, wanted_name, name_keys):
        _append_macos_debug_log(
            f"[AX_SELECT_NOTEBOOK] row_click_failed target={wanted_name!r}"
        )
        return False

    if not wait_for_visible:
        return True

    deadline = time.monotonic() + 4.0
    while time.monotonic() < deadline:
        if _is_target_notebook_visible(window, wanted_name):
            return True
        time.sleep(0.2)
    if _is_target_notebook_visible(window, wanted_name):
        return True

    # When the packaged .app is missing Apple Events visibility but still has
    # Accessibility access, the CG click can succeed while title verification
    # remains unreadable. The row was present and clicked, so prefer not to
    # report a false failure in the UI.
    _append_macos_debug_log(
        f"[AX_SELECT_NOTEBOOK] verify_timeout_assume_clicked target={wanted_name!r}"
    )
    return True


def _read_open_notebook_names_from_sidebar(window: Optional[MacWindow]) -> List[str]:
    if window is None:
        return []
    try:
        window.set_focus()
        time.sleep(0.08)
    except Exception:
        pass
    try:
        sidebar_ready, opened_sidebar = _ensure_notebook_sidebar(window)
    except Exception:
        return []
    if not sidebar_ready:
        return []

    script = _applescript_window_locator(window.process_id(), window.window_text()) + r'''
        set resultItems to {}
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
    try:
        raw = _run_osascript(script, timeout=6)
    except Exception:
        raw = ""
    finally:
        _restore_notebook_sidebar(window, opened_sidebar)

    names: List[str] = []
    seen = set()
    for line in (raw or "").splitlines():
        name = _clean_field(line)
        for marker in (", 동기화할 수 없습니다", ", 동기화 중"):
            if marker in name:
                name = name.split(marker, 1)[0].strip()
                break
        key = _normalize_text(name)
        if not key or key in seen:
            continue
        seen.add(key)
        names.append(name)
    return names


def _recent_notebook_records_from_cache() -> List[Dict[str, Any]]:
    if not _MAC_ONENOTE_RESOURCEINFOCACHE_JSON.is_file():
        return []
    try:
        payload = json.loads(_MAC_ONENOTE_RESOURCEINFOCACHE_JSON.read_text(encoding="utf-8"))
    except Exception:
        return []

    entries = payload.get("ResourceInfoCache") or []
    if not isinstance(entries, list):
        return []

    def _is_supported_recent_netloc(netloc: str) -> bool:
        host = str(netloc or "").strip().lower()
        if not host:
            return False
        if host in {"d.docs.live.net", "onedrive.live.com", "1drv.ms"}:
            return True
        if host.endswith(".sharepoint.com") or host.endswith(".sharepoint-df.com"):
            return True
        return False

    records: List[Dict[str, Any]] = []
    seen = set()
    domain_counts: Dict[str, int] = {}
    accepted_domains: Dict[str, int] = {}
    for item in sorted(
        (entry for entry in entries if isinstance(entry, dict)),
        key=lambda entry: int(entry.get("LastAccessedAt") or 0),
        reverse=True,
    ):
        raw_url = str(item.get("Url") or "").strip()
        if not raw_url:
            continue
        parsed = urlparse(raw_url)
        if parsed.scheme.lower() not in {"http", "https"}:
            continue
        host = parsed.netloc.lower()
        domain_counts[host] = int(domain_counts.get(host) or 0) + 1
        if not _is_supported_recent_netloc(host):
            continue
        accepted_domains[host] = int(accepted_domains.get(host) or 0) + 1
        notebook_name = _clean_field(str(item.get("Title") or item.get("title") or ""))
        if not notebook_name:
            notebook_name = _clean_field(unquote(parsed.path.rstrip("/").split("/")[-1]))
        key = _normalize_text(notebook_name)
        if not key or key in seen:
            continue
        seen.add(key)
        records.append(
            {
                "name": notebook_name,
                "url": raw_url,
                "last_accessed_at": int(item.get("LastAccessedAt") or 0),
            }
        )
    if not records or accepted_domains:
        ordered_domains = sorted(
            domain_counts.items(),
            key=lambda item: (-int(item[1]), item[0]),
        )
        ordered_accepted = sorted(
            accepted_domains.items(),
            key=lambda item: (-int(item[1]), item[0]),
        )
        _append_macos_debug_log(
            "[DBG][MAC][RECENT_CACHE] "
            f"records={len(records)} "
            f"domains={ordered_domains[:8]!r} "
            f"accepted={ordered_accepted[:8]!r} "
            f"sample={[record.get('name') for record in records[:8]]!r}"
        )
    return records

_publish_context(globals())

# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


_ROW_TEXT_ATTRS = ("AXValue", "AXTitle", "AXDescription")
_NOTEBOOK_MARKERS = ("(현재 전자필기장)", "(현재 전자 필기장)")
_FAST_OUTLINE_CACHE: Dict[str, Any] = {
    "key": None,
    "timestamp": 0.0,
    "snapshot": None,
}


def _fast_outline_cache_key(window: MacWindow) -> Tuple[int, int, str]:
    return (
        int(window.process_id() or 0),
        int(getattr(window, "info", {}).get("window_number") or 0),
        str(window.window_text() or ""),
    )


def _clone_fast_outline(snapshot: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "notebook": str(snapshot.get("notebook") or ""),
        "rows": list(snapshot.get("rows") or []),
    }


def _fast_outline_has_context(result: Dict[str, Any]) -> bool:
    return bool(result.get("notebook")) and len(result.get("rows") or []) >= 2


def _ax_rect(element) -> Optional[MacRect]:
    point = _ax_point_attribute(element, "AXPosition")
    size = _ax_size_attribute(element, "AXSize")
    if not point or not size:
        return None
    return MacRect(
        int(float(point.x)),
        int(float(point.y)),
        int(float(point.x + size.width)),
        int(float(point.y + size.height)),
    )


def _ax_first_text(element, depth: int = 0, max_depth: int = 4) -> str:
    for attr_name in _ROW_TEXT_ATTRS:
        text = _clean_field(_ax_text_attribute(element, attr_name))
        if text:
            return text
    if depth >= max_depth:
        return ""

    children = _ax_array_attribute(element, "AXChildren")
    try:
        for child in children[:80]:
            text = _ax_first_text(child, depth + 1, max_depth)
            if text:
                return text
    finally:
        _release_ax_refs(children)
    return ""


def _current_notebook_from_button(element) -> str:
    for attr_name in ("AXDescription", "AXValue", "AXTitle"):
        text = _clean_field(_ax_text_attribute(element, attr_name))
        if any(marker in text for marker in _NOTEBOOK_MARKERS):
            return _extract_current_notebook_name(text)
    return ""


def _append_selected_outline_rows(
    result: Dict[str, Any],
    window: MacWindow,
    outline,
    scroll_rect: Optional[MacRect],
    state: Dict[str, int],
) -> bool:
    selected_rows = _ax_array_attribute(outline, "AXSelectedRows")
    before_count = len(result.get("rows") or [])
    try:
        for row in selected_rows:
            rect = _ax_rect(row)
            text = _ax_first_text(row)
            if not (rect and text):
                continue
            state["order"] += 1
            result["rows"].append(
                MacRow(
                    window=window,
                    text=text,
                    selected=True,
                    rect=rect,
                    scroll_rect=scroll_rect or rect,
                    order=int(state["order"]),
                )
            )
    finally:
        _release_ax_refs(selected_rows)
    return len(result.get("rows") or []) > before_count


def _ax_child_at(element, index: int):
    if element is None:
        return None
    children = _ax_array_attribute(element, "AXChildren")
    if not (0 <= index < len(children)):
        _release_ax_refs(children)
        return None
    picked = children[index]
    for child_index, child in enumerate(children):
        if child_index != index:
            _cf_release(child)
    return picked


def _ax_preferred_window_root(window: MacWindow, title_hint: str):
    pid = int(window.process_id() or 0)
    if not pid:
        return None
    app_ref = c_void_p(_APP_SERVICES.AXUIElementCreateApplication(pid))
    if not app_ref:
        return None
    title_key = _normalize_text(title_hint)
    try:
        for attr_name in ("AXFocusedWindow", "AXMainWindow"):
            root = _ax_element_attribute(app_ref, attr_name)
            if not root:
                continue
            root_key = _normalize_text(_ax_text_attribute(root, "AXTitle"))
            if not title_key or root_key == title_key or title_key in root_key:
                return root
            _cf_release(root)
    finally:
        _cf_release(app_ref)
    return None


def _append_direct_group_outline(
    result: Dict[str, Any],
    window: MacWindow,
    group,
    state: Dict[str, int],
) -> bool:
    scroll = _ax_child_at(group, 0)
    outline = _ax_child_at(scroll, 0) if scroll else None
    try:
        if not outline:
            return False
        return _append_selected_outline_rows(
            result,
            window,
            outline,
            _ax_rect(scroll) if scroll else None,
            state,
        )
    finally:
        if outline:
            _cf_release(outline)
        if scroll:
            _cf_release(scroll)


def _walk_direct_outline(window: MacWindow, cache_key: Tuple[int, int, str]) -> Dict[str, Any]:
    result: Dict[str, Any] = {
        "notebook": str(cache_key[2] or "").strip(),
        "rows": [],
    }
    root = _ax_preferred_window_root(window, cache_key[2])
    roots = [root] if root else _ax_window_roots_for_onenote(window)
    state = {"nodes": 0, "order": 0}
    try:
        for root in roots[:1]:
            main_split = _ax_child_at(root, 0)
            left_group = _ax_child_at(main_split, 0) if main_split else None
            nav_split = _ax_child_at(left_group, 1) if left_group else None
            nested_split = _ax_child_at(nav_split, 0) if nav_split else None
            section_group = _ax_child_at(nested_split, 0) if nested_split else None
            page_group = _ax_child_at(nested_split, 2) if nested_split else None
            try:
                for group in (section_group, page_group):
                    if group is not None:
                        _append_direct_group_outline(result, window, group, state)
                if _fast_outline_has_context(result):
                    return result
            finally:
                for ref in (
                    page_group,
                    section_group,
                    nested_split,
                    nav_split,
                    left_group,
                    main_split,
                ):
                    if ref:
                        _cf_release(ref)
    finally:
        _release_ax_refs(roots)
    return result


def _walk_fast_outline(window: MacWindow) -> Dict[str, Any]:
    result: Dict[str, Any] = {"notebook": "", "rows": []}
    if not (IS_MACOS and _APP_SERVICES and window is not None):
        return result

    cache_key = _fast_outline_cache_key(window)
    result["notebook"] = str(cache_key[2] or "").strip()
    now = time.monotonic()
    if _FAST_OUTLINE_CACHE.get("key") == cache_key:
        age = now - float(_FAST_OUTLINE_CACHE.get("timestamp") or 0.0)
        cached = _FAST_OUTLINE_CACHE.get("snapshot")
        if age <= 0.4 and isinstance(cached, dict):
            return _clone_fast_outline(cached)

    direct = _walk_direct_outline(window, cache_key)
    if _fast_outline_has_context(direct):
        _FAST_OUTLINE_CACHE.update(
            {"key": cache_key, "timestamp": now, "snapshot": _clone_fast_outline(direct)}
        )
        return direct

    roots = _ax_window_roots_for_onenote(window)
    state = {"nodes": 0, "order": 0}

    def visit(element, depth: int, scroll_rect: Optional[MacRect]) -> None:
        if depth > 16 or state["nodes"] >= 1800 or _fast_outline_has_context(result):
            return
        state["nodes"] += 1

        role = _ax_text_attribute(element, "AXRole")
        current_scroll = scroll_rect
        if role == "AXScrollArea":
            current_scroll = _ax_rect(element) or scroll_rect
        elif role == "AXButton" and not result["notebook"]:
            result["notebook"] = _current_notebook_from_button(element)
        elif role == "AXOutline" and _append_selected_outline_rows(
            result,
            window,
            element,
            current_scroll,
            state,
        ):
            return
        elif role == "AXRow":
            state["order"] += 1
            if _ax_number_attribute(element, "AXSelected"):
                rect = _ax_rect(element)
                text = _ax_first_text(element)
                if rect and text:
                    result["rows"].append(
                        MacRow(
                            window=window,
                            text=text,
                            selected=True,
                            rect=rect,
                            scroll_rect=current_scroll or rect,
                            order=int(state["order"]),
                        )
                    )

        children = _ax_array_attribute(element, "AXChildren")
        try:
            for child in children:
                visit(child, depth + 1, current_scroll)
                if _fast_outline_has_context(result):
                    break
        finally:
            _release_ax_refs(children)

    try:
        for root in roots:
            visit(root, 0, None)
    finally:
        _release_ax_refs(roots)
    unique_rows = []
    seen_rows = set()
    for row in result["rows"]:
        key = (row.text, row.rect.left, row.rect.top, row.rect.right, row.rect.bottom)
        if key in seen_rows:
            continue
        seen_rows.add(key)
        unique_rows.append(row)
    result["rows"] = unique_rows
    _FAST_OUTLINE_CACHE.update(
        {"key": cache_key, "timestamp": now, "snapshot": _clone_fast_outline(result)}
    )
    return result


def pick_selected_row_fast(
    window: MacWindow,
    prefer_leftmost: bool = True,
) -> Optional[MacRow]:
    rows = _walk_fast_outline(window).get("rows") or []
    if not rows:
        return None
    return sorted(
        rows,
        key=lambda row: (
            row.scroll_rect.left if prefer_leftmost else -row.scroll_rect.left,
            row.rect.top,
            row.order,
        ),
    )[0]


def center_selected_row_fast(
    window: MacWindow,
    prefer_leftmost: bool = True,
    target_text: str = "",
) -> Tuple[bool, Optional[str]]:
    row = pick_selected_row_fast(window, prefer_leftmost=prefer_leftmost)
    if not row:
        return False, None
    wanted_key = _normalize_text(target_text)
    if wanted_key and _normalize_text(row.text) != wanted_key:
        return False, row.text
    return True, row.text


def current_outline_context_fast(window: MacWindow) -> Dict[str, str]:
    snapshot = _walk_fast_outline(window)
    rows = list(snapshot.get("rows") or [])
    rows.sort(key=lambda row: (row.scroll_rect.left, row.rect.left, row.order))
    notebook = str(snapshot.get("notebook") or window.window_text() or "").strip()
    section = rows[0].text if rows else ""
    page = rows[-1].text if len(rows) >= 2 else ""
    return {
        "notebook": notebook,
        "section": str(section or "").strip(),
        "page": str(page or "").strip(),
    }


_publish_context(globals())

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


def _walk_fast_outline(window: MacWindow) -> Dict[str, Any]:
    result: Dict[str, Any] = {"notebook": "", "rows": []}
    if not (IS_MACOS and _APP_SERVICES and window is not None):
        return result

    cache_key = _fast_outline_cache_key(window)
    now = time.monotonic()
    if _FAST_OUTLINE_CACHE.get("key") == cache_key:
        age = now - float(_FAST_OUTLINE_CACHE.get("timestamp") or 0.0)
        cached = _FAST_OUTLINE_CACHE.get("snapshot")
        if age <= 0.4 and isinstance(cached, dict):
            return _clone_fast_outline(cached)

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

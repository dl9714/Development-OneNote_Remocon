# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _mac_ensure_notebook_context_for_section(
    onenote_window,
    notebook_name: str,
    *,
    wait_for_visible: bool = True,
) -> Dict[str, Any]:
    requested_name = _strip_stale_favorite_prefix(str(notebook_name or "").strip())
    result = {
        "ok": True,
        "name": requested_name,
        "source": "",
        "error": "",
    }
    if not IS_MACOS or not requested_name:
        return result

    try:
        onenote_window.set_focus()
    except Exception:
        pass

    if _mac_outline_notebook_matches(onenote_window, requested_name):
        result["source"] = "current"
        return result

    try:
        if mac_select_open_notebook_by_name(
            onenote_window,
            requested_name,
            wait_for_visible=wait_for_visible,
        ):
            _clear_open_notebook_records_cache()
            if wait_for_visible:
                _mac_wait_for_notebook_context(onenote_window, requested_name)
            result["source"] = "open_sidebar"
            return result
    except Exception as e:
        print(f"[WARN][MAC][SECTION_NOTEBOOK][SIDEBAR] {e}")

    record = _mac_find_recent_notebook_record(onenote_window, requested_name)
    if record:
        try:
            if mac_open_recent_notebook_record(
                onenote_window,
                record,
                wait_for_visible=wait_for_visible,
            ):
                _clear_open_notebook_records_cache()
                if wait_for_visible:
                    _mac_wait_for_notebook_context(onenote_window, requested_name)
                result["source"] = "recent"
                return result
        except Exception as e:
            print(f"[WARN][MAC][SECTION_NOTEBOOK][RECENT] {e}")

    candidate_names: List[str] = []
    try:
        candidate_names = _get_open_notebook_names_via_com(refresh=True)
    except Exception:
        candidate_names = []
    result["ok"] = False
    result["error"] = (
        _build_notebook_not_found_error(requested_name, candidate_names)
        + " 섹션 복구를 중단했습니다."
    )
    return result


def _safe_window_text(ctrl) -> str:
    try:
        return ctrl.window_text() or ""
    except Exception:
        return ""


def _safe_control_type(ctrl) -> str:
    if IS_MACOS and ctrl is not None:
        return "TreeItem"
    try:
        return ctrl.element_info.control_type or ""
    except Exception:
        return ""


def _safe_rectangle(ctrl):
    try:
        return ctrl.rectangle()
    except Exception:
        return None


def _make_rect_proxy(left: int, top: int, right: int, bottom: int):
    return SimpleNamespace(left=left, top=top, right=right, bottom=bottom)


def _safe_parent(ctrl):
    try:
        return ctrl.parent()
    except Exception:
        return None


def _find_scrollable_ancestor(ctrl, max_depth: int = 8):
    current = ctrl
    fallback = None
    for _ in range(max(1, max_depth)):
        current = _safe_parent(current)
        if current is None:
            break
        fallback = current
        try:
            if getattr(current, "iface_scroll", None) is not None:
                return current
        except Exception:
            pass
        if _safe_control_type(current) in ("List", "Tree"):
            return current
    return fallback


def _find_center_anchor_element(selected_item):
    item_rect = _safe_rectangle(selected_item)
    if item_rect is None:
        return selected_item, "item_rect_missing"

    item_text = _extract_primary_accessible_text(_safe_window_text(selected_item))
    target_norm = _normalize_text(item_text)
    if not target_norm:
        return selected_item, "item_no_text"

    best = None
    best_score = None
    search_types = ("Text", "ListItem", "TreeItem", "Button", "Custom", "DataItem")
    top_limit = item_rect.top + min(260, max(120, (item_rect.bottom - item_rect.top) // 3))

    for cand in _iter_descendants_by_types(selected_item, search_types):
        if cand is None or cand == selected_item:
            continue
        rect = _safe_rectangle(cand)
        if rect is None:
            continue

        height = rect.bottom - rect.top
        width = rect.right - rect.left
        if height < 14 or height > 120 or width < 60:
            continue
        if rect.top < item_rect.top or rect.top > top_limit:
            continue

        cand_text = _extract_primary_accessible_text(_safe_window_text(cand))
        cand_norm = _normalize_text(cand_text)
        if not cand_norm:
            continue

        exact = 1 if cand_norm == target_norm else 0
        partial = 1 if (target_norm in cand_norm or cand_norm in target_norm) else 0
        if not exact and not partial:
            continue

        top_gap = abs(rect.top - item_rect.top)
        score = (exact, partial, -top_gap, width, -height)
        if best_score is None or score > best_score:
            best = cand
            best_score = score

    if best is not None:
        return best, "descendant_text"

    return selected_item, "item"


def _is_root_notebook_tree_item(selected_item, tree_control) -> bool:
    if selected_item is None or tree_control is None:
        return False

    try:
        if _safe_parent(selected_item) == tree_control:
            return True
    except Exception:
        pass

    try:
        for child in tree_control.children() or []:
            if child == selected_item:
                return True
    except Exception:
        pass

    try:
        return _control_depth_within_tree(selected_item, tree_control) <= 1
    except Exception:
        return False


def _get_scroll_placement_for_selected_item(selected_item, tree_control) -> str:
    if _is_root_notebook_tree_item(selected_item, tree_control):
        return "upper"
    return "center"


def _is_already_well_placed_in_view(
    rect_container,
    rect_item,
    rect_anchor,
    *,
    placement: str,
) -> bool:
    if rect_container is None or rect_item is None:
        return False

    container_height = max(1, rect_container.bottom - rect_container.top)
    top_bias = min(48, max(0, int(container_height * 0.08)))
    bottom_bias = min(24, max(0, int(container_height * 0.04)))
    visible_top = rect_container.top + top_bias
    visible_bottom = rect_container.bottom - bottom_bias

    if placement == "upper":
        anchor_top = rect_anchor.top if rect_anchor is not None else rect_item.top
        upper_band_bottom = visible_top + min(96, max(44, int(container_height * 0.22)))
        return visible_top <= anchor_top <= upper_band_bottom

    item_center_y = (
        (rect_anchor.top + rect_anchor.bottom) / 2
        if rect_anchor is not None
        else (rect_item.top + rect_item.bottom) / 2
    )
    container_center_y = (visible_top + visible_bottom) / 2
    return abs(item_center_y - container_center_y) <= 10


def _resolve_alignment_target_for_selected_item(selected_item, tree_control):
    placement = _get_scroll_placement_for_selected_item(selected_item, tree_control)
    if placement == "upper":
        return selected_item, "item_upper_fast", placement

    anchor_element, anchor_source = _find_center_anchor_element(selected_item)
    return anchor_element, anchor_source, placement


def _wait_until(predicate, timeout: float = 1.5, interval: float = 0.05):
    deadline = time.monotonic() + max(0.05, timeout)
    while time.monotonic() < deadline:
        try:
            value = predicate()
            if value:
                return value
        except Exception:
            pass
        time.sleep(interval)
    return None


def _iter_descendants_by_types(root, control_types):
    for control_type in control_types:
        try:
            for item in root.descendants(control_type=control_type):
                yield item
        except Exception:
            continue


def _extract_primary_accessible_text(text: Any) -> str:
    raw = "" if text is None else str(text)
    raw = raw.replace("\r", "\n")
    parts = [part.strip() for part in raw.split("\n") if part.strip()]
    if not parts:
        return raw.strip()
    for part in parts:
        norm = _normalize_text(part)
        if norm in ("shared with: just me", "shared with just me", "just me"):
            continue
        return part
    return parts[0]


def _find_descendant_by_text(root, texts, control_types=None):
    targets = {_normalize_text(t) for t in (texts or []) if t}
    if not targets:
        return None

    search_types = control_types or (
        "Text",
        "Button",
        "MenuItem",
        "ListItem",
        "TreeItem",
        "Hyperlink",
        "Group",
        "Pane",
        "Custom",
    )
    for item in _iter_descendants_by_types(root, search_types):
        txt = _normalize_text(_safe_window_text(item))
        if not txt:
            continue
        if txt in targets:
            return item
        for target in targets:
            if target and (target in txt or txt in target):
                return item
    return None


def _activate_ui_element(ctrl) -> bool:
    try:
        iface_invoke = getattr(ctrl, "iface_invoke", None)
        if iface_invoke is not None:
            iface_invoke.Invoke()
            return True
    except Exception:
        pass

    try:
        if hasattr(ctrl, "double_click_input"):
            ctrl.double_click_input()
            return True
    except Exception:
        pass

    try:
        if hasattr(ctrl, "click_input"):
            ctrl.click_input()
            return True
    except Exception:
        pass

    try:
        if hasattr(ctrl, "select"):
            ctrl.select()
            keyboard.send_keys("{ENTER}")
            return True
    except Exception:
        pass

    return False


def _is_onenote_open_notebook_view(onenote_window) -> bool:
    title = _find_descendant_by_text(
        onenote_window,
        ONENOTE_OPEN_NOTEBOOK_VIEW_TEXTS,
        control_types=("Text", "Button", "Pane", "Group", "Custom"),
    )
    if title:
        return True

    recent = _find_descendant_by_text(
        onenote_window,
        ONENOTE_RECENT_MENU_TEXTS,
        control_types=("Button", "ListItem", "Text", "Hyperlink"),
    )
    search = _find_descendant_by_text(
        onenote_window,
        ONENOTE_SEARCH_TEXTS,
        control_types=("Edit", "Button", "Text", "Custom"),
    )
    return bool(recent and search)

_publish_context(globals())

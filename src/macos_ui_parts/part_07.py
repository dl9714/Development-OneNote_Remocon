# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _read_notebook_names_with_timeout(
    reader: Callable[[], List[str]],
    timeout_sec: float,
) -> Tuple[Optional[List[str]], str, bool]:
    box: Dict[str, Any] = {}
    done = threading.Event()

    def _runner() -> None:
        try:
            box["value"] = reader()
        except Exception as exc:
            box["error"] = exc
        finally:
            done.set()

    worker = threading.Thread(target=_runner, daemon=True)
    worker.start()
    if not done.wait(max(0.2, float(timeout_sec or 0.0))):
        return None, "", True
    if "error" in box:
        return [], str(box["error"]), False
    value = box.get("value")
    return (value if isinstance(value, list) else []), "", False


def _ax_window_roots_for_onenote(window: Optional[MacWindow]) -> List[c_void_p]:
    if not (IS_MACOS and _APP_SERVICES and window is not None):
        return []
    pid = window.process_id()
    if not pid:
        return []

    app_ref = c_void_p(_APP_SERVICES.AXUIElementCreateApplication(int(pid)))
    if not app_ref:
        return []

    roots: List[c_void_p] = []
    try:
        for attr_name in ("AXFocusedWindow", "AXMainWindow"):
            ref = _ax_element_attribute(app_ref, attr_name)
            if ref:
                roots.append(ref)
        roots.extend(_ax_array_attribute(app_ref, "AXWindows"))
    finally:
        _cf_release(app_ref)

    if not roots:
        return []

    window_number_hint = int(window.info.get("window_number") or 0)
    title_hint = _clean_field(str(window.info.get("title") or ""))
    if not title_hint:
        try:
            title_hint = _clean_field(window.window_text())
        except Exception:
            title_hint = ""
    title_key = _normalize_text(title_hint)

    unique: List[c_void_p] = []
    seen_ptrs = set()
    for root in roots:
        ptr = int(root.value or 0)
        if not ptr:
            continue
        if ptr in seen_ptrs:
            _cf_release(root)
            continue
        seen_ptrs.add(ptr)
        unique.append(root)

    def _rank(root: c_void_p) -> Tuple[int, int, str]:
        ax_window_number = _ax_number_attribute(root, "AXWindowNumber")
        if window_number_hint and ax_window_number == window_number_hint:
            ax_title = _ax_text_attribute(root, "AXTitle")
            return (0, 0, _normalize_text(ax_title))
        ax_title = _ax_text_attribute(root, "AXTitle")
        key = _normalize_text(ax_title)
        if title_key and key == title_key:
            return (0, 1, key)
        if title_key and (title_key in key or key in title_key):
            return (0, 2, key)
        return (1, 0, key)

    return sorted(unique, key=_rank)


def _notebook_name_from_ax_label(raw_label: str) -> str:
    label = _clean_field(raw_label)
    if not label:
        return ""
    name = _extract_current_notebook_name(label)
    name = _clean_field(name)
    key = _normalize_text(name)
    if not key:
        return ""
    blocked_exact = {
        "전자 필기장",
        "전자필기장",
        "notebook",
        "notebooks",
        "open notebook",
        "close notebook",
        "열기",
        "닫기",
        "검색",
        "추가",
    }
    if key in blocked_exact:
        return ""
    blocked_fragments = (
        "sectiontab",
        "pagetab",
        "outline",
        "ax",
        "button",
        "scroll",
    )
    if any(fragment in key for fragment in blocked_fragments):
        return ""
    if len(name) > 120:
        return ""
    return name


def _ax_candidate_labels(element: c_void_p, depth: int = 0) -> List[str]:
    if not element or depth > 4:
        return []

    labels: List[str] = []
    for attr_name in ("AXDescription", "AXTitle", "AXValue"):
        label = _ax_text_attribute(element, attr_name)
        if label:
            labels.append(label)

    children = _ax_array_attribute(element, "AXChildren")
    try:
        preferred: List[str] = []
        fallback: List[str] = []
        for child in children:
            role = _ax_text_attribute(child, "AXRole")
            child_labels = _ax_candidate_labels(child, depth + 1)
            if role in {"AXStaticText", "AXTextField", "AXCell"}:
                preferred.extend(child_labels)
            else:
                fallback.extend(child_labels)
        labels.extend(preferred)
        labels.extend(fallback)
    finally:
        _release_ax_refs(children)

    deduped: List[str] = []
    seen = set()
    for label in labels:
        key = _normalize_text(label)
        if not key or key in seen:
            continue
        seen.add(key)
        deduped.append(label)
    return deduped


def _ax_row_notebook_name(row: c_void_p) -> str:
    for label in _ax_candidate_labels(row):
        name = _notebook_name_from_ax_label(label)
        if name:
            return name
    return ""


def _ax_outline_notebook_names(outline: c_void_p) -> List[str]:
    rows = _ax_array_attribute(outline, "AXRows")
    if not rows:
        rows = _ax_array_attribute(outline, "AXChildren")

    names: List[str] = []
    seen = set()
    try:
        for row in rows:
            name = _ax_row_notebook_name(row)
            key = _normalize_text(name)
            if not key or key in seen:
                continue
            seen.add(key)
            names.append(name)
    finally:
        _release_ax_refs(rows)
    return names


def _collect_ax_outline_name_groups(root: c_void_p) -> List[Dict[str, Any]]:
    groups: List[Dict[str, Any]] = []
    node_count = 0
    deadline = time.monotonic() + 4.0

    def _visit(element: c_void_p, depth: int) -> None:
        nonlocal node_count
        if not element or depth > 18 or node_count >= 1800:
            return
        if time.monotonic() > deadline:
            return
        if not int(element.value or 0):
            return
        node_count += 1

        role = _ax_text_attribute(element, "AXRole")
        if role == "AXOutline":
            names = _ax_outline_notebook_names(element)
            if names:
                groups.append(
                    {
                        "names": names,
                        "order": len(groups),
                        "depth": depth,
                    }
                )

        children = _ax_array_attribute(element, "AXChildren")
        try:
            for child in children:
                _visit(child, depth + 1)
        finally:
            _release_ax_refs(children)

    _visit(root, 0)
    return groups


def _read_open_notebook_names_from_ax(window: Optional[MacWindow]) -> List[str]:
    global _MAC_LAST_AX_NOTEBOOK_DEBUG
    debug: Dict[str, Any] = {
        "trusted": macos_accessibility_is_trusted(),
        "pid": 0,
        "title": "",
        "roots": 0,
        "groups": 0,
        "best_count": 0,
        "reason": "",
    }
    if not (IS_MACOS and window is not None):
        debug["reason"] = "not_macos_or_no_window"
        _MAC_LAST_AX_NOTEBOOK_DEBUG = debug
        return []

    current_title = _clean_field(str(window.info.get("title") or ""))
    debug["pid"] = window.process_id()
    if not current_title:
        try:
            current_title = _clean_field(window.window_text())
        except Exception:
            current_title = ""
    debug["title"] = current_title
    current_key = _normalize_text(current_title)

    roots = _ax_window_roots_for_onenote(window)
    debug["roots"] = len(roots)
    if not roots:
        debug["reason"] = "no_ax_roots"
        _MAC_LAST_AX_NOTEBOOK_DEBUG = debug
        return []

    groups: List[Dict[str, Any]] = []
    try:
        for root in roots[:3]:
            groups.extend(_collect_ax_outline_name_groups(root))
    finally:
        _release_ax_refs(roots)

    debug["groups"] = len(groups)
    if not groups:
        debug["reason"] = "no_ax_outline_groups"
        _MAC_LAST_AX_NOTEBOOK_DEBUG = debug
        return []

    def _score(group: Dict[str, Any]) -> Tuple[int, int, int]:
        names = [str(name or "") for name in group.get("names") or []]
        keys = {_normalize_text(name) for name in names}
        score = min(len(names), 80)
        if current_key and current_key in keys:
            score += 300
        if len(names) >= 3:
            score += 20
        score -= int(group.get("order") or 0) * 3
        score -= int(group.get("depth") or 0)
        return (score, len(names), -int(group.get("order") or 0))

    best = max(groups, key=_score)
    best_names = [str(name or "").strip() for name in best.get("names") or []]
    debug["best_count"] = len(best_names)
    if current_key and current_key not in {_normalize_text(name) for name in best_names}:
        if len(best_names) >= 10:
            debug["reason"] = "best_group_missing_current_title_fallback_large_group"
            _MAC_LAST_AX_NOTEBOOK_DEBUG = debug
            return best_names
        debug["reason"] = "best_group_missing_current_title"
        _MAC_LAST_AX_NOTEBOOK_DEBUG = debug
        return []
    if len(best_names) < 2 and not current_key:
        debug["reason"] = "best_group_too_small"
        _MAC_LAST_AX_NOTEBOOK_DEBUG = debug
        return []
    debug["reason"] = "ok"
    _MAC_LAST_AX_NOTEBOOK_DEBUG = debug
    return best_names


def _ax_press_notebook_sidebar_button(window: Optional[MacWindow]) -> bool:
    if not (IS_MACOS and _APP_SERVICES and window is not None):
        return False

    roots = _ax_window_roots_for_onenote(window)
    if not roots:
        return False

    def _visit(element: c_void_p, depth: int) -> bool:
        if not element or depth > 14:
            return False
        role = _ax_text_attribute(element, "AXRole")
        if role in {"AXButton", "AXMenuButton"}:
            help_text = _ax_text_attribute(element, "AXHelp")
            title_text = _ax_text_attribute(element, "AXTitle")
            combined = _normalize_text(f"{title_text} {help_text}")
            if (
                "전자 필기장 보기" in help_text
                or "view create or open notebooks" in combined
                or "view notebooks" in combined
            ):
                return _ax_perform_action(element, "AXPress")

        children = _ax_array_attribute(element, "AXChildren")
        try:
            for child in children:
                if _visit(child, depth + 1):
                    return True
        finally:
            _release_ax_refs(children)
        return False

    try:
        for root in roots[:3]:
            if _visit(root, 0):
                time.sleep(0.08)
                return True
    finally:
        _release_ax_refs(roots)
    return False

_publish_context(globals())

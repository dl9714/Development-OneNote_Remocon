# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



# ----------------- 10. 선택된 항목을 중앙으로 스크롤 -----------------
def scroll_selected_item_to_center(
    onenote_window,
    tree_control: Optional[object] = None,
    *,
    selected_item=None,
    expected_text: str = "",
    fast_windows: bool = False,
):
    if IS_MACOS:
        try:
            return mac_center_selected_row(
                onenote_window,
                prefer_leftmost=True,
                target_text=expected_text,
            )
        except Exception as e:
            print(f"[WARN] 중앙 정렬 중 오류(macOS): {e}")
            return False, None
    ensure_pywinauto()
    if not _pwa_ready:
        return False, None
    try:
        tree_control = tree_control or _find_tree_or_list(onenote_window)
        if not tree_control:
            return False, None

        selected_item = selected_item or get_selected_tree_item_fast(tree_control)
        if not selected_item and IS_WINDOWS and not fast_windows:
            current_tree_key = _wrapper_identity_key(tree_control)
            fallback_types = ("Tree", "List") if expected_text else ("Tree",)
            for candidate_tree in _iter_tree_or_list_controls(
                onenote_window,
                control_types=fallback_types,
            ):
                if _wrapper_identity_key(candidate_tree) == current_tree_key:
                    continue
                candidate_item = get_selected_tree_item_fast(candidate_tree)
                if candidate_item:
                    tree_control = candidate_tree
                    selected_item = candidate_item
                    break
        if not selected_item:
            if _debug_hotpaths_enabled():
                print("[DBG][CENTER][TARGET] selected_item=None")
            return False, None

        item_name = selected_item.window_text()
        anchor_element, anchor_source, placement = _resolve_alignment_target_for_selected_item(
            selected_item, tree_control
        )
        if _debug_hotpaths_enabled():
            try:
                has_focus = bool(selected_item.has_keyboard_focus())
            except Exception:
                has_focus = False
            try:
                is_selected = bool(selected_item.is_selected())
            except Exception:
                is_selected = False
            depth = _control_depth_within_tree(selected_item, tree_control)
            rect = _safe_rectangle(selected_item)
            height = None if rect is None else max(1, rect.bottom - rect.top)
            anchor_text = _safe_window_text(anchor_element)
            anchor_rect = _safe_rectangle(anchor_element)
            anchor_height = None if anchor_rect is None else max(1, anchor_rect.bottom - anchor_rect.top)
            print(
                "[DBG][CENTER][TARGET]",
                f"text={item_name!r}",
                f"type={_safe_control_type(selected_item)!r}",
                f"depth={depth}",
                f"height={height}",
                f"placement={placement}",
                f"anchor_source={anchor_source}",
                f"anchor_text={anchor_text!r}",
                f"anchor_height={anchor_height}",
                f"selected={is_selected}",
                f"focus={has_focus}",
            )
        _center_element_in_view(
            selected_item,
            tree_control,
            anchor_element=anchor_element,
            placement=placement,
        )
        return True, item_name
    except (ElementNotFoundError, TimeoutError):
        return False, None
    except Exception:
        return False, None


# ----------------- 11. 연결 시그니처 저장/스코어 기반 재획득 -----------------
def _window_info_dict(win) -> Dict[str, Any]:
    if win is None:
        return {}
    try:
        info = object.__getattribute__(win, "info")
        if isinstance(info, dict):
            return info
    except Exception:
        pass
    if IS_MACOS:
        try:
            info = getattr(win, "info", {}) or {}
            if isinstance(info, dict):
                return info
        except Exception:
            pass
    return {}


def _preferred_connected_window_title(
    win,
    fallback_sig: Optional[Dict[str, Any]] = None,
) -> str:
    info = _window_info_dict(win)

    def _clean(value: Any) -> str:
        return str(value or "").strip()

    def _non_generic(value: Any) -> str:
        text = _clean(value)
        if text and text.casefold() not in MACOS_GENERIC_ONENOTE_TITLES:
            return text
        return ""

    raw_title = ""
    try:
        raw_title = _clean(win.window_text())
    except Exception:
        raw_title = ""

    if IS_MACOS:
        preferred_title = _non_generic(raw_title)
        if preferred_title:
            return preferred_title
        try:
            for name in mac_current_open_notebook_names(win):
                preferred_title = _non_generic(name)
                if preferred_title:
                    return preferred_title
        except Exception:
            pass
        info_title = _non_generic(info.get("title"))
        if info_title:
            return info_title

    if raw_title:
        return raw_title

    if isinstance(fallback_sig, dict):
        for key in ("title", "app_name", "bundle_id", "class_name"):
            text = _clean(fallback_sig.get(key))
            if text:
                return text

    for key in ("title", "app_name", "bundle_id", "class_name"):
        text = _clean(info.get(key))
        if text:
            return text

    return ""


def _merge_connection_signature(
    new_sig: Dict[str, Any],
    previous_sig: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    merged = dict(new_sig or {})
    if not IS_MACOS or not isinstance(previous_sig, dict):
        return merged

    new_title = str(merged.get("title") or "").strip()
    prev_title = str(previous_sig.get("title") or "").strip()
    if (
        prev_title
        and (
            not new_title
            or (
                new_title.casefold() in MACOS_GENERIC_ONENOTE_TITLES
                and prev_title.casefold() not in MACOS_GENERIC_ONENOTE_TITLES
            )
        )
    ):
        merged["title"] = prev_title
    if (
        not int(merged.get("handle") or 0)
        and int(previous_sig.get("handle") or 0)
    ):
        merged["handle"] = int(previous_sig.get("handle") or 0)
    if (
        not int(merged.get("window_number") or 0)
        and int(previous_sig.get("window_number") or 0)
    ):
        merged["window_number"] = int(previous_sig.get("window_number") or 0)
    return merged


def _preferred_connected_window_title_quick(
    win,
    fallback_sig: Optional[Dict[str, Any]] = None,
) -> str:
    info = _window_info_dict(win)

    for source in (info, fallback_sig or {}):
        if not isinstance(source, dict):
            continue
        for key in ("title", "app_name", "bundle_id", "class_name"):
            text = str(source.get(key) or "").strip()
            if text and text.casefold() not in MACOS_GENERIC_ONENOTE_TITLES:
                return text

    for source in (info, fallback_sig or {}):
        if not isinstance(source, dict):
            continue
        for key in ("title", "app_name", "bundle_id", "class_name"):
            text = str(source.get(key) or "").strip()
            if text:
                return text

    return ""


def _is_macos_recent_notebooks_dialog_title(title: Optional[str]) -> bool:
    text = str(title or "").strip().casefold()
    if not text:
        return False
    return any(token.casefold() in text for token in MACOS_RECENT_NOTEBOOK_DIALOG_TITLE_TOKENS)


def _resolve_macos_primary_notebook_window(
    win: Optional[object],
    fallback_sig: Optional[Dict[str, Any]] = None,
) -> Optional[MacWindow]:
    if not IS_MACOS or win is None:
        return win if isinstance(win, MacWindow) else None

    current_info = dict(_window_info_dict(win))
    current_title = _preferred_connected_window_title_quick(win, fallback_sig)
    if current_title and not _is_macos_recent_notebooks_dialog_title(current_title):
        return win if isinstance(win, MacWindow) else MacWindow(current_info or dict(fallback_sig or {}))

    current_pid = int(current_info.get("pid") or (fallback_sig or {}).get("pid") or 0)
    candidates = [
        info
        for info in enumerate_macos_windows_quick(filter_title_substr=None)
        if is_macos_onenote_window_info(info, os.getpid())
    ]
    if not candidates:
        return win if isinstance(win, MacWindow) else None

    def _score(info: Dict[str, Any]) -> int:
        title = str(info.get("title") or "").strip()
        title_cf = title.casefold()
        score = 0
        if title and title_cf not in MACOS_GENERIC_ONENOTE_TITLES:
            score += 20
        if not _is_macos_recent_notebooks_dialog_title(title):
            score += 80
        if current_pid and int(info.get("pid") or 0) == current_pid:
            score += 20
        if bool(info.get("frontmost")):
            score += 10
        return score

    best = max(candidates, key=_score)
    if _score(best) <= 0:
        return win if isinstance(win, MacWindow) else None
    try:
        return MacWindow(dict(best))
    except Exception:
        return win if isinstance(win, MacWindow) else None


def build_window_signature_quick(
    win,
    fallback_sig: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    info = _window_info_dict(win)

    if IS_WINDOWS:
        try:
            pid = int(win.process_id() or 0) or None
        except Exception:
            pid = int(info.get("pid") or 0) or None
        try:
            cls_name = str(win.class_name() or "").strip()
        except Exception:
            cls_name = str(info.get("class_name") or "").strip()
        try:
            handle = int(win.handle or 0) or None
        except Exception:
            handle = int(info.get("handle") or 0) or None
        bundle_id = ""
        window_number = None
    else:
        try:
            pid = int(info.get("pid") or win.process_id() or 0) or None
        except Exception:
            pid = None
        try:
            bundle_id = str(info.get("bundle_id") or win.bundle_id() or "").strip()
        except Exception:
            bundle_id = str(info.get("bundle_id") or "").strip()
        try:
            cls_name = str(
                info.get("class_name") or win.class_name() or info.get("app_name") or ""
            ).strip()
        except Exception:
            cls_name = str(info.get("class_name") or info.get("app_name") or "").strip()
        try:
            handle = int(info.get("handle") or win.handle or 0) or None
        except Exception:
            handle = int(info.get("handle") or 0) or None
        try:
            window_number = int(info.get("window_number") or 0) or None
        except Exception:
            window_number = int(info.get("window_number") or 0) or None

    exe_name = os.path.basename(bundle_id or cls_name or "").lower()
    title = _preferred_connected_window_title_quick(win, fallback_sig)
    return _merge_connection_signature(
        {
            "handle": handle,
            "window_number": window_number,
            "pid": pid,
            "class_name": cls_name,
            "title": title,
            "exe_path": "",
            "exe_name": exe_name,
            "bundle_id": bundle_id,
        },
        fallback_sig if isinstance(fallback_sig, dict) else None,
    )


def build_window_signature(win) -> dict:
    try:
        pid = win.process_id()
    except Exception:
        pid = None
    if IS_MACOS:
        bundle_id = getattr(win, "bundle_id", lambda: "")() or ""
        exe_path = ""
        exe_name = os.path.basename(bundle_id or win.class_name() or "").lower()
    else:
        bundle_id = ""
        exe_path = get_process_image_path(pid) if pid else None
        exe_name = os.path.basename(exe_path).lower() if exe_path else None
    try:
        handle = win.handle
    except Exception:
        handle = None
    try:
        window_number = int(_window_info_dict(win).get("window_number") or 0) or None
    except Exception:
        window_number = None
    title = _preferred_connected_window_title(win)
    try:
        cls_name = win.class_name()
    except Exception:
        cls_name = None

    return {
        "handle": handle,
        "window_number": window_number,
        "pid": pid,
        "class_name": cls_name,
        "title": title,
        "exe_path": exe_path,
        "exe_name": exe_name,
        "bundle_id": bundle_id,
    }


def _build_connection_signature_for_save(
    window_element,
    previous_sig: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    if not IS_WINDOWS:
        return build_window_signature(window_element)
    info = build_window_signature_quick(window_element, previous_sig)
    try:
        current_title = str(window_element.window_text() or "").strip()
        if current_title:
            info["title"] = current_title
    except Exception:
        pass
    if isinstance(previous_sig, dict):
        if previous_sig.get("exe_path"):
            info["exe_path"] = previous_sig.get("exe_path")
        previous_exe_name = str(previous_sig.get("exe_name") or "").strip()
        current_exe_name = str(info.get("exe_name") or "").strip()
        class_name = str(info.get("class_name") or "").strip()
        if previous_exe_name and (
            not current_exe_name
            or current_exe_name.casefold() == class_name.casefold()
        ):
            info["exe_name"] = previous_exe_name
    return info


def save_connection_info(window_element):
    try:
        current_settings = load_settings()
        current_sig = current_settings.get("connection_signature")
        info = _build_connection_signature_for_save(
            window_element,
            current_sig if isinstance(current_sig, dict) else None,
        )
        info = _merge_connection_signature(info, current_sig)
        if current_settings.get("connection_signature") == info:
            return
        current_settings["connection_signature"] = info
        save_settings(current_settings)
    except Exception as e:
        print(f"[ERROR] 연결 정보 저장 실패: {e}")

_publish_context(globals())

# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def get_selected_tree_item_fast(tree_control):
    ensure_pywinauto()
    if not _pwa_ready:
        return None
    candidates = []
    seen = set()

    def _push(item):
        if item is None:
            return
        key = _wrapper_identity_key(item)
        if key in seen:
            return
        seen.add(key)
        candidates.append(item)

    def _best_candidate():
        if not candidates:
            return None
        return _pick_best_tree_item_candidate(tree_control, candidates)

    try:
        if hasattr(tree_control, "get_selection"):
            sel = tree_control.get_selection()
            if sel:
                for item in sel:
                    _push(item)
    except Exception:
        pass

    best = _best_candidate()
    if best is not None:
        return best

    try:
        iface_sel = getattr(tree_control, "iface_selection", None)
        if iface_sel:
            arr = iface_sel.GetSelection()
            length = getattr(arr, "Length", 0)
            if length and length > 0:
                for idx in range(length):
                    try:
                        el = arr.GetElement(idx)
                        _push(UIAWrapper(UIAElementInfo(el)))
                    except Exception:
                        continue
    except Exception:
        pass

    best = _best_candidate()
    if best is not None:
        return best

    try:
        for item in tree_control.children():
            try:
                if item.is_selected():
                    _push(item)
            except Exception:
                pass
    except Exception:
        pass

    best = _best_candidate()
    if best is not None:
        return best

    try:
        for control_type in ("TreeItem", "ListItem"):
            for item in tree_control.descendants(control_type=control_type):
                try:
                    if item.is_selected() or item.has_keyboard_focus():
                        _push(item)
                except Exception:
                    pass
    except Exception:
        pass

    return _best_candidate()


# ----------------- 8. 페이지/섹션 컨테이너(Tree/List) 찾기 - ensure 호출 -----------------
def _find_tree_or_list(onenote_window):
    ensure_pywinauto()
    if not _pwa_ready:
        return None
    for ctype in ("Tree", "List"):
        try:
            return onenote_window.child_window(
                control_type=ctype, found_index=0
            ).wrapper_object()
        except Exception:
            continue
    return None


# ----------------- 8.1 지정 텍스트 섹션 찾기/선택 -----------------
def _normalize_text(s: Optional[str]) -> str:
    return " ".join(((s or "").strip().split())).lower()


_NOTEBOOK_NAME_KEY_CACHE = {}
_NOTEBOOK_NAME_NORMALIZE = None
_NOTEBOOK_NAME_SEPARATOR_SUB = None


def _normalize_notebook_name_key(s: Optional[str]) -> str:
    global _NOTEBOOK_NAME_NORMALIZE, _NOTEBOOK_NAME_SEPARATOR_SUB
    raw = s if isinstance(s, str) else str(s or "")
    cached = _NOTEBOOK_NAME_KEY_CACHE.get(raw)
    if cached is not None:
        return cached
    if _NOTEBOOK_NAME_NORMALIZE is None:
        _NOTEBOOK_NAME_NORMALIZE = unicodedata.normalize
    if _NOTEBOOK_NAME_SEPARATOR_SUB is None:
        _NOTEBOOK_NAME_SEPARATOR_SUB = re.compile(r"[\s\-_]+").sub
    text = _NOTEBOOK_NAME_NORMALIZE("NFKC", raw).casefold()
    key = _NOTEBOOK_NAME_SEPARATOR_SUB("", text)
    if len(_NOTEBOOK_NAME_KEY_CACHE) >= 4096:
        _NOTEBOOK_NAME_KEY_CACHE.clear()
    _NOTEBOOK_NAME_KEY_CACHE[raw] = key
    return key


def _normalize_project_search_key(s: Optional[str]) -> str:
    text = unicodedata.normalize("NFKC", s or "").casefold()
    return re.sub(r"\s+", "", text)


def _strip_stale_favorite_prefix(text: Optional[str]) -> str:
    raw = (text or "").strip()
    for prefix in ("(구) ", "(old) "):
        if raw.startswith(prefix):
            return raw[len(prefix) :].strip()
    return raw


def select_section_by_text(
    onenote_window, text: str, tree_control: Optional[object] = None
) -> bool:
    ensure_pywinauto()
    if not _pwa_ready:
        return False
    if IS_MACOS:
        try:
            return mac_select_row_by_text(onenote_window, text)
        except Exception as e:
            print(f"[ERROR] 섹션 선택 실패(macOS): {e}")
            return False
    try:
        tree_control = tree_control or _find_tree_or_list(onenote_window)
        if not tree_control:
            return False

        target_norm = _normalize_text(text)

        def _scan(types: List[str]):
            for t in types:
                try:
                    for itm in tree_control.descendants(control_type=t):
                        try:
                            if _normalize_text(itm.window_text()) == target_norm:
                                try:
                                    itm.select()
                                    return True
                                except Exception:
                                    try:
                                        itm.click_input()
                                        return True
                                    except Exception:
                                        pass
                        except Exception:
                            pass
                except Exception:
                    pass
            return False

        if _scan(["TreeItem"]):
            return True
        if _scan(["ListItem"]):
            return True
        return False
    except Exception as e:
        print(f"[ERROR] 섹션 선택 실패: {e}")
        return False


def select_notebook_item_by_text(
    onenote_window,
    text: str,
    tree_control: Optional[object] = None,
    *,
    center_after_select: bool = False,
):
    """
    전자필기장(노트북) 이름으로 찾고 선택합니다.
    - root children 우선 탐색(전자필기장은 보통 루트에 있음)
    - 실패하면 descendants(TreeItem/ListItem)로 fallback
    """
    ensure_pywinauto()
    if not _pwa_ready:
        return False
    if IS_MACOS:
        try:
            if not mac_select_row_by_text(onenote_window, text):
                return None
            if center_after_select:
                mac_center_selected_row(onenote_window, prefer_leftmost=True)
            return mac_pick_selected_row(onenote_window, prefer_leftmost=True)
        except Exception as e:
            print(f"[ERROR] 전자필기장 선택 실패(macOS): {e}")
            return None
    try:
        tree_control = tree_control or _find_tree_or_list(onenote_window)
        if not tree_control:
            return False
        target_norm = _normalize_text(text)

        def _select_and_center(item):
            try:
                item.select()
            except Exception:
                try:
                    item.click_input()
                except Exception:
                    return None

            if center_after_select:
                _center_element_in_view(
                    item,
                    tree_control,
                    placement=_get_scroll_placement_for_selected_item(item, tree_control),
                )
            return item

        # 1) root-level children 우선
        try:
            for item in (tree_control.children() or []):
                try:
                    if _normalize_text(item.window_text()) == target_norm:
                        return _select_and_center(item)
                except Exception:
                    continue
        except Exception:
            pass

        # 2) descendants fallback
        for control_type in ["TreeItem", "ListItem"]:
            try:
                for item in tree_control.descendants(control_type=control_type):
                    try:
                        if _normalize_text(item.window_text()) == target_norm:
                            return _select_and_center(item)
                    except Exception:
                        pass
            except Exception:
                pass

        return None
    except Exception as e:
        print(f"[ERROR] 전자필기장 선택 실패: {e}")
        return None


def select_notebook_by_text(
    onenote_window, text: str, tree_control: Optional[object] = None, *, center_after_select: bool = False
) -> bool:
    return (
        select_notebook_item_by_text(
            onenote_window,
            text,
            tree_control,
            center_after_select=center_after_select,
        )
        is not None
    )


def _mac_outline_notebook_matches(onenote_window, notebook_name: str) -> bool:
    if not IS_MACOS:
        return False
    expected_key = _normalize_notebook_name_key(notebook_name)
    if not expected_key:
        return False
    current_name = mac_current_notebook_name(onenote_window)
    if current_name:
        return _normalize_notebook_name_key(current_name) == expected_key
    for current_name in (
        _get_macos_primary_notebook_title(),
        str(getattr(onenote_window, "info", {}).get("title") or "").strip(),
    ):
        if _normalize_notebook_name_key(current_name) == expected_key:
            return True
    return False


def _mac_wait_for_notebook_context(
    onenote_window,
    notebook_name: str,
    *,
    timeout_sec: float = 4.0,
) -> bool:
    if not IS_MACOS:
        return True
    expected_key = _normalize_notebook_name_key(notebook_name)
    if not expected_key:
        return True
    deadline = time.monotonic() + max(0.1, float(timeout_sec or 0.0))
    while time.monotonic() < deadline:
        if _mac_outline_notebook_matches(onenote_window, notebook_name):
            return True
        time.sleep(0.2)
    return _mac_outline_notebook_matches(onenote_window, notebook_name)


def _mac_find_recent_notebook_record(
    onenote_window,
    notebook_name: str,
) -> Optional[Dict[str, Any]]:
    expected_key = _normalize_notebook_name_key(notebook_name)
    if not expected_key:
        return None
    try:
        records = mac_recent_notebook_records(onenote_window)
    except Exception as e:
        print(f"[WARN][MAC][RECENT_MATCH] {e}")
        return None
    best = None
    for record in records or []:
        name = str((record or {}).get("name") or "").strip()
        key = _normalize_notebook_name_key(name)
        if not key:
            continue
        if key == expected_key:
            return dict(record)
        if best is None and (expected_key in key or key in expected_key):
            best = dict(record)
    return best

_publish_context(globals())

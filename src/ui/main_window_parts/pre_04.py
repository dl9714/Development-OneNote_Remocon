# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    finalize_context as _finalize_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _ensure_default_and_aggregate_inplace(settings: Dict[str, Any]) -> None:
    """
    1패널(버퍼 트리)에
      - Default 그룹(최상단 고정)
      - 종합(가상 버퍼)
    를 항상 보장합니다.
    """
    bufs = settings.get("favorites_buffers")
    if not isinstance(bufs, list):
        bufs = []
        settings["favorites_buffers"] = bufs

    # Default 그룹 찾기/생성
    default_idx = None
    default_node = None
    for i, n in enumerate(bufs):
        if isinstance(n, dict) and n.get("type") == "group" and n.get("id") == DEFAULT_GROUP_ID:
            default_idx = i
            default_node = n
            break
    if default_node is None:
        default_node = {
            "type": "group",
            "id": DEFAULT_GROUP_ID,
            "name": DEFAULT_GROUP_NAME,
            "locked": True,
            "children": []
        }
        bufs.insert(0, default_node)
    else:
        # 항상 0번으로 이동 + 이름 강제
        default_node["name"] = DEFAULT_GROUP_NAME
        default_node["locked"] = True
        if default_idx != 0:
            bufs.pop(default_idx)
            bufs.insert(0, default_node)

    # Default 그룹 children 보정
    children = default_node.get("children")
    if not isinstance(children, list):
        children = []
        default_node["children"] = children

    # 종합(가상 버퍼) 찾기/생성 (Default 그룹의 첫 번째로 고정)
    agg_idx = None
    agg_node = None
    for i, c in enumerate(children):
        if isinstance(c, dict) and c.get("type") == "buffer" and c.get("id") == AGG_BUFFER_ID:
            agg_idx = i
            agg_node = c
            break
    if agg_node is None:
        agg_node = {
            "type": "buffer",
            "id": AGG_BUFFER_ID,
            "name": AGG_BUFFER_NAME,
            "virtual": "aggregate",
            "locked": True,
            "data": []
        }
        children.insert(0, agg_node)
    else:
        # 항상 children[0]으로 이동 + 이름/속성 강제
        agg_node["name"] = AGG_BUFFER_NAME
        agg_node["locked"] = True
        agg_node["virtual"] = "aggregate"
        if agg_idx != 0:
            children.pop(agg_idx)
            children.insert(0, agg_node)

    # active_buffer_id가 종합/기타 버퍼에 대해 유효하지 않으면 첫 "일반 buffer"로 보정
    # (종합으로 시작하고 싶으면 이 부분을 제거하면 됨)
    if settings.get("active_buffer_id") is None:
        # 기본은 종합이 아니라, 첫 일반 buffer를 찾는다
        first_normal = _find_first_normal_buffer_id(bufs)
        if first_normal:
            settings["active_buffer_id"] = first_normal


def _find_first_normal_buffer_id(nodes: List[Dict[str, Any]]) -> Optional[str]:
    def _walk(lst):
        for n in lst:
            if not isinstance(n, dict):
                continue
            if n.get("type") == "buffer":
                if n.get("id") != AGG_BUFFER_ID:
                    return n.get("id")
            if n.get("type") == "group":
                cid = _walk(n.get("children") or [])
                if cid:
                    return cid
        return None
    return _walk(nodes)


def _find_buffer_node_by_id(nodes: Any, buffer_id: Optional[str]) -> Optional[Dict[str, Any]]:
    if not buffer_id or not isinstance(nodes, list):
        return None

    stack = list(reversed(nodes))
    while stack:
        node = stack.pop()
        if not isinstance(node, dict):
            continue
        node_type = node.get("type")
        if node_type == "buffer" and node.get("id") == buffer_id:
            return node
        if node_type == "group":
            children = node.get("children")
            if isinstance(children, list) and children:
                stack.extend(reversed(children))
    return None


def _collect_all_sections_dedup(settings: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    모든 버퍼의 data에서 section/notebook들을 수집해서 (sig + text) 기준 중복 제거 후 반환.
    - 반환 형태: 중앙트리(_load_favorites_into_center_tree)에 바로 넣을 수 있는 node 리스트
    """
    bufs = settings.get("favorites_buffers", [])
    out: List[Dict[str, Any]] = []
    seen = set()

    def _freeze_key(value: Any):
        if isinstance(value, dict):
            try:
                items = sorted(value.items(), key=lambda pair: str(pair[0]))
            except Exception:
                items = value.items()
            return tuple((str(k), _freeze_key(v)) for k, v in items)
        if isinstance(value, list):
            return tuple(_freeze_key(v) for v in value)
        if isinstance(value, (str, int, float, bool, type(None))):
            return value
        return str(value)

    def _section_key(node: Dict[str, Any]):
        t = node.get("target") or {}
        sig = t.get("sig") or {}
        ty = node.get("type", "section")
        # 섹션명 또는 노트북명 수집
        text = t.get("section_text") or t.get("notebook_text") or ""
        return (ty, _freeze_key(sig), text)

    def _walk_fav_nodes(nodes: Any):
        if not isinstance(nodes, list):
            return
        for n in nodes:
            if not isinstance(n, dict):
                continue
            ty = n.get("type")
            if ty in ("section", "notebook"):
                k = _section_key(n)
                if k not in seen:
                    seen.add(k)
                    # 종합은 납작하게(flat) 보여주기
                    out.append({
                        "type": ty,
                        "id": n.get("id") or str(uuid.uuid4()),
                        "name": n.get("name") or (n.get("target") or {}).get("section_text") or (n.get("target") or {}).get("notebook_text") or "항목",
                        "target": n.get("target") or {}
                    })
            # 그룹 아래 children 순회
            ch = n.get("children")
            if isinstance(ch, list):
                _walk_fav_nodes(ch)

    def _walk_buffers(nodes: Any):
        if not isinstance(nodes, list):
            return
        for b in nodes:
            if not isinstance(b, dict):
                continue
            if b.get("type") == "buffer":
                # ✅ 더 이상 aggregate를 스킵하지 않음 (그 안에 직접 등록한 노트북 등 보존 목적)
                _walk_fav_nodes(b.get("data") or [])
            elif b.get("type") == "group":
                _walk_buffers(b.get("children") or [])

    _walk_buffers(bufs)
    # ✅ 종합은 중앙트리에서 '이름순'으로 보이도록 정렬
    try:
        out.sort(key=lambda n: _name_sort_key((n or {}).get("name", "")))
    except Exception:
        pass
    return out


def _collect_known_notebook_name_records(settings: Dict[str, Any]) -> List[Dict[str, Any]]:
    def _coerce_last_accessed_at(value: Any) -> int:
        try:
            return max(0, int(value or 0))
        except Exception:
            return 0

    def _merge_record(existing: Dict[str, Any], incoming: Dict[str, Any]) -> None:
        incoming_id = str(incoming.get("id") or incoming.get("notebook_id") or "").strip()
        incoming_path = str(incoming.get("path") or "").strip()
        incoming_url = str(
            incoming.get("url") or incoming.get("notebook_url") or ""
        ).strip()
        incoming_source = str(incoming.get("source") or "").strip()
        incoming_last_accessed_at = _coerce_last_accessed_at(
            incoming.get("last_accessed_at")
        )

        if not str(existing.get("id") or "").strip() and incoming_id:
            existing["id"] = incoming_id
        if not str(existing.get("path") or "").strip() and incoming_path:
            existing["path"] = incoming_path
        if not str(existing.get("url") or "").strip() and incoming_url:
            existing["url"] = incoming_url
        if not str(existing.get("source") or "").strip() and incoming_source:
            existing["source"] = incoming_source
        if incoming_last_accessed_at > _coerce_last_accessed_at(
            existing.get("last_accessed_at")
        ):
            existing["last_accessed_at"] = incoming_last_accessed_at

    records: Dict[str, Dict[str, Any]] = {}
    for node in _collect_all_sections_dedup(settings):
        if not isinstance(node, dict):
            continue
        target = node.get("target") or {}
        notebook_name = (
            str(target.get("notebook_text") or "").strip()
            or (
                str(node.get("name") or "").strip()
                if str(node.get("type") or "").strip() == "notebook"
                else ""
            )
        )
        key = _normalize_notebook_name_key(notebook_name)
        if not key:
            continue
        candidate = {
            "name": notebook_name,
            "id": str(target.get("notebook_id") or "").strip(),
            "path": str(target.get("path") or "").strip(),
            "url": str(target.get("url") or target.get("notebook_url") or "").strip(),
            "last_accessed_at": _coerce_last_accessed_at(
                target.get("last_accessed_at")
            ),
            "source": "SETTINGS_BUFFER",
        }
        existing = records.get(key)
        if existing is None:
            records[key] = candidate
            continue
        _merge_record(existing, candidate)
    return list(records.values())


# ----------------- 0.1 pywinauto 지연 로딩 -----------------
Desktop = None
WindowNotFoundError = None
ElementNotFoundError = None
TimeoutError = None
UIAWrapper = None
UIAElementInfo = None
mouse = None
keyboard = None

_pwa_ready = False
_pwa_import_error = ""
_pwa_context_synced = False


def _sync_pywinauto_context_for_windows() -> None:
    global _pwa_context_synced
    if not IS_WINDOWS:
        return
    if _pwa_context_synced:
        return
    try:
        _publish_context(globals())
        _finalize_context()
        import sys as _sys

        facade = _sys.modules.get("src.ui.main_window")
        if facade is not None:
            for name in (
                "Desktop",
                "WindowNotFoundError",
                "ElementNotFoundError",
                "TimeoutError",
                "UIAWrapper",
                "UIAElementInfo",
                "mouse",
                "keyboard",
                "_pwa_ready",
                "_pwa_import_error",
            ):
                setattr(facade, name, globals().get(name))
        _pwa_context_synced = True
    except Exception:
        pass


def ensure_pywinauto():
    global _pwa_ready, _pwa_import_error, _pwa_context_synced, Desktop, WindowNotFoundError, ElementNotFoundError, TimeoutError, UIAWrapper, UIAElementInfo, mouse, keyboard
    # NameError 수정: _ppa_ready -> _pwa_ready
    if _pwa_ready:
        _sync_pywinauto_context_for_windows()
        return
    if IS_MACOS:
        from src import macos_ui as _macos_ui

        Desktop = _macos_ui.MacDesktop
        WindowNotFoundError = _macos_ui.MacAutomationError
        ElementNotFoundError = _macos_ui.MacAutomationError
        TimeoutError = TimeoutError or RuntimeError
        UIAWrapper = None
        UIAElementInfo = None
        mouse = None
        keyboard = None
        _pwa_ready = True
        _pwa_import_error = ""
        _sync_pywinauto_context_for_windows()
        return
    try:
        _pwa_context_synced = False
        from pywinauto import (
            Desktop as _Desktop,
            mouse as _mouse,
            keyboard as _keyboard,
        )
        from pywinauto.findwindows import (
            WindowNotFoundError as _WNF,
            ElementNotFoundError as _ENF,
        )
        from pywinauto.timings import TimeoutError as _TO
        from pywinauto.controls.uiawrapper import UIAWrapper as _UIAWrapper
        from pywinauto.uia_element_info import UIAElementInfo as _UIAElementInfo

        Desktop = _Desktop
        WindowNotFoundError = _WNF
        ElementNotFoundError = _ENF
        TimeoutError = _TO
        UIAWrapper = _UIAWrapper
        UIAElementInfo = _UIAElementInfo
        mouse = _mouse
        keyboard = _keyboard
        _pwa_ready = True
        _pwa_import_error = ""
        _sync_pywinauto_context_for_windows()
    except ImportError as e:
        _pwa_import_error = str(e)
        _sync_pywinauto_context_for_windows()
        print(f"[WARN][PWA] import failed: {_pwa_import_error}")


# ----------------- 0.2 Win32 빠른 창 열거 -----------------
_user32 = ctypes.windll.user32 if IS_WINDOWS else None


def _win_get_window_text(hwnd):
    if _user32 is None:
        return ""
    length = _user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(length + 1 if length > 0 else 1)
    _user32.GetWindowTextW(hwnd, buf, len(buf))
    return buf.value


def _win_get_class_name(hwnd):
    if _user32 is None:
        return ""
    buf = ctypes.create_unicode_buffer(256)
    _user32.GetClassNameW(hwnd, buf, 256)
    return buf.value

_publish_context(globals())

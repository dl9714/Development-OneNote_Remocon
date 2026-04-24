# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _ensure_onenote_open_notebook_view(onenote_window) -> bool:
    ensure_pywinauto()
    if not _pwa_ready:
        return False

    try:
        onenote_window.set_focus()
    except Exception:
        pass

    if _is_onenote_open_notebook_view(onenote_window):
        return True

    try:
        keyboard.send_keys("^o")
    except Exception:
        pass
    if _wait_until(lambda: _is_onenote_open_notebook_view(onenote_window), timeout=2.5, interval=0.1):
        return True

    file_item = _find_descendant_by_text(
        onenote_window,
        ONENOTE_FILE_MENU_TEXTS,
        control_types=("Button", "MenuItem", "TabItem", "Hyperlink", "Text"),
    )
    if file_item and _activate_ui_element(file_item):
        time.sleep(0.25)
        open_item = _find_descendant_by_text(
            onenote_window,
            ONENOTE_OPEN_MENU_TEXTS,
            control_types=("Button", "MenuItem", "ListItem", "Hyperlink", "Text"),
        )
        if open_item:
            _activate_ui_element(open_item)
            if _wait_until(lambda: _is_onenote_open_notebook_view(onenote_window), timeout=2.0, interval=0.1):
                return True

    try:
        keyboard.send_keys("^o")
    except Exception:
        pass
    return bool(
        _wait_until(
            lambda: _is_onenote_open_notebook_view(onenote_window),
            timeout=1.5,
            interval=0.1,
        )
    )


def _collect_open_notebook_candidates(onenote_window):
    ensure_pywinauto()
    if not _pwa_ready:
        return [], None

    win_rect = _safe_rectangle(onenote_window)
    if not win_rect:
        return [], None

    left_min = win_rect.left + 60
    left_max = win_rect.left + int((win_rect.right - win_rect.left) * 0.85)
    top_min = win_rect.top + 80
    bottom_max = min(
        win_rect.bottom - 10,
        win_rect.top + int((win_rect.bottom - win_rect.top) * 0.72),
    )

    raw_candidates = []
    seen = set()
    for item in _iter_descendants_by_types(onenote_window, ONENOTE_NOTEBOOK_ITEM_CONTROL_TYPES):
        raw_text = _safe_window_text(item)
        primary_text = _extract_primary_accessible_text(raw_text)
        norm_text = _normalize_text(primary_text)
        if not norm_text:
            continue
        if norm_text in ONENOTE_NOTEBOOK_SKIP_EXACT_TEXTS:
            continue
        if any(skip in norm_text for skip in ONENOTE_NOTEBOOK_SKIP_CONTAINS):
            continue

        rect = _safe_rectangle(item)
        if not rect:
            continue
        width = rect.right - rect.left
        height = rect.bottom - rect.top
        if width < 120 or height < 16 or height > 120:
            continue
        if rect.left < left_min or rect.left > left_max:
            continue
        if rect.top < top_min or rect.bottom > bottom_max:
            continue

        dedupe_key = (norm_text, rect.left // 6, rect.top // 6)
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)

        parent = _safe_parent(item)
        container = _find_scrollable_ancestor(item) or parent
        container_rect = _safe_rectangle(container) if container else None
        container_type = _safe_control_type(container) if container else ""
        raw_candidates.append(
            {
                "item": item,
                "text": primary_text,
                "norm_text": norm_text,
                "rect": rect,
                "parent": parent,
                "container": container,
                "container_rect": container_rect,
                "container_type": container_type,
            }
        )

    if not raw_candidates:
        return [], None

    grouped = {}
    for cand in raw_candidates:
        container_rect = cand.get("container_rect")
        if container_rect is None:
            key = ("", 0, 0, 0, 0)
        else:
            key = (
                cand.get("container_type") or "",
                container_rect.left // 10,
                container_rect.top // 10,
                container_rect.right // 10,
                container_rect.bottom // 10,
            )
        grouped.setdefault(key, []).append(cand)

    best_group = None
    best_score = None
    for group in grouped.values():
        uniq_count = len({cand["norm_text"] for cand in group})
        list_like = sum(
            1
            for cand in group
            if cand.get("container_type") in ("List", "Tree", "Pane", "Group", "Custom")
        )
        score = (uniq_count, list_like, -min(cand["rect"].top for cand in group))
        if best_score is None or score > best_score:
            best_score = score
            best_group = group

    final_candidates = best_group if best_group else raw_candidates
    final_candidates.sort(key=lambda cand: (cand["rect"].top, cand["rect"].left, cand["text"]))
    container = final_candidates[0].get("container") if final_candidates else None
    return final_candidates, container


def _scroll_open_notebook_candidates(container) -> bool:
    if container is None:
        return False

    try:
        container.set_focus()
    except Exception:
        pass

    if _scroll_vertical_via_pattern(container, "down", small=True, repeats=4):
        time.sleep(0.2)
        return True

    try:
        _safe_wheel(container, -4)
        time.sleep(0.2)
        return True
    except Exception:
        pass

    try:
        keyboard.send_keys("{PGDN}")
        time.sleep(0.2)
        return True
    except Exception:
        return False


def _find_next_unattempted_open_notebook(onenote_window, attempted_norms: Set[str]):
    tried_snapshots = set()
    last_candidates = []
    container = None

    for _ in range(24):
        candidates, container = _collect_open_notebook_candidates(onenote_window)
        last_candidates = candidates
        if not candidates:
            return None, [], container

        for cand in candidates:
            if cand["norm_text"] not in attempted_norms:
                return cand, candidates, container

        snapshot = tuple(cand["norm_text"] for cand in candidates[:20])
        if not snapshot or snapshot in tried_snapshots:
            break
        tried_snapshots.add(snapshot)

        if not _scroll_open_notebook_candidates(container):
            break

    return None, last_candidates, container


def _open_one_notebook_from_backstage(
    onenote_window, attempted_norms: Optional[Set[str]] = None
) -> Optional[Dict[str, Any]]:
    if not _ensure_onenote_open_notebook_view(onenote_window):
        return None

    attempted_norms = attempted_norms or set()
    target, candidates, _container = _find_next_unattempted_open_notebook(
        onenote_window, attempted_norms
    )
    if not candidates:
        return {
            "done": False,
            "opened_text": "",
            "visible_names": [],
            "error": "전자 필기장 목록 항목을 찾지 못했습니다.",
        }

    snapshot = [cand["text"] for cand in candidates[:10]]
    if target is None:
        return {"done": True, "opened_text": "", "visible_names": snapshot}

    if not _activate_ui_element(target["item"]):
        return {
            "done": False,
            "opened_text": target["text"],
            "visible_names": snapshot,
            "attempted_norm": target["norm_text"],
            "error": f"전자필기장 열기 실패: '{target['text']}'",
        }

    _wait_until(
        lambda: not _is_onenote_open_notebook_view(onenote_window),
        timeout=2.5,
        interval=0.1,
    )
    time.sleep(0.35)
    return {
        "done": False,
        "opened_text": target["text"],
        "visible_names": snapshot,
        "attempted_norm": target["norm_text"],
    }


def _ps_quote(text: str) -> str:
    return "'" + (text or "").replace("'", "''") + "'"


def _run_powershell(script: str, timeout: int = 30) -> str:
    full_script = (
        "$ErrorActionPreference='Stop';"
        "[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new($false);"
        "$OutputEncoding=[Console]::OutputEncoding;"
        + script
    )
    startupinfo = None
    creationflags = 0
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    completed = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-WindowStyle",
            "Hidden",
            "-Command",
            full_script,
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=max(5, timeout),
        startupinfo=startupinfo,
        creationflags=creationflags,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            (completed.stderr or completed.stdout or "").strip()
            or f"PowerShell exit code {completed.returncode}"
        )
    return (completed.stdout or "").strip()


def _load_json_output(raw_text: str):
    text = (raw_text or "").strip()
    if not text:
        return None
    return json.loads(text)


def _iter_onedrive_notebook_shortcut_dirs() -> List[str]:
    roots = []
    candidates = [
        os.environ.get("OneDrive"),
        os.path.join(os.path.expanduser("~"), "OneDrive"),
    ]
    cloud_storage = os.path.join(os.path.expanduser("~"), "Library", "CloudStorage")
    try:
        for name in os.listdir(cloud_storage):
            if "onedrive" in str(name or "").casefold():
                candidates.append(os.path.join(cloud_storage, name))
    except Exception:
        pass
    for base in candidates:
        if not base or not os.path.isdir(base):
            continue
        for rel in ("문서", "Documents", ""):
            path = os.path.join(base, rel)
            if os.path.isdir(path) and path not in roots:
                roots.append(path)
    return roots


def _read_internet_shortcut_url(path: str) -> str:
    for encoding in ("utf-8-sig", "cp949", "utf-8", "latin-1"):
        try:
            with open(path, "r", encoding=encoding) as f:
                for line in f:
                    if line.startswith("URL="):
                        return line[4:].strip()
        except Exception:
            continue
    return ""


def _read_webloc_url(path: str) -> str:
    try:
        import plistlib

        with open(path, "rb") as f:
            data = plistlib.load(f)
        return str(data.get("URL") or "").strip()
    except Exception:
        return ""


def _looks_like_onenote_shortcut_url(url: str) -> bool:
    lower = (url or "").strip().lower()
    if not lower:
        return False
    return (
        "callerscenarioid=onenote-prod" in lower
        or lower.startswith("onenote:")
        or ("onedrive.live.com/redir.aspx" in lower and "onenote" in lower)
    )

_publish_context(globals())

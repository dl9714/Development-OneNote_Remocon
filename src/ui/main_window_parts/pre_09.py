# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _extract_onedrive_cid(url: str) -> str:
    try:
        parsed = urlparse(url or "")
        query = parse_qs(parsed.query)
        cid = (query.get("cid") or [""])[0].strip()
        return cid.lower()
    except Exception:
        return ""


def _encode_onenote_protocol_segment(text: str) -> str:
    segment = (text or "").strip()
    return (
        segment.replace("%", "%25")
        .replace("#", "%23")
        .replace(" ", "%20")
    )


def _build_onenote_protocol_url(shortcut_path: str, web_url: str, notebook_name: str) -> str:
    cid = _extract_onedrive_cid(web_url)
    if not cid:
        return ""

    root_label = "문서"
    for root in _iter_onedrive_notebook_shortcut_dirs():
        try:
            if os.path.commonpath([os.path.abspath(shortcut_path), os.path.abspath(root)]) == os.path.abspath(root):
                root_label = os.path.basename(root.rstrip("\\/")) or root_label
                break
        except Exception:
            continue

    encoded_root = _encode_onenote_protocol_segment(root_label)
    encoded_name = _encode_onenote_protocol_segment(notebook_name)
    return f"onenote:https://d.docs.live.net/{cid}/{encoded_root}/{encoded_name}/"


def _get_onenote_exe_path() -> str:
    if winreg is None:
        return ""
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"onenote\shell\Open\command") as key:
            command, _ = winreg.QueryValueEx(key, None)
    except Exception:
        return ""

    command = str(command or "").strip()
    if not command:
        return ""

    if command.startswith('"'):
        end = command.find('"', 1)
        if end > 1:
            exe_path = command[1:end]
        else:
            exe_path = ""
    else:
        exe_path = command.split(" ", 1)[0].strip()

    return exe_path if exe_path and os.path.isfile(exe_path) else ""


def _collect_onenote_notebook_shortcuts() -> List[Dict[str, str]]:
    results: Dict[str, Dict[str, str]] = {}
    for root in _iter_onedrive_notebook_shortcut_dirs():
        try:
            names = sorted(os.listdir(root), key=_name_sort_key)
        except Exception:
            continue
        for name in names:
            lower_name = name.lower()
            if not (lower_name.endswith(".url") or lower_name.endswith(".webloc")):
                continue
            path = os.path.join(root, name)
            if not os.path.isfile(path):
                continue
            if lower_name.endswith(".webloc"):
                url = _read_webloc_url(path)
            else:
                url = _read_internet_shortcut_url(path)
            if not _looks_like_onenote_shortcut_url(url):
                continue
            display_name = os.path.splitext(name)[0].strip()
            norm_name = _normalize_notebook_name_key(display_name)
            if not norm_name or norm_name in results:
                continue
            results[norm_name] = {
                "name": display_name,
                "path": path,
                "url": url,
            }
    return list(results.values())


def _get_open_notebook_names_via_com(
    refresh: bool = False,
    max_age_sec: float = _OPEN_NOTEBOOK_RECORDS_CACHE_TTL_SEC,
) -> List[str]:
    records = _get_open_notebook_records_via_com(
        refresh=refresh, max_age_sec=max_age_sec
    )
    return [
        str(record.get("name") or "").strip()
        for record in records
        if str(record.get("name") or "").strip()
    ]


def _get_macos_primary_notebook_title() -> str:
    if not IS_MACOS:
        return ""

    try:
        wins = [
            info
            for info in enumerate_macos_windows(filter_title_substr=None)
            if is_macos_onenote_window_info(info, os.getpid())
        ]
        wins.sort(key=lambda item: (not bool(item.get("frontmost")), item.get("title", "")))
        for info in wins:
            title = str(info.get("title") or "").strip()
            if title:
                return title
    except Exception as e:
        print(f"[WARN][MAC][PRIMARY_NOTEBOOK] {e}")
    return ""


def _is_macos_notebook_visible(expected_name: str) -> bool:
    expected_key = _normalize_notebook_name_key(expected_name)
    if not expected_key or not IS_MACOS:
        return False

    current_title = _get_macos_primary_notebook_title()
    if _normalize_notebook_name_key(current_title) == expected_key:
        return True

    try:
        open_keys = {
            _normalize_notebook_name_key(name)
            for name in _get_open_notebook_names_via_com(refresh=True)
            if _normalize_notebook_name_key(name)
        }
    except Exception:
        open_keys = set()
    return expected_key in open_keys


def _wait_for_macos_notebook_visible(expected_name: str, timeout_sec: float = 6.0) -> bool:
    if not IS_MACOS:
        return False

    deadline = time.monotonic() + max(0.1, float(timeout_sec or 0.0))
    while time.monotonic() < deadline:
        if _is_macos_notebook_visible(expected_name):
            return True
        time.sleep(0.25)
    return _is_macos_notebook_visible(expected_name)


def _normalize_notebook_record(raw: Any) -> Optional[Dict[str, str]]:
    if not isinstance(raw, dict):
        return None

    notebook_id = str(raw.get("id") or raw.get("ID") or "").strip()
    name = str(raw.get("name") or "").strip()
    path = str(raw.get("path") or "").strip()

    if not notebook_id and not name and not path:
        return None

    return {
        "id": notebook_id,
        "name": name,
        "path": path,
    }

_publish_context(globals())

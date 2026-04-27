# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


_MAC_RECENT_RECORDS_CACHE_TTL_SEC = 120.0
_MAC_RECENT_RECORDS_CACHE: Dict[str, Any] = {
    "sig": None,
    "timestamp": 0.0,
    "records": None,
}


def _recent_cache_file_signature() -> Optional[Tuple[int, int]]:
    try:
        st = os.stat(_MAC_ONENOTE_RESOURCEINFOCACHE_JSON)
        return int(st.st_mtime_ns), int(st.st_size)
    except Exception:
        return None


def _recent_notebook_records_from_cache_with_timeout(
    timeout_sec: float = 0.8,
) -> Optional[List[Dict[str, Any]]]:
    global _MAC_RECENT_CACHE_TIMED_OUT
    if _MAC_RECENT_CACHE_TIMED_OUT:
        return None

    file_sig = _recent_cache_file_signature()
    if file_sig is None:
        _MAC_RECENT_RECORDS_CACHE.update(
            {"sig": None, "timestamp": time.monotonic(), "records": []}
        )
        return []
    cached_records = _MAC_RECENT_RECORDS_CACHE.get("records")
    if _MAC_RECENT_RECORDS_CACHE.get("sig") == file_sig and isinstance(cached_records, list):
        age = time.monotonic() - float(_MAC_RECENT_RECORDS_CACHE.get("timestamp") or 0.0)
        if age <= _MAC_RECENT_RECORDS_CACHE_TTL_SEC:
            return [dict(record) for record in cached_records]

    reader_script = r'''
import json
import os
from pathlib import Path
from urllib.parse import unquote, urlparse
path = Path(os.path.expanduser("""__CACHE_PATH__"""))
if not path.is_file():
    print("[]")
    raise SystemExit(0)

payload = json.loads(path.read_text(encoding="utf-8"))
entries = payload.get("ResourceInfoCache") if isinstance(payload, dict) else []
if not isinstance(entries, list):
    print("[]")
    raise SystemExit(0)

def clean_field(value):
    return " ".join(str(value or "").strip().split())

def normalize_text(value):
    return " ".join(str(value or "").strip().split()).casefold()

def supported_recent_netloc(netloc):
    host = str(netloc or "").strip().lower()
    if not host:
        return False
    if host in {"d.docs.live.net", "onedrive.live.com", "1drv.ms"}:
        return True
    return host.endswith(".sharepoint.com") or host.endswith(".sharepoint-df.com")

records = []
seen = set()
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
    if not supported_recent_netloc(parsed.netloc):
        continue
    notebook_name = clean_field(item.get("Title") or item.get("title") or "")
    if not notebook_name:
        notebook_name = clean_field(unquote(parsed.path.rstrip("/").split("/")[-1]))
    key = normalize_text(notebook_name)
    if not key or key in seen:
        continue
    seen.add(key)
    records.append({
        "name": notebook_name,
        "url": raw_url,
        "last_accessed_at": int(item.get("LastAccessedAt") or 0),
        "source": "MAC_RECENT_CACHE",
    })
print(json.dumps(records, ensure_ascii=False))
'''.replace("__CACHE_PATH__", str(_MAC_ONENOTE_RESOURCEINFOCACHE_JSON).replace("\\", "\\\\").replace('"', '\\"'))

    last_error = ""
    for executable in _MAC_RECENT_CACHE_READER_PATHS:
        if not executable or not os.path.exists(executable):
            continue
        try:
            completed = subprocess.run(
                [executable, "-c", reader_script],
                text=True,
                capture_output=True,
                encoding="utf-8",
                errors="replace",
                timeout=max(0.25, float(timeout_sec)),
            )
        except subprocess.TimeoutExpired:
            _MAC_RECENT_CACHE_TIMED_OUT = True
            _append_macos_debug_log(
                "[DBG][MAC][RECENT_CACHE] subprocess-timeout "
                f"timeout={timeout_sec:.2f}s"
            )
            return None
        except Exception as exc:
            last_error = str(exc)
            continue
        if completed.returncode != 0:
            last_error = (completed.stderr or completed.stdout or "").strip()
            continue
        try:
            data = json.loads(completed.stdout or "[]")
        except Exception as exc:
            last_error = str(exc)
            continue
        if not isinstance(data, list):
            return []
        records = [
            dict(record)
            for record in data
            if isinstance(record, dict)
            and str(record.get("name") or "").strip()
        ]
        _MAC_RECENT_RECORDS_CACHE.update(
            {
                "sig": file_sig,
                "timestamp": time.monotonic(),
                "records": [dict(record) for record in records],
            }
        )
        return records

    if last_error:
        raise MacAutomationError(last_error)
    return []


def _onenote_protocol_url_from_web_url(web_url: str) -> str:
    parsed = urlparse(str(web_url or "").strip())
    if not parsed.scheme or not parsed.netloc:
        return ""
    path = quote(unquote(parsed.path or ""), safe="/-_.~")
    query = f"?{parsed.query}" if parsed.query else ""
    fragment = f"#{parsed.fragment}" if parsed.fragment else ""
    return f"onenote:{parsed.scheme}://{parsed.netloc}{path}{query}{fragment}"


def recent_notebook_records(
    window: Optional[MacWindow] = None,
    *,
    fast: bool = False,
) -> List[Dict[str, Any]]:
    try:
        records = _recent_notebook_records_from_cache_with_timeout()
    except Exception:
        records = []
    if records is None:
        print("[DBG][MAC][RECENT_CACHE] timeout; fallback=dialog")
        records = []
    if records:
        return [dict(record) for record in records]

    if window is None:
        return []

    names = recent_notebook_names(window, fast=fast)
    return [
        {
            "name": name,
            "url": "",
            "last_accessed_at": 0,
            "source": "MAC_RECENT_DIALOG",
        }
        for name in names
        if str(name or "").strip()
    ]


def open_recent_notebook_record(
    window: Optional[MacWindow],
    record: Dict[str, Any],
    wait_for_visible: bool = True,
    fast: bool = False,
) -> bool:
    name = _clean_field(str((record or {}).get("name") or ""))
    web_url = str((record or {}).get("url") or "").strip()
    protocol_url = _onenote_protocol_url_from_web_url(web_url)

    if protocol_url:
        try:
            open_url_in_system(protocol_url)
            return True
        except Exception:
            pass

    if window is None or not name:
        return False
    return open_recent_notebook_by_name(
        window,
        name,
        wait_for_visible=wait_for_visible,
        fast=fast,
    )


def _ax_onenote_notebook_dialog_roots(window: Optional[MacWindow]) -> List[c_void_p]:
    roots = _ax_window_roots_for_onenote(window)
    if not roots:
        return []

    dialog_roots: List[c_void_p] = []
    other_roots: List[c_void_p] = []
    for root in roots:
        title = _ax_text_attribute(root, "AXTitle")
        if _is_recent_notebook_dialog_title(title):
            dialog_roots.append(root)
        else:
            other_roots.append(root)
    _release_ax_refs(other_roots)
    return dialog_roots


def _ax_find_dialog_element(
    window: Optional[MacWindow],
    *,
    roles: Sequence[str],
    labels: Sequence[str],
    max_depth: int = 12,
    timeout_sec: float = 4.0,
) -> Optional[c_void_p]:
    wanted_roles = {str(role or "").strip() for role in roles if str(role or "").strip()}
    wanted_labels = [
        _normalize_text(str(label or ""))
        for label in labels
        if _normalize_text(str(label or ""))
    ]
    if not (wanted_roles and wanted_labels):
        return None

    roots = _ax_onenote_notebook_dialog_roots(window)
    if not roots:
        return None
    deadline = time.monotonic() + max(0.5, float(timeout_sec or 0.0))
    node_count = 0

    def _matches(element: c_void_p) -> bool:
        combined_parts = []
        for attr_name in ("AXTitle", "AXValue", "AXDescription", "AXHelp"):
            text = _ax_text_attribute(element, attr_name)
            if text:
                combined_parts.append(text)
        combined = _normalize_text(" ".join(combined_parts))
        return any(label and label in combined for label in wanted_labels)

    def _visit(element: c_void_p, depth: int) -> Optional[c_void_p]:
        nonlocal node_count
        if (
            not element
            or depth > max_depth
            or node_count >= 1800
            or time.monotonic() > deadline
        ):
            return None
        node_count += 1
        role = _ax_text_attribute(element, "AXRole")
        if role in wanted_roles and _matches(element):
            retained = _cf_retain(element)
            if retained:
                return retained

        children = _ax_array_attribute(element, "AXChildren")
        try:
            for child in children:
                result = _visit(child, depth + 1)
                if result:
                    return result
        finally:
            _release_ax_refs(children)
        return None

    try:
        for root in roots:
            result = _visit(root, 0)
            if result:
                return result
    finally:
        _release_ax_refs(roots)
    return None


def _ax_press_dialog_radio(
    window: Optional[MacWindow],
    labels: Sequence[str],
    *,
    timeout_sec: float = 4.0,
) -> bool:
    element = _ax_find_dialog_element(
        window,
        roles=("AXRadioButton",),
        labels=labels,
        max_depth=10,
        timeout_sec=timeout_sec,
    )
    if not element:
        return False
    try:
        return _ax_perform_action(element, "AXPress") or _ax_click_element_center(element)
    finally:
        _cf_release(element)


def _ax_click_open_tab_documents_folder(window: Optional[MacWindow]) -> bool:
    roots = _ax_onenote_notebook_dialog_roots(window)
    if not roots:
        return False
    wanted_names = {_normalize_text("문서"), _normalize_text("Documents")}
    deadline = time.monotonic() + 3.0
    node_count = 0

    def _visit(element: c_void_p, depth: int) -> Optional[c_void_p]:
        nonlocal node_count
        if not element or depth > 14 or node_count >= 1400 or time.monotonic() > deadline:
            return None
        node_count += 1
        role = _ax_text_attribute(element, "AXRole")
        if role == "AXStaticText":
            value = _clean_field(_ax_text_attribute(element, "AXValue"))
            desc = _clean_field(_ax_text_attribute(element, "AXDescription"))
            if (
                _normalize_text(value) in wanted_names
                and "d.docs.live.net" in desc.casefold()
            ):
                retained = _cf_retain(element)
                if retained:
                    return retained
        children = _ax_array_attribute(element, "AXChildren")
        try:
            for child in children:
                result = _visit(child, depth + 1)
                if result:
                    return result
        finally:
            _release_ax_refs(children)
        return None

    try:
        for root in roots:
            target = _visit(root, 0)
            if target:
                try:
                    return _ax_click_element_center(target, click_count=2)
                finally:
                    _cf_release(target)
    finally:
        _release_ax_refs(roots)
    return False

_publish_context(globals())

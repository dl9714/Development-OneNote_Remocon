# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _get_open_notebook_records_via_com(
    refresh: bool = False,
    max_age_sec: float = _OPEN_NOTEBOOK_RECORDS_CACHE_TTL_SEC,
) -> List[Dict[str, str]]:
    now = time.monotonic()
    cache_records = _OPEN_NOTEBOOK_RECORDS_CACHE.get("records") or []
    cache_expires_at = float(_OPEN_NOTEBOOK_RECORDS_CACHE.get("expires_at") or 0.0)
    if not refresh and now < cache_expires_at:
        return [dict(record) for record in cache_records]

    if IS_MACOS:
        def _coerce_last_accessed_at(value: Any) -> int:
            try:
                return max(0, int(value or 0))
            except Exception:
                return 0

        def _merge_record(existing: Dict[str, Any], incoming: Dict[str, Any]) -> None:
            incoming_url = str(incoming.get("url") or "").strip()
            incoming_path = str(incoming.get("path") or "").strip()
            incoming_source = str(incoming.get("source") or "").strip()
            incoming_last_accessed_at = _coerce_last_accessed_at(
                incoming.get("last_accessed_at")
            )
            if not str(existing.get("url") or "").strip() and incoming_url:
                existing["url"] = incoming_url
            if not str(existing.get("path") or "").strip() and incoming_path:
                existing["path"] = incoming_path
            if not str(existing.get("source") or "").strip() and incoming_source:
                existing["source"] = incoming_source
            if incoming_last_accessed_at > _coerce_last_accessed_at(
                existing.get("last_accessed_at")
            ):
                existing["last_accessed_at"] = incoming_last_accessed_at

        def _enrich_mac_records(base_records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
            ordered_records: List[Dict[str, Any]] = []
            records_by_key: Dict[str, Dict[str, Any]] = {}
            for raw_record in base_records:
                name = str((raw_record or {}).get("name") or "").strip()
                key = _normalize_notebook_name_key(name)
                if not key:
                    continue
                existing = records_by_key.get(key)
                if existing is None:
                    existing = dict(raw_record)
                    existing["name"] = name
                    existing["url"] = str(existing.get("url") or "").strip()
                    existing["last_accessed_at"] = _coerce_last_accessed_at(
                        existing.get("last_accessed_at")
                    )
                    records_by_key[key] = existing
                    ordered_records.append(existing)
                    continue
                _merge_record(existing, raw_record)

            known_sources: List[List[Dict[str, Any]]] = []
            try:
                cache_records = [
                    dict(record)
                    for record in mac_recent_notebook_records(None)
                    if str((record or {}).get("name") or "").strip()
                ]
            except Exception:
                cache_records = []
            if not cache_records:
                try:
                    cache_records = [
                        dict(record)
                        for record in mac_recent_notebook_records(None)
                        if str((record or {}).get("name") or "").strip()
                    ]
                except Exception:
                    cache_records = []
            if cache_records:
                known_sources.append(cache_records)

            try:
                shortcut_records = [
                    dict(record)
                    for record in _collect_onenote_notebook_shortcuts()
                    if str((record or {}).get("name") or "").strip()
                ]
            except Exception:
                shortcut_records = []
            if shortcut_records:
                known_sources.append(shortcut_records)

            for source_records in known_sources:
                source_by_key = {}
                for source_record in source_records:
                    source_name = str((source_record or {}).get("name") or "").strip()
                    source_key = _normalize_notebook_name_key(source_name)
                    if source_key and source_key not in source_by_key:
                        source_by_key[source_key] = dict(source_record)
                for record in ordered_records:
                    record_key = _normalize_notebook_name_key(record.get("name"))
                    matched = source_by_key.get(record_key)
                    if matched is not None:
                        _merge_record(record, matched)

            return ordered_records

        records: List[Dict[str, Any]] = []
        try:
            ax_box: Dict[str, Any] = {}
            ax_done = threading.Event()

            def _read_ax_records() -> None:
                try:
                    wins = [
                        info
                        for info in enumerate_macos_windows(filter_title_substr=None)
                        if is_macos_onenote_window_info(info, os.getpid())
                    ]
                    wins.sort(
                        key=lambda item: (
                            not bool(item.get("frontmost")),
                            item.get("title", ""),
                        )
                    )
                    ax_records: List[Dict[str, Any]] = []
                    for info in wins:
                        win = MacWindow(dict(info))
                        ax_records = [
                            {
                                "id": "",
                                "name": name,
                                "path": name,
                                "url": "",
                                "last_accessed_at": 0,
                                "source": "MAC_AX_OPEN_NOTEBOOKS",
                            }
                            for name in mac_current_open_notebook_names(win)
                            if str(name).strip()
                        ]
                        if ax_records:
                            break
                    ax_box["records"] = ax_records
                except Exception as exc:
                    ax_box["error"] = exc
                finally:
                    ax_done.set()

            threading.Thread(
                target=_read_ax_records,
                name="onenote-mac-open-notebooks-cache",
                daemon=True,
            ).start()
            if ax_done.wait(8.0):
                if "error" in ax_box:
                    raise ax_box["error"]
                records = [dict(record) for record in (ax_box.get("records") or [])]
            else:
                print("[WARN][MAC][NOTEBOOKS] AX open notebook read timed out")
                records = []
        except Exception as e:
            print(f"[WARN][MAC][NOTEBOOKS] {e}")
            records = []

        if records:
            records = _enrich_mac_records(records)

        _OPEN_NOTEBOOK_RECORDS_CACHE["records"] = [dict(record) for record in records]
        _OPEN_NOTEBOOK_RECORDS_CACHE["expires_at"] = now + max(0.0, max_age_sec)
        return [dict(record) for record in records]

    script = """
$one = New-Object -ComObject OneNote.Application
$xml = ''
$one.GetHierarchy('', 4, [ref]$xml)
[xml]$doc = $xml
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace('one', 'http://schemas.microsoft.com/office/onenote/2013/onenote')
$items = @(
  $doc.SelectNodes('//one:Notebook', $ns) | ForEach-Object {
    [pscustomobject]@{
      id = $_.GetAttribute('ID')
      name = $_.GetAttribute('name')
      path = $_.GetAttribute('path')
    }
  }
)
$items | ConvertTo-Json -Compress
"""
    data = _load_json_output(_run_powershell(script, timeout=30))
    if data is None:
        _OPEN_NOTEBOOK_RECORDS_CACHE["records"] = []
        _OPEN_NOTEBOOK_RECORDS_CACHE["expires_at"] = now + max(0.0, max_age_sec)
        return []
    if isinstance(data, dict):
        data = [data]

    records: List[Dict[str, str]] = []
    if not isinstance(data, list):
        _OPEN_NOTEBOOK_RECORDS_CACHE["records"] = []
        _OPEN_NOTEBOOK_RECORDS_CACHE["expires_at"] = now + max(0.0, max_age_sec)
        return records

    for entry in data:
        record = _normalize_notebook_record(entry)
        if record and record.get("name"):
            records.append(record)

    _OPEN_NOTEBOOK_RECORDS_CACHE["records"] = [dict(record) for record in records]
    _OPEN_NOTEBOOK_RECORDS_CACHE["expires_at"] = now + max(0.0, max_age_sec)
    return [dict(record) for record in records]


def _pick_notebook_name_suggestion(
    requested_name: str, records: List[Dict[str, str]]
) -> str:
    requested_key = _normalize_notebook_name_key(requested_name)
    if not requested_key:
        return ""

    best_name = ""
    best_score = 0.0
    for record in records:
        name = str(record.get("name") or "").strip()
        name_key = _normalize_notebook_name_key(name)
        if not name_key:
            continue

        score = difflib.SequenceMatcher(None, requested_key, name_key).ratio()
        if requested_key in name_key or name_key in requested_key:
            score = max(score, 0.93)
        if score > best_score:
            best_score = score
            best_name = name

    return best_name if best_score >= 0.72 else ""


def _collect_root_notebook_names_from_tree(tree_control, limit: int = 32) -> List[str]:
    names: List[str] = []
    seen = set()
    if not tree_control:
        return names
    try:
        roots = list(tree_control.children() or [])
    except Exception:
        roots = []

    for item in roots:
        try:
            name = _extract_primary_accessible_text(item.window_text()).strip()
        except Exception:
            name = ""
        key = _normalize_notebook_name_key(name)
        if not key or key in seen:
            continue
        seen.add(key)
        names.append(name)
        if len(names) >= max(1, limit):
            break
    return names


def _build_notebook_not_found_error(
    requested_name: str, candidate_names: List[str]
) -> str:
    shown_name = requested_name or "알 수 없는 전자필기장"
    records = [{"name": name} for name in (candidate_names or []) if name]
    suggestion = _pick_notebook_name_suggestion(requested_name, records)
    if suggestion:
        return (
            f"전자필기장 '{shown_name}'을(를) 찾지 못했습니다. "
            f"이름이 바뀌었을 수 있습니다. 현재 열려 있는 비슷한 이름: '{suggestion}'."
        )
    return (
        f"전자필기장 '{shown_name}'을(를) 찾지 못했습니다. "
        "현재 연결된 OneNote에서 해당 전자필기장이 보이지 않습니다."
    )


def _resolve_notebook_target_for_activation(
    target: Dict[str, Any], fallback_name: str = ""
) -> Dict[str, Any]:
    requested_name = _strip_stale_favorite_prefix(
        str((target or {}).get("notebook_text") or fallback_name or "").strip()
    )
    requested_id = str((target or {}).get("notebook_id") or "").strip()
    result = {
        "requested_name": requested_name,
        "resolved_name": requested_name,
        "notebook_id": requested_id,
        "renamed": False,
        "should_abort": False,
        "com_failed": False,
        "error": "",
    }

    try:
        records = _get_open_notebook_records_via_com()
    except Exception as e:
        result["com_failed"] = True
        result["error"] = str(e)
        return result

    if not records:
        return result

    requested_key = _normalize_notebook_name_key(requested_name)
    records_by_key: Dict[str, List[Dict[str, str]]] = {}
    for record in records:
        record_key = _normalize_notebook_name_key(record.get("name"))
        if record_key:
            records_by_key.setdefault(record_key, []).append(record)

    matched_record = None
    if requested_id:
        matched_record = next(
            (record for record in records if (record.get("id") or "") == requested_id),
            None,
        )

    if matched_record is None and requested_key:
        exact_matches = records_by_key.get(requested_key) or []
        if exact_matches:
            matched_record = exact_matches[0]

    if matched_record is not None:
        resolved_name = str(matched_record.get("name") or requested_name).strip()
        resolved_id = str(matched_record.get("id") or requested_id).strip()
        result["resolved_name"] = resolved_name or requested_name
        result["notebook_id"] = resolved_id
        if requested_name and resolved_name:
            result["renamed"] = (
                _normalize_notebook_name_key(requested_name)
                != _normalize_notebook_name_key(resolved_name)
            )
        return result

    suggestion = _pick_notebook_name_suggestion(requested_name, records)
    shown_name = requested_name or fallback_name or "알 수 없는 전자필기장"
    if suggestion:
        hint = (
            f"전자필기장 '{shown_name}'을(를) 찾지 못했습니다. "
            f"이름이 바뀌었을 수 있습니다. 현재 열려 있는 비슷한 이름: '{suggestion}'."
        )
    else:
        hint = (
            f"전자필기장 '{shown_name}'을(를) 찾지 못했습니다. "
            "이름이 바뀌었거나 현재 열려 있지 않을 수 있습니다. "
            "이름을 바꿨다면 다시 등록해 주세요."
        )

    result["should_abort"] = True
    result["error"] = hint
    return result

_publish_context(globals())

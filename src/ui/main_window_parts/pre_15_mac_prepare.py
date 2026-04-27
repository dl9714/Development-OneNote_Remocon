# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


class OpenAllNotebooksWorkerMacPrepareMixin:

    def _prepare_macos_open_all_records(self, result, state):
        candidate_limited_mode = state["candidate_limited_mode"]
        mac_open_action_label = state["mac_open_action_label"]
        _mac_open_all_debug = state["_mac_open_all_debug"]
        open_names = state["open_names"]
        open_detected_names = state["open_detected_names"]
        sidebar_error = state["sidebar_error"]
        accessibility_trusted = state["accessibility_trusted"]
        recent_records = state["recent_records"]
        shortcut_records = state["shortcut_records"]
        settings_records = state["settings_records"]
        open_tab_records = state["open_tab_records"]

        merged_records_by_key: Dict[str, Dict[str, Any]] = {}

        def _track_open_candidate_sources(
            existing: Dict[str, Any],
            incoming: Dict[str, Any],
        ) -> None:
            ordered_sources: List[str] = []
            for candidate in (
                list(existing.get("_candidate_sources") or [])
                if isinstance(existing.get("_candidate_sources"), list)
                else []
            ):
                token = str(candidate or "").strip()
                if token and token not in ordered_sources:
                    ordered_sources.append(token)
            for token in sorted(_notebook_record_source_hints(existing)):
                if token and token not in ordered_sources:
                    ordered_sources.append(token)
            for token in sorted(_notebook_record_source_hints(incoming)):
                if token and token not in ordered_sources:
                    ordered_sources.append(token)
            existing["_candidate_sources"] = ordered_sources

        def _merge_open_candidate(record: Dict[str, Any], *, allow_new: bool) -> None:
            name = str((record or {}).get("name") or "").strip()
            key = _normalize_notebook_name_key(name)
            if not key:
                return
            existing = merged_records_by_key.get(key)
            if existing is None:
                if allow_new:
                    merged = dict(record)
                    _track_open_candidate_sources(merged, record)
                    merged_records_by_key[key] = merged
                return
            _track_open_candidate_sources(existing, record)
            for field in ("url", "path", "id", "source"):
                if (
                    not str(existing.get(field) or "").strip()
                    and str(record.get(field) or "").strip()
                ):
                    existing[field] = str(record.get(field) or "").strip()
            try:
                incoming_last = int(record.get("last_accessed_at") or 0)
                existing_last = int(existing.get("last_accessed_at") or 0)
                if incoming_last > existing_last:
                    existing["last_accessed_at"] = incoming_last
            except Exception:
                pass

        if candidate_limited_mode and settings_records:
            for record in settings_records:
                _merge_open_candidate(record, allow_new=True)
            for source_records in (recent_records, shortcut_records, open_tab_records):
                for record in source_records:
                    _merge_open_candidate(record, allow_new=False)
        else:
            for source_records in (
                recent_records,
                shortcut_records,
                settings_records,
            ):
                for record in source_records:
                    _merge_open_candidate(record, allow_new=True)

        recent_records = list(merged_records_by_key.values())
        _mac_open_all_debug(
            f"[DBG][OPEN_ALL][MAC] merged-records={len(recent_records)}"
        )
        if not recent_records:
            if len(open_detected_names) > 1:
                result["ok"] = True
                result["verified_open_count"] = len(open_detected_names)
                result["opened_names"] = list(open_detected_names)
                result["remaining_names"] = []
                result["refresh_open_notebooks"] = True
                self.done.emit(result)
                return

            if not accessibility_trusted:
                result["error"] = (
                    "macOS 손쉬운 사용 권한이 현재 앱 빌드에 적용되지 않았습니다. "
                    "개인정보 보호 및 보안 > 손쉬운 사용에서 OneNote_Remocon.app을 "
                    "다시 추가/허용해야 합니다."
                )
                result["refresh_open_notebooks"] = False
                self.done.emit(result)
                return

            if sidebar_error and not open_detected_names:
                result["error"] = (
                    f"{sidebar_error} "
                    f"앱이 멈추지 않도록 {mac_open_action_label} 확인을 건너뜁니다."
                )
                result["refresh_open_notebooks"] = False
                self.done.emit(result)
                return

            result["error"] = (
                "최근 전자필기장 목록을 읽지 못했습니다. "
                "OneNote 최근 목록/캐시/바로가기 후보가 모두 비어 있어 "
                f"{mac_open_action_label} 요청을 건너뜁니다."
            )
            result["refresh_open_notebooks"] = False
            self.done.emit(result)
            return
        records_by_key: Dict[str, Dict[str, Any]] = {}
        pending_records: List[Dict[str, Any]] = []
        for record in recent_records:
            name = str(record.get("name") or "").strip()
            key = _normalize_notebook_name_key(name)
            if not key or key in records_by_key:
                continue
            records_by_key[key] = dict(record)
            if key not in open_names:
                pending_records.append(dict(record))
        overlap_keys = [key for key in records_by_key if key in open_names]
        result["candidate_total"] = len(records_by_key)
        result["pending_count"] = len(pending_records)
        result["already_open_count"] = len(overlap_keys)
        _mac_open_all_debug(
            f"[DBG][OPEN_ALL][MAC] overlap={len(overlap_keys)}",
            f"overlap-sample={[str((records_by_key.get(key) or {}).get('name') or '').strip() for key in overlap_keys[:8]]!r}",
        )
        _mac_open_all_debug(
            f"[DBG][OPEN_ALL][MAC] pending-records={len(pending_records)}"
        )

        total_targets = len(pending_records)
        if candidate_limited_mode:
            self.progress.emit(
                f"{mac_open_action_label} 준비 완료... "
                f"후보 {len(records_by_key)}개, 열 대상 {total_targets}개, "
                f"이미 열림 {len(overlap_keys)}개"
            )
        else:
            self.progress.emit(
                f"{mac_open_action_label} 준비 완료... 대상 {total_targets}개"
            )

        if not pending_records:
            result["ok"] = True
            if overlap_keys:
                result["verified_open_count"] = len(overlap_keys)
                result["opened_names"] = [
                    str((records_by_key.get(key) or {}).get("name") or "").strip()
                    for key in overlap_keys
                    if str((records_by_key.get(key) or {}).get("name") or "").strip()
                ]
            elif open_detected_names:
                result["verified_open_count"] = len(open_detected_names)
                result["opened_names"] = list(open_detected_names)
            result["remaining_names"] = []
            self.done.emit(result)
            return

        failed_names = []
        failed_details = []
        failed_records = []

        bulk_records = []
        ui_pending_records = []
        missing_launch_records = []
        for record in pending_records:
            record_copy = dict(record)
            if str((record or {}).get("url") or "").strip():
                bulk_records.append(record_copy)
            elif candidate_limited_mode and _mac_record_is_app_only_without_launch_info(
                record
            ):
                record_copy["_open_all_name_search"] = True
                ui_pending_records.append(record_copy)
            else:
                ui_pending_records.append(record_copy)
        if candidate_limited_mode:
            self.progress.emit(
                f"{mac_open_action_label} 실행 방식 확정... "
                f"URL 일괄 {len(bulk_records)}개, "
                f"UI 보조 {len(ui_pending_records)}개, "
                f"정보 부족 {len(missing_launch_records)}개"
            )
            _mac_open_all_debug(
                "[DBG][OPEN_ALL][MAC]",
                f"bulk-records={len(bulk_records)}",
                f"ui-records={len(ui_pending_records)}",
                f"missing-launch-records={len(missing_launch_records)}",
                f"missing-sample={[str((record or {}).get('name') or '').strip() for record in missing_launch_records[:8]]!r}",
            )
            if not bulk_records and ui_pending_records:
                self.progress.emit(
                    f"{mac_open_action_label} URL 후보 없음... "
                    "최근 목록 UI 보조 자동화로 진행"
                )
            if not bulk_records and not ui_pending_records and missing_launch_records:
                missing_names = [
                    str((record or {}).get("name") or "").strip()
                    for record in missing_launch_records
                    if str((record or {}).get("name") or "").strip()
                ]
                result["error"] = (
                    f"자동으로 열 URL/최근 목록 정보가 없는 후보 {len(missing_names)}개는 "
                    "앱이 멈추지 않도록 건너뜁니다. "
                    "OneNote에서 해당 전자필기장을 한 번 열거나 열린 전자필기장 새로고침 후 다시 시도하세요."
                )
                result["remaining_names"] = missing_names
                self.done.emit(result)
                return
        state.update({
            "records_by_key": records_by_key,
            "pending_records": pending_records,
            "overlap_keys": overlap_keys,
            "failed_names": failed_names,
            "failed_details": failed_details,
            "failed_records": failed_records,
            "bulk_records": bulk_records,
            "ui_pending_records": ui_pending_records,
            "missing_launch_records": missing_launch_records,
        })
        return state

_publish_context(globals())

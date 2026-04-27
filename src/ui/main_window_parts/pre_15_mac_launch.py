# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


class OpenAllNotebooksWorkerMacLaunchMixin:

    def _launch_macos_open_all_records(self, result, win, state) -> None:
        candidate_limited_mode = state["candidate_limited_mode"]
        mac_open_action_label = state["mac_open_action_label"]
        _mac_open_all_debug = state["_mac_open_all_debug"]
        initial_notebook_name = state["initial_notebook_name"]
        records_by_key = state["records_by_key"]
        failed_names = state["failed_names"]
        failed_details = state["failed_details"]
        failed_records = state["failed_records"]
        bulk_records = state["bulk_records"]
        ui_pending_records = state["ui_pending_records"]
        missing_launch_records = state["missing_launch_records"]

        def _record_prefers_name_search(record: Dict[str, Any]) -> bool:
            return bool(
                record.get("_open_all_name_search")
                or _mac_record_is_app_only_without_launch_info(record)
            )

        def _try_open_macos_ui_record(record: Dict[str, Any]) -> bool:
            name = str(record.get("name") or "").strip()
            if not name:
                return False
            if _record_prefers_name_search(record):
                try:
                    if mac_open_tab_notebook_by_name(
                        win,
                        name,
                        wait_for_visible=False,
                        fast=candidate_limited_mode,
                    ):
                        _mac_open_all_debug(
                            "[DBG][OPEN_ALL][MAC]",
                            "open-tab-name-search",
                            f"name={name!r}",
                        )
                        return True
                except Exception as e:
                    _mac_open_all_debug(
                        "[DBG][OPEN_ALL][MAC]",
                        "open-tab-name-search-error",
                        f"name={name!r}",
                        f"error={e!r}",
                    )
                return False
            try:
                return bool(
                    mac_open_recent_notebook_record(
                        win,
                        record,
                        wait_for_visible=False,
                        fast=candidate_limited_mode,
                    )
                )
            except Exception as e:
                failed_details.append(f"{name}: {e}")
                return False

        if ui_pending_records:
            deduped_ui_records = []
            seen_ui_keys = set()
            for record in ui_pending_records:
                name = str(record.get("name") or "").strip()
                key = _normalize_notebook_name_key(name)
                if not key or key in seen_ui_keys:
                    continue
                seen_ui_keys.add(key)
                deduped_ui_records.append(dict(record))
            ui_pending_records = deduped_ui_records

        def _process_ui_records(records: List[Dict[str, Any]], phase_label: str) -> None:
            ui_total = len(records)
            if records:
                self.progress.emit(
                    f"{mac_open_action_label} {phase_label} 중... {ui_total}개"
                )

            for index, record in enumerate(records, start=1):
                name = str(record.get("name") or "").strip()
                print(f"[DBG][OPEN_ALL][MAC] try-open-ui {index}/{ui_total}: {name}")
                _mac_open_all_debug(
                    "[DBG][OPEN_ALL][MAC]",
                    f"try-open-ui={index}/{ui_total}",
                    f"name={name!r}",
                )
                if self.isInterruptionRequested():
                    result["error"] = "사용자 중단"
                    self.done.emit(result)
                    return

                self.progress.emit(
                    f"{mac_open_action_label} 진행 중... "
                    f"{index}/{ui_total} 시도 - {name}"
                )
                opened = _try_open_macos_ui_record(record)
                if not opened and not _record_prefers_name_search(record):
                    try:
                        notebook_result = _mac_ensure_notebook_context_for_section(
                            win,
                            name,
                            wait_for_visible=False,
                        )
                        opened = bool(notebook_result.get("ok"))
                        _mac_open_all_debug(
                            "[DBG][OPEN_ALL][MAC]",
                            "fallback-context",
                            f"ok={opened}",
                            f"source={str(notebook_result.get('source') or '')!r}",
                            f"name={name!r}",
                        )
                    except Exception as e:
                        _mac_open_all_debug(
                            "[DBG][OPEN_ALL][MAC]",
                            "fallback-context-error",
                            f"name={name!r}",
                            f"error={e!r}",
                        )
                        if not any(detail.startswith(f"{name}:") for detail in failed_details):
                            failed_details.append(f"{name}: {e}")
                print(f"[DBG][OPEN_ALL][MAC] launched={opened} name={name}")
                _mac_open_all_debug(
                    "[DBG][OPEN_ALL][MAC]",
                    f"launched={opened}",
                    f"name={name!r}",
                )

                if not opened:
                    failed_names.append(name)
                    failed_records.append(dict(record))
                    if not any(detail.startswith(f"{name}:") for detail in failed_details):
                        failed_details.append(f"{name}: 최근 전자필기장 창에서 열기 실패")
                    continue

                result["opened_names"].append(name)
                result["opened_count"] += 1
                self.progress.emit(
                    f"{mac_open_action_label} 진행 중... "
                    f"{index}/{ui_total} 요청 완료 - {name}"
                )

        pre_bulk_ui_records: List[Dict[str, Any]] = []
        if candidate_limited_mode and bulk_records and ui_pending_records:
            pre_bulk_ui_records = [
                record
                for record in ui_pending_records
                if _record_prefers_name_search(record)
            ]
            if pre_bulk_ui_records:
                pre_bulk_keys = {
                    _normalize_notebook_name_key(record.get("name"))
                    for record in pre_bulk_ui_records
                }
                ui_pending_records = [
                    record
                    for record in ui_pending_records
                    if _normalize_notebook_name_key(record.get("name")) not in pre_bulk_keys
                ]
                _mac_open_all_debug(
                    "[DBG][OPEN_ALL][MAC]",
                    f"pre-bulk-ui-records={len(pre_bulk_ui_records)}",
                )
                _process_ui_records(pre_bulk_ui_records, "이름 검색 UI 선행")
                if result.get("error"):
                    return

        if bulk_records:
            bulk_total = len(bulk_records)
            self.progress.emit(
                f"{mac_open_action_label} 일괄 실행 중... {bulk_total}개"
            )
            for bulk_index, record in enumerate(bulk_records, start=1):
                if self.isInterruptionRequested():
                    result["error"] = "사용자 중단"
                    self.done.emit(result)
                    return

                name = str(record.get("name") or "").strip()
                try:
                    launched = mac_open_recent_notebook_record(None, record)
                except Exception as e:
                    launched = False
                    failed_details.append(f"{name}: {e}")
                print(
                    "[DBG][OPEN_ALL][MAC]",
                    f"bulk-launched={launched}",
                    f"{bulk_index}/{bulk_total}",
                    f"name={name}",
                )
                if launched:
                    result["opened_names"].append(name)
                    result["opened_count"] += 1
                else:
                    if (
                        not candidate_limited_mode
                        or _mac_record_has_ui_open_hint(record)
                    ):
                        ui_pending_records.append(dict(record))
                    else:
                        failed_names.append(name)
                        failed_records.append(dict(record))
                        failed_details.append(f"{name}: URL 실행 실패")

                time.sleep(0.08)

            self.progress.emit(
                f"{mac_open_action_label} 일괄 실행 완료..."
                f" URL {result['opened_count']}개 요청"
            )
            time.sleep(1.5)

        if ui_pending_records:
            _process_ui_records(ui_pending_records, "최근 목록 UI 실행")
            if result.get("error"):
                return

        if failed_records:
            retry_names = []
            retry_details = []
            retry_records = []
            seen_retry_keys = set()
            for record in failed_records:
                name = str(record.get("name") or "").strip()
                key = _normalize_notebook_name_key(name)
                if not key or key in seen_retry_keys:
                    continue
                seen_retry_keys.add(key)
                retry_records.append(record)

            for retry_index, record in enumerate(retry_records, start=1):
                if self.isInterruptionRequested():
                    result["error"] = "사용자 중단"
                    self.done.emit(result)
                    return

                name = str(record.get("name") or "").strip()
                self.progress.emit(
                    f"{mac_open_action_label} 재시도 중... "
                    f"{retry_index}/{len(retry_records)} - {name}"
                )
                if _is_macos_notebook_visible(name):
                    _mac_open_all_debug(
                        "[DBG][OPEN_ALL][MAC]",
                        "retry-visible",
                        f"name={name!r}",
                    )
                    result["opened_names"].append(name)
                    result["opened_count"] += 1
                    continue

                opened = _try_open_macos_ui_record(record)
                if not opened and not _record_prefers_name_search(record):
                    try:
                        notebook_result = _mac_ensure_notebook_context_for_section(
                            win,
                            name,
                            wait_for_visible=False,
                        )
                        opened = bool(notebook_result.get("ok"))
                        _mac_open_all_debug(
                            "[DBG][OPEN_ALL][MAC]",
                            "retry-fallback-context",
                            f"ok={opened}",
                            f"source={str(notebook_result.get('source') or '')!r}",
                            f"name={name!r}",
                        )
                    except Exception as e:
                        _mac_open_all_debug(
                            "[DBG][OPEN_ALL][MAC]",
                            "retry-fallback-context-error",
                            f"name={name!r}",
                            f"error={e!r}",
                        )
                        if not any(detail.startswith(f"{name}:") for detail in retry_details):
                            retry_details.append(f"{name}: {e}")

                if opened:
                    _mac_open_all_debug(
                        "[DBG][OPEN_ALL][MAC]",
                        "retry-opened",
                        f"name={name!r}",
                    )
                    result["opened_names"].append(name)
                    result["opened_count"] += 1
                else:
                    retry_names.append(name)
                    if not any(detail.startswith(f"{name}:") for detail in retry_details):
                        retry_details.append(f"{name}: 재시도 후에도 열기 실패")

            failed_names = retry_names
            failed_details = retry_details

        if missing_launch_records:
            missing_names = []
            for record in missing_launch_records:
                name = str(record.get("name") or "").strip()
                if not name or name in missing_names:
                    continue
                missing_names.append(name)
            failed_names.extend(missing_names)
            failed_details.extend(
                f"{name}: URL/최근 목록 정보 없음" for name in missing_names
            )

        if failed_names:
            result["error"] = (
                f"자동 열기 미완료 {len(failed_names)}개 - "
                + "; ".join(failed_details[:3])
            )
            result["remaining_names"] = failed_names
            self.done.emit(result)
            return

        restore_name = str(initial_notebook_name or "").strip()
        restore_key = _normalize_notebook_name_key(restore_name)
        restore_record = records_by_key.get(restore_key) if restore_key else None
        if restore_record and str(restore_record.get("url") or "").strip():
            if mac_open_recent_notebook_record(None, restore_record):
                self.progress.emit(
                    f"{mac_open_action_label} 마무리 요청... 원래 전자필기장 복구 - {restore_name}"
                )

        result["ok"] = True
        result["remaining_names"] = []
        self.done.emit(result)
        return

_publish_context(globals())

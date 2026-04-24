# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


class OpenAllNotebooksWorkerWindowsMixin:

    def _run_windows_open_all(self, result, win) -> None:
        windows_candidate_limited_mode = self.candidate_scope == "AGG_UNCHECKED"
        windows_open_action_label = (
            _open_unchecked_notebooks_button_label()
            if windows_candidate_limited_mode
            else "실제 OneNote 전체 열기"
        )
        settings_targets = [
            dict(record)
            for record in self.notebook_candidates
            if str((record or {}).get("name") or "").strip()
        ]
        shortcut_targets = _collect_onenote_notebook_shortcuts()
        merged_targets_by_key: Dict[str, Dict[str, Any]] = {}

        def _merge_windows_target(record: Dict[str, Any], *, allow_new: bool) -> None:
            name = str((record or {}).get("name") or "").strip()
            key = _normalize_notebook_name_key(name)
            if not key:
                return
            existing = merged_targets_by_key.get(key)
            if existing is None:
                if allow_new:
                    merged_targets_by_key[key] = dict(record)
                return
            for field in ("path", "url", "id", "source"):
                if (
                    not str(existing.get(field) or "").strip()
                    and str(record.get(field) or "").strip()
                ):
                    existing[field] = str(record.get(field) or "").strip()

        if windows_candidate_limited_mode and settings_targets:
            for target in settings_targets:
                _merge_windows_target(target, allow_new=True)
            for target in shortcut_targets:
                _merge_windows_target(target, allow_new=False)
        else:
            for target in shortcut_targets:
                _merge_windows_target(target, allow_new=True)
            for target in settings_targets:
                _merge_windows_target(target, allow_new=True)

        shortcut_targets = list(merged_targets_by_key.values())

        if shortcut_targets:
            try:
                open_names = {
                    _normalize_notebook_name_key(name)
                    for name in _get_open_notebook_names_via_com(refresh=True)
                    if _normalize_notebook_name_key(name)
                }
            except Exception as e:
                print(f"[WARN][OPEN_ALL_NOTEBOOKS][COM][LIST] {e}")
                open_names = set()

            targets_by_key = {
                _normalize_notebook_name_key(t.get("name")): t
                for t in shortcut_targets
                if _normalize_notebook_name_key(t.get("name"))
            }
            pending_targets = [
                t
                for t in shortcut_targets
                if _normalize_notebook_name_key(t.get("name")) not in open_names
            ]
            overlap_keys = [key for key in targets_by_key if key in open_names]
            result["candidate_total"] = len(targets_by_key)
            result["pending_count"] = len(pending_targets)
            result["already_open_count"] = len(overlap_keys)

            total_targets = len(pending_targets)
            self.progress.emit(
                f"{windows_open_action_label} 준비 완료... "
                f"후보 {len(targets_by_key)}개, 열 대상 {total_targets}개, "
                f"이미 열림 {len(overlap_keys)}개"
            )

            if not pending_targets:
                result["ok"] = True
                result["remaining_names"] = []
                self.done.emit(result)
                return

            failed_names = []
            failed_details = []

            for index, target in enumerate(pending_targets, start=1):
                if self.isInterruptionRequested():
                    result["error"] = "사용자 중단"
                    self.done.emit(result)
                    return

                name = (target.get("name") or "").strip()
                path = (target.get("path") or "").strip()
                url = (target.get("url") or "").strip()
                if not name or (not path and not url):
                    failed_names.append(name or "이름 없음")
                    failed_details.append(f"{name or '이름 없음'}: 바로가기 정보 없음")
                    continue

                self.progress.emit(
                    f"{windows_open_action_label} 진행 중... {index}/{total_targets} 시도 - {name}"
                )
                try:
                    step = _open_notebook_shortcut_via_shell(
                        path,
                        url,
                        name,
                        progress_callback=lambda msg, idx=index, total=total_targets: self.progress.emit(
                            f"{windows_open_action_label} 진행 중... {idx}/{total} - {msg}"
                        ),
                    )
                except Exception as e:
                    step = {"ok": False, "already": False, "name": name, "error": str(e)}

                if step.get("ok"):
                    result["opened_names"].append(name)
                    result["opened_count"] += 1
                    self.progress.emit(
                        f"{windows_open_action_label} 진행 중... {index}/{total_targets} 완료 - {name}"
                    )
                else:
                    failed_names.append(name)
                    failed_details.append(
                        f"{name}: {step.get('error') or '열기 실패'}"
                    )

            if failed_names:
                result["error"] = (
                    f"바로가기 기반 열기 실패 {len(failed_names)}개 - "
                    + "; ".join(failed_details[:3])
                )
                result["remaining_names"] = failed_names
                self.done.emit(result)
                return

            result["ok"] = True
            result["remaining_names"] = []
            self.done.emit(result)
            return

        last_snapshot = None
        stale_rounds = 0
        attempted_norms: Set[str] = set()

        for _ in range(200):
            if self.isInterruptionRequested():
                result["error"] = "사용자 중단"
                self.done.emit(result)
                return

            step = _open_one_notebook_from_backstage(win, attempted_norms)
            if step is None:
                result["error"] = "OneNote의 '전자 필기장 열기' 화면을 찾지 못했습니다."
                self.done.emit(result)
                return

            visible_names = step.get("visible_names") or []
            if step.get("done"):
                result["ok"] = True
                result["remaining_names"] = []
                self.done.emit(result)
                return

            attempted_norm = (step.get("attempted_norm") or "").strip()
            current_snapshot = tuple(_normalize_text(name) for name in visible_names[:8])
            if (not attempted_norm) and current_snapshot and current_snapshot == last_snapshot:
                stale_rounds += 1
            else:
                stale_rounds = 0
            last_snapshot = current_snapshot

            if attempted_norm:
                attempted_norms.add(attempted_norm)

            opened_text = (step.get("opened_text") or "").strip()
            if opened_text:
                result["opened_names"].append(opened_text)
                result["opened_count"] += 1
                self.progress.emit(
                    f"실제 OneNote 전체 열기 중... {result['opened_count']}개 - {opened_text}"
                )

            if step.get("error"):
                result["error"] = step["error"]
                result["remaining_names"] = visible_names
                self.done.emit(result)
                return

            if stale_rounds >= 2:
                result["error"] = "전자필기장 목록이 더 이상 변하지 않아 중단했습니다."
                result["remaining_names"] = visible_names
                self.done.emit(result)
                return

        result["error"] = "안전 제한(200개)에 도달해 중단했습니다."
        self.done.emit(result)

_publish_context(globals())

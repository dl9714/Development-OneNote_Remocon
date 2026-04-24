# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin03:

    def _populate_codex_location_group_combo(
        self, selected_group: str = "", selected_section: str = ""
    ) -> None:
        combo = getattr(self, "codex_location_group_combo", None)
        if combo is None:
            return

        notebook = self._codex_location_selected_notebook()
        seen: Set[str] = set()
        groups: List[Dict[str, str]] = []
        for profile in getattr(self, "_codex_location_lookup_targets", []):
            if not isinstance(profile, dict):
                continue
            if notebook and profile.get("notebook") != notebook:
                continue
            group_name = str(profile.get("section_group", "") or "").strip()
            if not group_name or group_name in seen:
                continue
            if profile.get("kind") == "section_group":
                group_profile = dict(profile)
            else:
                group_profile = {
                    "kind": "section_group",
                    "name": group_name,
                    "path": " > ".join(
                        part for part in [profile.get("notebook", ""), group_name] if part
                    ),
                    "notebook": profile.get("notebook", ""),
                    "section_group": group_name,
                    "section": "",
                    "section_group_id": profile.get("section_group_id", ""),
                    "section_id": "",
                }
            groups.append(group_profile)
            seen.add(group_name)

        combo.blockSignals(True)
        combo.clear()
        combo.addItem("섹션 그룹 없음", {
            "kind": "section_group",
            "name": "섹션 그룹 없음",
            "path": notebook,
            "notebook": notebook,
            "section_group": "",
            "section": "",
            "section_group_id": "",
            "section_id": "",
        })
        selected_idx = 0
        for idx, profile in enumerate(groups, start=1):
            label = profile.get("section_group", "") or profile.get("path", "")
            combo.addItem(label, profile)
            if selected_group and profile.get("section_group") == selected_group:
                selected_idx = idx
        combo.setCurrentIndex(selected_idx)
        combo.blockSignals(False)
        self._populate_codex_location_section_combo(selected_section)

    def _populate_codex_location_section_combo(self, selected_section: str = "") -> None:
        combo = getattr(self, "codex_location_section_combo", None)
        if combo is None:
            return

        notebook = self._codex_location_selected_notebook()
        group_name = self._codex_location_selected_group()
        sections = [
            profile
            for profile in getattr(self, "_codex_location_lookup_targets", [])
            if isinstance(profile, dict)
            and profile.get("kind") == "section"
            and (not notebook or profile.get("notebook") == notebook)
            and profile.get("section_group", "") == group_name
        ]

        combo.blockSignals(True)
        combo.clear()
        combo.addItem("섹션 선택", None)
        selected_idx = 0
        for idx, profile in enumerate(sections, start=1):
            label = profile.get("section", "") or profile.get("path", "")
            combo.addItem(label, profile)
            if selected_section and profile.get("section") == selected_section:
                selected_idx = idx
        combo.setCurrentIndex(selected_idx)
        combo.blockSignals(False)

    def _populate_codex_location_lookup_combo(
        self, selected_path: str = ""
    ) -> None:
        targets = getattr(self, "_codex_location_lookup_targets", [])
        notebook_combo = getattr(self, "codex_location_notebook_combo", None)
        if notebook_combo is None:
            return

        selected_profile = None
        matched_selected_path = False
        if selected_path:
            selected_profile = next(
                (
                    profile
                    for profile in targets
                    if isinstance(profile, dict) and profile.get("path") == selected_path
                ),
                None,
            )
            matched_selected_path = selected_profile is not None
        if selected_profile is None:
            selected_profile = self._codex_location_first_profile(kind="section")
        if selected_profile is None:
            selected_profile = self._codex_location_first_profile()

        notebooks: List[str] = []
        seen: Set[str] = set()
        for profile in targets:
            if not isinstance(profile, dict):
                continue
            notebook = str(profile.get("notebook", "") or "").strip()
            if notebook and notebook not in seen:
                notebooks.append(notebook)
                seen.add(notebook)

        selected_notebook = (selected_profile or {}).get("notebook", "")
        notebook_combo.blockSignals(True)
        notebook_combo.clear()
        for notebook in notebooks:
            notebook_combo.addItem(notebook, notebook)
        selected_idx = max(0, notebook_combo.findData(selected_notebook))
        notebook_combo.setCurrentIndex(selected_idx if notebook_combo.count() else -1)
        notebook_combo.blockSignals(False)

        self._populate_codex_location_group_combo(
            (selected_profile or {}).get("section_group", ""),
            (selected_profile or {}).get("section", ""),
        )

        if selected_profile is not None and matched_selected_path:
            self._apply_codex_location_profile(selected_profile)

    def _set_codex_location_lookup_enabled(self, enabled: bool) -> None:
        toggle = getattr(self, "codex_location_lookup_toggle", None)
        lookup_widgets = (
            getattr(self, "codex_location_notebook_combo", None),
            getattr(self, "codex_location_group_combo", None),
            getattr(self, "codex_location_section_combo", None),
        )
        refresh_btn = getattr(self, "codex_location_lookup_refresh_btn", None)

        if toggle is not None:
            if toggle.isChecked() != enabled:
                toggle.blockSignals(True)
                toggle.setChecked(enabled)
                toggle.blockSignals(False)
            toggle.setText("OneNote 조회 ON" if enabled else "OneNote 조회 OFF")

        for widget in (*lookup_widgets, refresh_btn):
            if widget is not None:
                widget.setVisible(enabled)

        notebook_combo = getattr(self, "codex_location_notebook_combo", None)
        if enabled and notebook_combo is not None and notebook_combo.count() == 0:
            current_path = ""
            path_input = getattr(self, "codex_target_path_input", None)
            if path_input is not None:
                current_path = path_input.text().strip()
            if self._load_codex_location_lookup_cache_into_ui(current_path):
                try:
                    count = len(getattr(self, "_codex_location_lookup_targets", []))
                    self.connection_status_label.setText(
                        f"저장된 OneNote 위치 {count}개를 불러왔습니다."
                    )
                except Exception:
                    pass
            else:
                try:
                    self.connection_status_label.setText(
                        "저장된 OneNote 위치가 없습니다. 조회를 눌러 한 번 갱신하세요."
                    )
                except Exception:
                    pass

    def _refresh_codex_location_lookup(self) -> None:
        worker = getattr(self, "_codex_location_lookup_worker", None)
        try:
            if worker is not None and worker.isRunning():
                self.connection_status_label.setText("OneNote 위치 조회가 이미 진행 중입니다.")
                return
        except Exception:
            pass

        toggle = getattr(self, "codex_location_lookup_toggle", None)
        refresh_btn = getattr(self, "codex_location_lookup_refresh_btn", None)
        lookup_widgets = (
            getattr(self, "codex_location_notebook_combo", None),
            getattr(self, "codex_location_group_combo", None),
            getattr(self, "codex_location_section_combo", None),
        )
        current_path = ""
        path_input = getattr(self, "codex_target_path_input", None)
        if path_input is not None:
            current_path = path_input.text().strip()

        if refresh_btn is not None:
            refresh_btn.setEnabled(False)
            refresh_btn.setText("조회 중")
        if toggle is not None:
            toggle.setEnabled(False)
        try:
            self.connection_status_label.setText(
                "OneNote 위치 조회 중... 저장된 위치 목록은 계속 사용할 수 있습니다."
            )
        except Exception:
            pass

        worker = CodexLocationLookupWorker(
            self._codex_onenote_location_lookup_script(),
            timeout=60,
            parent=self,
        )
        self._codex_location_lookup_worker = worker
        self._retain_qthread_until_finished(worker, "_codex_location_lookup_worker")
        worker.done.connect(
            lambda result, selected_path=current_path, lookup_widgets=lookup_widgets, refresh_btn=refresh_btn, toggle=toggle, worker=worker: self._on_codex_location_lookup_done(
                result,
                selected_path,
                lookup_widgets,
                refresh_btn,
                toggle,
                worker,
            )
        )
        worker.start()

    def _on_codex_location_lookup_done(
        self,
        result: Dict[str, Any],
        selected_path: str,
        lookup_widgets: Tuple[Optional[QWidget], ...],
        refresh_btn: Optional[QToolButton],
        toggle: Optional[QToolButton],
        worker: CodexLocationLookupWorker,
    ) -> None:
        active_worker = getattr(self, "_codex_location_lookup_worker", None)
        if active_worker is not None and active_worker is not worker:
            return

        if refresh_btn is not None:
            refresh_btn.setEnabled(True)
            refresh_btn.setText("조회")
        if toggle is not None:
            toggle.setEnabled(True)
        for widget in lookup_widgets:
            if widget is not None:
                widget.setEnabled(True)

        if not result.get("ok"):
            QMessageBox.warning(
                self,
                "OneNote 위치 조회 실패",
                str(result.get("error") or "알 수 없는 오류"),
            )
            return

        try:
            targets = self._codex_location_lookup_targets_from_json_text(
                str(result.get("raw") or "")
            )
            self._codex_location_lookup_targets = targets
            self._save_codex_location_lookup_cache(targets)
            self._populate_codex_location_lookup_combo(selected_path)
            elapsed = max(0.0, float(result.get("elapsed_ms") or 0) / 1000.0)
            try:
                self.connection_status_label.setText(
                    f"OneNote 위치 조회 완료: {len(targets)}개 저장 ({elapsed:.1f}초)"
                )
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "OneNote 위치 조회 결과 처리 실패", str(e))

    def _apply_codex_location_profile(self, profile: Dict[str, str]) -> None:
        if not isinstance(profile, dict):
            return
        self._populate_codex_target_fields(profile)
        self._apply_codex_target_to_request()
        try:
            self._update_codex_status_summary()
            self.connection_status_label.setText(
                f"코덱스 세부 위치 반영: {profile.get('path', '')}"
            )
        except Exception:
            pass

    def _on_codex_location_notebook_selected(self) -> None:
        self._populate_codex_location_group_combo()
        profile = self._codex_location_first_profile(
            kind="notebook",
            notebook=self._codex_location_selected_notebook(),
        )
        if profile is not None:
            self._apply_codex_location_profile(profile)

    def _on_codex_location_group_selected(self) -> None:
        self._populate_codex_location_section_combo()
        combo = getattr(self, "codex_location_group_combo", None)
        profile = combo.currentData() if combo is not None else None
        if isinstance(profile, dict):
            self._apply_codex_location_profile(profile)

    def _on_codex_location_section_selected(self) -> None:
        combo = getattr(self, "codex_location_section_combo", None)
        profile = combo.currentData() if combo is not None else None
        if isinstance(profile, dict):
            self._apply_codex_location_profile(profile)

    def _codex_target_profile_json_text(self) -> str:
        return json.dumps(
            {
                "version": 1,
                "targets": [self._codex_target_from_fields()],
            },
            ensure_ascii=False,
            indent=2,
        )

    def _copy_codex_target_profile_json_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_current_target_copy_text())
        try:
            self.connection_status_label.setText(
                "현재 작업 위치를 한국어로 복사했습니다."
            )
        except Exception:
            pass

    def _codex_all_targets_json_text(self) -> str:
        return json.dumps(
            {
                "version": 1,
                "exported_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "targets": self._load_codex_targets(),
            },
            ensure_ascii=False,
            indent=2,
        )

    def _copy_codex_all_targets_json_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_all_targets_copy_text())
        try:
            self.connection_status_label.setText("저장된 작업 위치 목록을 한국어로 복사했습니다.")
        except Exception:
            pass

_publish_context(globals())

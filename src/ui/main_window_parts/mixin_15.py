# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin15:

    def _build_codex_target_group(self) -> QWidget:
        group = QWidget()
        group.setObjectName("CodexCard")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(8)

        header = QLabel("📍 OneNote 작업 위치 설정")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        desc = QLabel("왼쪽 패널에서 선택한 전자필기장/섹션이 작업 위치로 들어옵니다.")
        desc.setObjectName("CodexHeroSubtitle")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        lookup_row = QHBoxLayout()
        lookup_row.setSpacing(6)
        self.codex_location_lookup_toggle = QToolButton()
        self.codex_location_lookup_toggle.setText("OneNote 조회 OFF")
        self.codex_location_lookup_toggle.setMinimumWidth(0)
        self.codex_location_lookup_toggle.setSizePolicy(
            QSizePolicy.Policy.Ignored,
            QSizePolicy.Policy.Fixed,
        )
        self.codex_location_lookup_toggle.setCheckable(True)
        self.codex_location_lookup_toggle.toggled.connect(
            self._set_codex_location_lookup_enabled
        )
        lookup_row.addWidget(self.codex_location_lookup_toggle, stretch=1)

        self.codex_location_lookup_refresh_btn = QToolButton()
        self.codex_location_lookup_refresh_btn.setText("조회")
        self.codex_location_lookup_refresh_btn.setMinimumWidth(0)
        self.codex_location_lookup_refresh_btn.setSizePolicy(
            QSizePolicy.Policy.Ignored,
            QSizePolicy.Policy.Fixed,
        )
        self.codex_location_lookup_refresh_btn.clicked.connect(
            self._refresh_codex_location_lookup
        )
        lookup_row.addWidget(self.codex_location_lookup_refresh_btn, stretch=1)
        layout.addLayout(lookup_row)

        lookup_form = QGridLayout()
        lookup_form.setHorizontalSpacing(6)
        lookup_form.setVerticalSpacing(6)

        def add_lookup_combo(row: int, label: str, attr: str, handler) -> WheelSafeComboBox:
            lbl = QLabel(label)
            lbl.setMinimumWidth(54)
            combo = WheelSafeComboBox()
            self._configure_codex_lookup_combo(combo)
            combo.currentIndexChanged.connect(handler)
            setattr(self, attr, combo)
            lookup_form.addWidget(lbl, row, 0)
            lookup_form.addWidget(combo, row, 1)
            return combo

        add_lookup_combo(
            0,
            "필기장",
            "codex_location_notebook_combo",
            self._on_codex_location_notebook_selected,
        )
        add_lookup_combo(
            1,
            "섹션그룹",
            "codex_location_group_combo",
            self._on_codex_location_group_selected,
        )
        add_lookup_combo(
            2,
            "섹션",
            "codex_location_section_combo",
            self._on_codex_location_section_selected,
        )
        layout.addLayout(lookup_form)
        self._set_codex_location_lookup_enabled(False)

        form_layout = QGridLayout()
        form_layout.setSpacing(6)

        def add_line(label_text: str, attr: str, placeholder: str = "", hidden: bool = False) -> QLineEdit:
            lbl = QLabel(label_text)
            lbl.setMinimumWidth(64)
            field = QLineEdit()
            field.setPlaceholderText(placeholder)
            setattr(self, attr, field)

            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0,0,0,0)
            row_layout.setSpacing(4)
            row_layout.addWidget(lbl)
            row_layout.addWidget(field, stretch=1)

            setattr(self, f"{attr}_row", row_widget)
            form_layout.addWidget(row_widget, form_layout.rowCount(), 0)
            row_widget.setVisible(not hidden)
            return field

        self.codex_target_name_input = add_line("위치 이름", "codex_target_name_input", "예: 업무일지 전용")
        self.codex_target_path_input = add_line("작업 경로", "codex_target_path_input", "전자필기장 > 섹션그룹 > 섹션")
        self.codex_target_notebook_input = add_line("전자필기장", "codex_target_notebook_input")
        self.codex_target_group_input = add_line("섹션 그룹", "codex_target_group_input")
        self.codex_target_section_input = add_line("섹션", "codex_target_section_input")
        self.codex_target_group_id_input = add_line("그룹 ID", "codex_target_group_id_input", hidden=True)
        self.codex_target_section_id_input = add_line("섹션 ID", "codex_target_section_id_input", hidden=True)
        layout.addLayout(form_layout)

        actions = QHBoxLayout()
        actions.setSpacing(6)

        apply_btn = QToolButton()
        apply_btn.setText("🎯 요청에 넣기")
        apply_btn.setProperty("variant", "primary")
        apply_btn.clicked.connect(self._apply_codex_target_to_request)

        save_btn = QToolButton()
        save_btn.setText("💾 위치 저장")
        save_btn.clicked.connect(self._save_codex_target_profile)

        actions.addWidget(apply_btn, stretch=2)
        actions.addWidget(save_btn, stretch=1)
        layout.addLayout(actions)

        util_buttons = QGridLayout()
        util_buttons.setHorizontalSpacing(6)
        util_buttons.setVerticalSpacing(6)
        target_tools = [
            ("📋 위치 복사", self._copy_codex_target_profile_json_to_clipboard),
            ("🗂️ 저장위치 복사", self._copy_codex_all_targets_json_to_clipboard),
            ("🧭 위치조회 요청", self._copy_codex_onenote_inventory_script_to_clipboard),
            ("📄 페이지목록 요청", self._copy_codex_page_reader_script_to_clipboard),
        ]
        for index, (text, cb) in enumerate(target_tools):
            btn = QToolButton()
            btn.setText(text)
            btn.setMinimumHeight(28)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            btn.clicked.connect(cb)
            util_buttons.addWidget(btn, index // 2, index % 2)
        layout.addLayout(util_buttons)

        for attr in (
            "codex_target_name_input",
            "codex_target_path_input",
            "codex_target_notebook_input",
            "codex_target_group_input",
            "codex_target_section_input",
            "codex_target_group_id_input",
            "codex_target_section_id_input",
        ):
            widget = getattr(self, attr, None)
            if widget is not None:
                widget.textChanged.connect(self._schedule_codex_codegen_previews)
        self.codex_target_path_input.textChanged.connect(self._apply_codex_target_to_request)
        self._refresh_codex_target_combo()
        self._update_codex_codegen_previews()

        return group

    def _codex_visible_request_text(self) -> str:
        action = self.codex_request_action_combo.currentText()
        target = self.codex_request_target_input.text().strip()
        title = self.codex_request_title_input.text().strip()
        body = self.codex_request_body_editor.toPlainText().strip()

        return f"""작업:
{action}

대상 경로:
{target or "생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리"}

제목/이름:
{title or "코덱스 작업"}

본문/내용:
{body or "- 필요한 내용을 여기에 작성한다."}
"""

    def _codex_request_text(self) -> str:
        return self._codex_visible_request_text()

    def _update_codex_request_preview(self) -> None:
        preview = getattr(self, "codex_request_preview", None)
        if preview is None:
            return
        self._set_plain_text_if_changed(preview, self._codex_visible_request_text())

    def _copy_codex_request_to_clipboard(self) -> None:
        text = self._codex_request_text()
        QApplication.clipboard().setText(text)
        try:
            self.connection_status_label.setText("코덱스 OneNote 요청문을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _codex_request_draft_payload(self) -> Dict[str, Any]:
        def line(attr: str) -> str:
            widget = getattr(self, attr, None)
            if widget is None:
                return ""
            try:
                return widget.text().strip()
            except Exception:
                return ""

        def plain(attr: str) -> str:
            widget = getattr(self, attr, None)
            if widget is None:
                return ""
            try:
                return widget.toPlainText()
            except Exception:
                return ""

        action = ""
        action_combo = getattr(self, "codex_request_action_combo", None)
        if action_combo is not None:
            try:
                action = action_combo.currentText().strip()
            except Exception:
                action = ""

        preset_key = ""
        preset_combo = getattr(self, "codex_request_preset_combo", None)
        if preset_combo is not None:
            try:
                preset_key = str(preset_combo.currentData() or "")
            except Exception:
                preset_key = ""

        return {
            "version": 1,
            "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "preset": preset_key,
            "request": {
                "action": action,
                "target": line("codex_request_target_input"),
                "title": line("codex_request_title_input"),
                "body": plain("codex_request_body_editor"),
            },
            "target_profile": self._codex_target_from_fields(),
        }

    def _save_codex_request_draft(self) -> None:
        try:
            path = self._codex_request_draft_path()
            self._write_json_file_atomic(path, self._codex_request_draft_payload())
            try:
                self.connection_status_label.setText(f"코덱스 요청 초안 저장 완료: {path}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "요청 초안 저장 실패", str(e))

    def _load_codex_request_draft(self) -> None:
        path = self._codex_request_draft_path()
        if not os.path.exists(path):
            QMessageBox.information(self, "요청 초안 없음", "저장된 코덱스 요청 초안이 없습니다.")
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            request = payload.get("request", {}) if isinstance(payload, dict) else {}
            target_profile = payload.get("target_profile", {}) if isinstance(payload, dict) else {}
            if not isinstance(request, dict):
                request = {}
            if not isinstance(target_profile, dict):
                target_profile = {}

            action_combo = getattr(self, "codex_request_action_combo", None)
            if action_combo is not None:
                idx = action_combo.findText(str(request.get("action", "")))
                if idx >= 0:
                    action_combo.setCurrentIndex(idx)

            target_input = getattr(self, "codex_request_target_input", None)
            if target_input is not None:
                target_input.setText(str(request.get("target", "")))

            title_input = getattr(self, "codex_request_title_input", None)
            if title_input is not None:
                title_input.setText(str(request.get("title", "")))

            body_editor = getattr(self, "codex_request_body_editor", None)
            if body_editor is not None:
                body_editor.setPlainText(str(request.get("body", "")))

            if target_profile:
                self._populate_codex_target_fields(target_profile)
            self._update_codex_codegen_previews()
            try:
                self.connection_status_label.setText(f"코덱스 요청 초안 불러오기 완료: {path}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "요청 초안 불러오기 실패", str(e))

    def _codex_execution_checklist_text(self) -> str:
        payload = self._codex_request_draft_payload()
        request = payload.get("request", {})
        target = payload.get("target_profile", {})
        return f"""# OneNote 작업 실행 순서
생성 시각: {time.strftime("%Y-%m-%d %H:%M:%S")}

## 현재 작업

- 작업: {request.get("action", "") or "미지정"}
- 제목/이름: {request.get("title", "") or "미지정"}
- 대상 경로: {request.get("target", "") or target.get("path", "") or "미지정"}

## 실행 전

- 대상 전자필기장/섹션/섹션 그룹이 맞는지 확인한다.
- 삭제/이동 작업은 작업 전 구조를 기록한다.
- 내부 처리 방식은 코덱스 전용 지침을 따른다.

## 실행 중

- 실제 OneNote 조작은 코덱스 전용 지침의 작업별 절차를 따른다.
- 사용자 스킬은 형식/내용 처리 기준으로만 적용한다.
- 예외가 나면 OneNote 상태와 대상 캐시를 다시 조회한다.

## 실행 후 검증

- 코덱스 전용 지침의 검증 기준으로 결과를 확인한다.
- 생성/수정한 제목과 본문 일부가 실제 반영됐는지 확인한다.
- 검증 결과와 다음 확인 항목을 사용자에게 짧게 보고한다.
"""

    def _copy_codex_execution_checklist_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_execution_checklist_text())
        try:
            self.connection_status_label.setText("OneNote 작업 실행 순서를 한국어로 복사했습니다.")
        except Exception:
            pass

    def _append_clipboard_to_codex_request_body(self) -> None:
        body_editor = getattr(self, "codex_request_body_editor", None)
        if body_editor is None:
            return
        clip = QApplication.clipboard().text().strip()
        if not clip:
            QMessageBox.information(self, "클립보드 비어 있음", "요청 본문에 추가할 클립보드 텍스트가 없습니다.")
            return

        current = body_editor.toPlainText().rstrip()
        stamp = time.strftime("%Y-%m-%d %H:%M:%S")
        addition = f"클립보드 자료 ({stamp}):\n{clip}"
        body_editor.setPlainText((current + "\n\n" + addition).strip())
        self._update_codex_codegen_previews()
        try:
            self.connection_status_label.setText("클립보드 텍스트를 코덱스 요청 본문에 추가했습니다.")
        except Exception:
            pass

_publish_context(globals())

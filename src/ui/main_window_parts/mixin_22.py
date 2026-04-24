# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin22:


    def _build_codex_skill_package_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 8, 10, 10)
        layout.setSpacing(8)

        header = QLabel("스킬 패키지")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        desc = QLabel(
            "사용자 스킬, Windows 코덱스 스킬, macOS 코덱스 스킬, 실행 지침을 하나의 템플릿처럼 묶어 저장합니다. "
            "윈도우 자료는 유지하고 맥 흐름을 따로 확장하는 구조입니다."
        )
        desc.setObjectName("CodexPageSubtitle")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        self.codex_skill_package_list = QListWidget()
        self.codex_skill_package_list.setMinimumHeight(72)
        self.codex_skill_package_list.setMaximumHeight(120)
        self.codex_skill_package_list.itemDoubleClicked.connect(
            self._load_selected_codex_skill_package
        )
        layout.addWidget(self.codex_skill_package_list)

        template_row = QHBoxLayout()
        template_row.setSpacing(6)
        template_row.addWidget(QLabel("패키지 템플릿"))
        self.codex_skill_package_template_combo = WheelSafeComboBox()
        self.codex_skill_package_template_combo.setMinimumContentsLength(12)
        for key, template in self._codex_skill_package_templates().items():
            self.codex_skill_package_template_combo.addItem(
                str(template.get("name") or key),
                key,
            )
        template_row.addWidget(self.codex_skill_package_template_combo, stretch=1)
        apply_template_btn = QToolButton()
        apply_template_btn.setText("템플릿 적용")
        apply_template_btn.setMinimumHeight(30)
        apply_template_btn.clicked.connect(self._apply_codex_skill_package_template)
        template_row.addWidget(apply_template_btn)
        layout.addLayout(template_row)

        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(8)
        form.setColumnStretch(1, 1)

        self.codex_skill_package_name_input = QLineEdit()
        self.codex_skill_package_name_input.setPlaceholderText("예: 기본 메모 작성 패키지")
        self.codex_skill_package_name_input.textChanged.connect(
            self._schedule_codex_skill_package_preview
        )
        form.addWidget(QLabel("이름"), 0, 0)
        form.addWidget(self.codex_skill_package_name_input, 0, 1)

        self.codex_skill_package_desc_editor = QTextEdit()
        self.codex_skill_package_desc_editor.setMinimumHeight(54)
        self.codex_skill_package_desc_editor.setPlaceholderText(
            "이 패키지를 언제 쓰는지 적어두세요."
        )
        self.codex_skill_package_desc_editor.textChanged.connect(
            self._schedule_codex_skill_package_preview
        )
        form.addWidget(QLabel("설명"), 1, 0)
        form.addWidget(self.codex_skill_package_desc_editor, 1, 1)

        self.codex_skill_package_user_skill_list = QListWidget()
        self.codex_skill_package_user_skill_list.setSelectionMode(
            QAbstractItemView.SelectionMode.NoSelection
        )
        self.codex_skill_package_user_skill_list.setMinimumHeight(110)
        self.codex_skill_package_user_skill_list.itemChanged.connect(
            self._schedule_codex_skill_package_preview
        )
        form.addWidget(QLabel("사용자 스킬"), 2, 0)
        form.addWidget(self.codex_skill_package_user_skill_list, 2, 1)

        self.codex_skill_package_windows_skill_list = QListWidget()
        self.codex_skill_package_windows_skill_list.setSelectionMode(
            QAbstractItemView.SelectionMode.NoSelection
        )
        self.codex_skill_package_windows_skill_list.setMinimumHeight(108)
        for name in self._codex_skill_package_default_codex_skills(
            CODEX_PLATFORM_WINDOWS
        ):
            self.codex_skill_package_windows_skill_list.addItem(
                self._make_codex_checkable_item(name, name)
            )
        self.codex_skill_package_windows_skill_list.itemChanged.connect(
            self._schedule_codex_skill_package_preview
        )
        form.addWidget(QLabel("Windows 스킬"), 3, 0)
        form.addWidget(self.codex_skill_package_windows_skill_list, 3, 1)

        self.codex_skill_package_macos_skill_list = QListWidget()
        self.codex_skill_package_macos_skill_list.setSelectionMode(
            QAbstractItemView.SelectionMode.NoSelection
        )
        self.codex_skill_package_macos_skill_list.setMinimumHeight(108)
        for name in self._codex_skill_package_default_codex_skills(
            CODEX_PLATFORM_MACOS
        ):
            self.codex_skill_package_macos_skill_list.addItem(
                self._make_codex_checkable_item(name, name)
            )
        self.codex_skill_package_macos_skill_list.itemChanged.connect(
            self._schedule_codex_skill_package_preview
        )
        form.addWidget(QLabel("macOS 스킬"), 4, 0)
        form.addWidget(self.codex_skill_package_macos_skill_list, 4, 1)

        self.codex_skill_package_extra_skills_editor = QTextEdit()
        self.codex_skill_package_extra_skills_editor.setMinimumHeight(54)
        self.codex_skill_package_extra_skills_editor.setPlaceholderText(
            "혼합 운영용 추가 스킬을 한 줄에 하나씩 추가"
        )
        self.codex_skill_package_extra_skills_editor.textChanged.connect(
            self._schedule_codex_skill_package_preview
        )
        form.addWidget(QLabel("추가/혼합"), 5, 0)
        form.addWidget(self.codex_skill_package_extra_skills_editor, 5, 1)

        self.codex_skill_package_instructions_editor = QTextEdit()
        self.codex_skill_package_instructions_editor.setMinimumHeight(96)
        self.codex_skill_package_instructions_editor.setPlaceholderText(
            "코덱스 지침을 한 줄에 하나씩 입력. 예: 현재 보이는 위치 먼저 확인, 쓰기 직후 자동 검증"
        )
        self.codex_skill_package_instructions_editor.textChanged.connect(
            self._schedule_codex_skill_package_preview
        )
        form.addWidget(QLabel("코덱스 지침"), 6, 0)
        form.addWidget(self.codex_skill_package_instructions_editor, 6, 1)

        layout.addLayout(form)

        self.codex_skill_package_preview = QTextEdit()
        self.codex_skill_package_preview.setReadOnly(True)
        self.codex_skill_package_preview.setMinimumHeight(150)
        self.codex_skill_package_preview.setPlaceholderText(
            "스킬 패키지 미리보기가 여기에 표시됩니다."
        )
        layout.addWidget(QLabel("패키지 호출문 미리보기"))
        layout.addWidget(self.codex_skill_package_preview)

        actions = QGridLayout()
        actions.setHorizontalSpacing(6)
        actions.setVerticalSpacing(6)
        specs = [
            ("새 패키지", self._new_codex_skill_package),
            ("저장", self._save_codex_skill_package),
            ("불러오기", self._load_selected_codex_skill_package),
            ("삭제", self._delete_selected_codex_skill_package),
            ("복사", self._copy_codex_skill_package_to_clipboard),
            ("폴더", self._open_codex_skill_packages_folder),
        ]
        for index, (text, cb) in enumerate(specs):
            btn = QToolButton()
            btn.setText(text)
            btn.setMinimumHeight(30)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            if text in ("저장", "복사"):
                btn.setProperty("variant", "primary" if text == "복사" else "secondary")
            btn.clicked.connect(cb)
            actions.addWidget(btn, index // 3, index % 3)
        layout.addLayout(actions)

        self._refresh_codex_skill_package_list()
        if self.codex_skill_package_list.count() > 0:
            self.codex_skill_package_list.setCurrentRow(0)
            self._load_selected_codex_skill_package()
        else:
            self._new_codex_skill_package()
        return card


    def _build_codex_work_order_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 8, 10, 10)
        layout.setSpacing(8)

        header = QLabel("📝 작업 주문서")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        self.codex_work_order_preview = QTextEdit()
        self.codex_work_order_preview.setReadOnly(True)
        self.codex_work_order_preview.setMinimumHeight(96)
        self.codex_work_order_preview.setPlaceholderText("생성된 주문서 미리보기가 여기에 표시됩니다.")
        layout.addWidget(self.codex_work_order_preview)

        actions = QHBoxLayout()
        actions.setSpacing(6)
        copy_btn = QToolButton()
        copy_btn.setText("🚀 주문서 복사")
        copy_btn.setProperty("variant", "primary")
        copy_btn.clicked.connect(self._copy_codex_work_order_to_clipboard)

        save_btn = QToolButton()
        save_btn.setText("💾 저장")
        save_btn.clicked.connect(self._save_codex_work_order)

        for btn in (copy_btn, save_btn):
            btn.setMinimumWidth(0)
            btn.setMinimumHeight(28)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        actions.addWidget(copy_btn, stretch=2)
        actions.addWidget(save_btn, stretch=1)
        layout.addLayout(actions)

        self._update_codex_work_order_preview()
        return card


    def _build_codex_work_order_history_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        header_row = QHBoxLayout()
        header_row.setSpacing(6)

        header = QLabel("주문서 기록")
        header.setObjectName("CodexCardTitle")
        header_row.addWidget(header)
        header_row.addStretch(1)

        refresh_btn = QToolButton()
        refresh_btn.setText("갱신")
        refresh_btn.clicked.connect(self._refresh_codex_work_order_list)
        header_row.addWidget(refresh_btn)

        folder_btn = QToolButton()
        folder_btn.setText("폴더")
        folder_btn.clicked.connect(self._open_codex_requests_folder)
        header_row.addWidget(folder_btn)
        layout.addLayout(header_row)

        search_row = QHBoxLayout()
        search_row.setSpacing(6)
        self.codex_work_order_search_input = QLineEdit()
        self.codex_work_order_search_input.setPlaceholderText("기록 검색")
        self.codex_work_order_search_input.textChanged.connect(self._schedule_refresh_codex_work_order_list)
        search_row.addWidget(self.codex_work_order_search_input, stretch=1)
        layout.addLayout(search_row)

        self.codex_work_order_list = QListWidget()
        self.codex_work_order_list.setMinimumHeight(110)
        self.codex_work_order_list.setMaximumHeight(170)
        self.codex_work_order_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.codex_work_order_list.currentItemChanged.connect(lambda current, prev=None: self._on_codex_work_order_selected(current))
        layout.addWidget(self.codex_work_order_list)

        preview_label = QLabel("선택 주문서 내용")
        preview_label.setObjectName("CodexFieldLabel")
        layout.addWidget(preview_label)

        self.codex_work_order_history_preview = QTextEdit()
        self.codex_work_order_history_preview.setReadOnly(True)
        self.codex_work_order_history_preview.setMinimumHeight(190)
        self.codex_work_order_history_preview.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        layout.addWidget(self.codex_work_order_history_preview, stretch=1)

        actions = QGridLayout()
        actions.setHorizontalSpacing(6)
        actions.setVerticalSpacing(6)

        copy_btn = QToolButton()
        copy_btn.setText("복사")
        copy_btn.setProperty("variant", "secondary")
        copy_btn.clicked.connect(self._copy_selected_codex_work_order_to_clipboard)

        load_btn = QToolButton()
        load_btn.setText("요청 불러오기")
        load_btn.clicked.connect(self._load_selected_codex_work_order_into_request)

        followup_btn = QToolButton()
        followup_btn.setText("후속 요청")
        followup_btn.clicked.connect(self._copy_selected_codex_work_order_followup_prompt)

        open_btn = QToolButton()
        open_btn.setText("열기")
        open_btn.clicked.connect(self._open_selected_codex_work_order_file)

        delete_btn = QToolButton()
        delete_btn.setText("삭제")
        delete_btn.clicked.connect(self._delete_selected_codex_work_order)

        for index, btn in enumerate((copy_btn, load_btn, followup_btn, open_btn, delete_btn)):
            btn.setMinimumHeight(28)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            actions.addWidget(btn, index // 3, index % 3)
        layout.addLayout(actions)

        self._refresh_codex_work_order_list()
        return card


    def _scroll_codex_to_widget(self, attr_name: str) -> None:
        page_mapping = {
            "codex_status_summary_widget": ("codex_remocon_stacked_widget", "_codex_remocon_nav_buttons", 1, 0),
            "codex_quick_tools_widget": ("codex_remocon_stacked_widget", "_codex_remocon_nav_buttons", 1, 0),
            "codex_work_order_widget": ("codex_remocon_stacked_widget", "_codex_remocon_nav_buttons", 1, 0),
            "codex_request_group_widget": ("codex_remocon_stacked_widget", "_codex_remocon_nav_buttons", 1, 1),
            "codex_target_group_widget": ("codex_remocon_stacked_widget", "_codex_remocon_nav_buttons", 1, 1),
            "codex_work_order_history_widget": ("codex_remocon_stacked_widget", "_codex_remocon_nav_buttons", 1, 2),
            "codex_skill_package_widget": ("codex_harness_stacked_widget", "_codex_harness_nav_buttons", 2, 0),
            "codex_context_pack_widget": ("codex_harness_stacked_widget", "_codex_harness_nav_buttons", 2, 0),
            "codex_skill_editor_widget": ("codex_harness_stacked_widget", "_codex_harness_nav_buttons", 2, 1),
            "codex_skill_editor_area": ("codex_harness_stacked_widget", "_codex_harness_nav_buttons", 2, 1),
            "codex_template_group_widget": ("codex_harness_stacked_widget", "_codex_harness_nav_buttons", 2, 2),
            "codex_internal_instructions_widget": ("codex_harness_stacked_widget", "_codex_harness_nav_buttons", 2, 3),
        }
        stack_attr, buttons_attr, workspace_index, idx = page_mapping.get(
            attr_name, ("", "", -1, -1)
        )
        if workspace_index >= 0:
            self._ensure_remocon_workspace_tab_loaded(workspace_index)
        widget = getattr(self, attr_name, None)
        if widget is None:
            return

        stacked = getattr(self, stack_attr, None)
        if stacked is None:
            return

        workspace_tabs = getattr(self, "remocon_workspace_tabs", None)
        if workspace_tabs is not None and workspace_index >= 0:
            try:
                workspace_tabs.setCurrentIndex(workspace_index)
            except Exception:
                pass

        if idx >= 0:
            stacked.setCurrentIndex(idx)
            buttons = getattr(self, buttons_attr, [])
            for i, b in enumerate(buttons):
                b.setChecked(i == idx)

            try:
                scroll_area = stacked.currentWidget()
                content = scroll_area.widget()
                if content:
                    target_y = widget.mapTo(content, widget.rect().topLeft()).y()
                    scroll_area.verticalScrollBar().setValue(max(0, target_y - 16))
            except Exception:
                pass

    def _workspace_mode_from_tab_index(self, index: int) -> str:
        return "codex" if index in (1, 2) else "remocon"

_publish_context(globals())

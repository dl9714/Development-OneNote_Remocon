# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin18:



    def _build_codex_template_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(8)

        header = QLabel("📄 플랫폼별 OneNote 작업 양식")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        subtitle = QLabel(
            "Windows OneNote COM 기준 스킬과 macOS OneNote 화면/UI 기준 스킬을 분리해서 관리합니다."
        )
        subtitle.setObjectName("CodexPageSubtitle")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        platform_row = QHBoxLayout()
        platform_row.setSpacing(6)
        platform_row.addWidget(QLabel("스킬 플랫폼"))
        self.codex_template_platform_combo = WheelSafeComboBox()
        for key, label in _codex_platform_variants():
            self.codex_template_platform_combo.addItem(f"{label} 스킬", key)
        default_platform = _codex_active_platform_key()
        default_index = max(
            0,
            self.codex_template_platform_combo.findData(default_platform),
        )
        self.codex_template_platform_combo.setCurrentIndex(default_index)
        self.codex_template_platform_combo.currentIndexChanged.connect(
            self._on_codex_template_platform_changed
        )
        platform_row.addWidget(self.codex_template_platform_combo, stretch=1)
        layout.addLayout(platform_row)

        self.codex_template_scope_label = QLabel("")
        self.codex_template_scope_label.setObjectName("CodexPageSubtitle")
        self.codex_template_scope_label.setWordWrap(True)
        layout.addWidget(self.codex_template_scope_label)

        row = QHBoxLayout()
        self.codex_template_combo = WheelSafeComboBox()
        self.codex_template_combo.currentIndexChanged.connect(self._update_codex_template_preview)
        row.addWidget(self.codex_template_combo, stretch=1)

        copy_button = QToolButton()
        copy_button.setText("📋 양식 복사")
        copy_button.clicked.connect(self._copy_codex_template_to_clipboard)
        row.addWidget(copy_button)
        layout.addLayout(row)

        self.codex_template_preview = QTextEdit()
        self.codex_template_preview.setReadOnly(True)
        self.codex_template_preview.setMinimumHeight(140)
        layout.addWidget(self.codex_template_preview)
        self._populate_codex_template_combo(default_platform)
        self._update_codex_template_preview()

        return card


    def _build_codex_skill_editor_group(self) -> QWidget:
        root = QWidget()
        root.setObjectName("CodexSkillEditorRoot")
        root.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        main_layout = QVBoxLayout(root)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        list_panel = QWidget()
        list_panel.setObjectName("CodexCard")
        list_panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        list_panel.setMinimumHeight(156)
        list_layout = QVBoxLayout(list_panel)
        list_layout.setContentsMargins(8, 8, 8, 8)
        list_layout.setSpacing(4)

        list_header_row = QHBoxLayout()
        list_header_row.setSpacing(6)
        list_header = QLabel("스킬 목록")
        list_header.setObjectName("CodexCardTitle")
        list_header_row.addWidget(list_header)
        list_header_row.addStretch(1)
        for text, cb in [
            ("새로고침", self._refresh_codex_skill_list),
            ("폴더", self._open_codex_skills_folder),
        ]:
            btn = QToolButton()
            btn.setText(text)
            btn.setMinimumHeight(26)
            btn.clicked.connect(cb)
            list_header_row.addWidget(btn)
        list_layout.addLayout(list_header_row)

        search_row = QHBoxLayout()
        search_row.setSpacing(6)
        self.codex_skill_search_input = QLineEdit()
        self.codex_skill_search_input.setPlaceholderText("스킬/카테고리 검색")
        self.codex_skill_search_input.textChanged.connect(self._schedule_filter_codex_skill_list)
        search_row.addWidget(self.codex_skill_search_input, stretch=2)

        self.codex_skill_order_lookup_input = QLineEdit()
        self.codex_skill_order_lookup_input.setPlaceholderText("주문번호 바로가기")
        self.codex_skill_order_lookup_input.returnPressed.connect(self._select_codex_skill_by_order_input)
        search_row.addWidget(self.codex_skill_order_lookup_input, stretch=1)
        list_layout.addLayout(search_row)

        self.codex_skill_list = QListWidget()
        self.codex_skill_list.itemClicked.connect(self._open_selected_codex_skill_in_editor)
        self.codex_skill_list.itemDoubleClicked.connect(self._open_selected_codex_skill_in_editor)
        self.codex_skill_list.setMinimumHeight(96)
        self.codex_skill_list.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.codex_skill_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.codex_skill_list.setToolTip("카테고리별로 정리된 스킬 목록입니다. 스킬을 클릭하면 바로 아래 편집 영역으로 불러옵니다.")
        list_layout.addWidget(self.codex_skill_list)

        list_actions = QHBoxLayout()
        list_actions.setSpacing(6)
        for index, (text, cb) in enumerate(
            [
                ("주문표 복사", self._copy_codex_skill_order_index_to_clipboard),
            ]
        ):
            btn = QToolButton()
            btn.setText(text)
            btn.setMinimumHeight(26)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            btn.clicked.connect(cb)
            list_actions.addWidget(btn, stretch=1)
        list_layout.addLayout(list_actions)

        editor_area = QWidget()
        self.codex_skill_editor_area = editor_area
        editor_area.setObjectName("CodexCard")
        editor_area.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        editor_layout = QVBoxLayout(editor_area)
        editor_layout.setContentsMargins(8, 8, 8, 8)
        editor_layout.setSpacing(6)

        header_layout = QHBoxLayout()
        header_title = QLabel("스킬 편집")
        header_title.setObjectName("CodexCardTitle")
        header_layout.addWidget(header_title)
        header_layout.addStretch()
        editor_layout.addLayout(header_layout)

        template_row = QHBoxLayout()
        template_row.setSpacing(6)
        self.codex_skill_template_combo = WheelSafeComboBox()
        for key, template in self._codex_skill_templates().items():
            self.codex_skill_template_combo.addItem(template.get("name", key), key)
        template_row.addWidget(QLabel("양식"))
        template_row.addWidget(self.codex_skill_template_combo, stretch=1)

        apply_btn = QToolButton()
        apply_btn.setText("적용")
        apply_btn.setProperty("variant", "secondary")
        apply_btn.clicked.connect(self._apply_codex_skill_template)
        template_row.addWidget(apply_btn)
        editor_layout.addLayout(template_row)

        form_layout = QGridLayout()
        form_layout.setHorizontalSpacing(6)
        form_layout.setVerticalSpacing(6)

        def add_field(row: int, col: int, label: str, field: QLineEdit, placeholder: str = "", stretch: int = 1):
            lbl = QLabel(label)
            lbl.setObjectName("CodexFieldLabel")
            field.setPlaceholderText(placeholder)
            form_layout.addWidget(lbl, row, col)
            form_layout.addWidget(field, row, col + 1)
            form_layout.setColumnStretch(col + 1, stretch)

        self.codex_skill_order_input = QLineEdit()
        self.codex_skill_order_input.textChanged.connect(self._schedule_codex_skill_call_preview)
        self.codex_skill_order_input.setMaximumWidth(92)
        add_field(0, 0, "번호", self.codex_skill_order_input, "SK-001", 0)

        self.codex_skill_name_input = QLineEdit()
        self.codex_skill_name_input.textChanged.connect(self._schedule_codex_skill_call_preview)
        add_field(0, 2, "이름", self.codex_skill_name_input, "스킬명", 2)

        self.codex_skill_trigger_input = QLineEdit()
        self.codex_skill_trigger_input.textChanged.connect(self._schedule_codex_skill_call_preview)
        trigger_label = QLabel("조건")
        trigger_label.setObjectName("CodexFieldLabel")
        self.codex_skill_trigger_input.setPlaceholderText("이 스킬이 사용될 상황")
        form_layout.addWidget(trigger_label, 1, 0)
        form_layout.addWidget(self.codex_skill_trigger_input, 1, 1, 1, 3)

        editor_layout.addLayout(form_layout)

        body_panel = QWidget()
        body_layout = QVBoxLayout(body_panel)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(6)
        body_layout.addWidget(QLabel("스킬 내용"))
        self.codex_skill_body_editor = QTextEdit()
        self.codex_skill_body_editor.setPlaceholderText("사용자 요청에 적용할 처리 기준을 작성하세요.")
        self.codex_skill_body_editor.textChanged.connect(self._schedule_codex_skill_call_preview)
        self.codex_skill_body_editor.setMinimumHeight(220)
        self.codex_skill_body_editor.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self.codex_skill_body_editor.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        body_layout.addWidget(self.codex_skill_body_editor, stretch=1)

        preview_toggle = QToolButton()
        preview_toggle.setText("스킬 요청 미리보기 펼치기")
        preview_toggle.setCheckable(True)
        preview_toggle.setMinimumHeight(28)
        body_layout.addWidget(preview_toggle)

        self.codex_skill_call_preview = QTextEdit()
        self.codex_skill_call_preview.setReadOnly(True)
        self.codex_skill_call_preview.setMinimumHeight(130)
        self.codex_skill_call_preview.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self.codex_skill_call_preview.setVisible(False)
        body_layout.addWidget(self.codex_skill_call_preview)

        def toggle_preview(checked: bool) -> None:
            self.codex_skill_call_preview.setVisible(checked)
            preview_toggle.setText(
                "스킬 요청 미리보기 접기" if checked else "스킬 요청 미리보기 펼치기"
            )

        preview_toggle.toggled.connect(toggle_preview)
        editor_layout.addWidget(body_panel, stretch=1)

        actions = QGridLayout()
        actions.setHorizontalSpacing(6)
        actions.setVerticalSpacing(6)

        save_btn = QToolButton()
        save_btn.setText("저장")
        save_btn.setProperty("variant", "primary")
        save_btn.clicked.connect(self._save_codex_skill_draft)

        new_btn = QToolButton()
        new_btn.setText("새로")
        new_btn.clicked.connect(self._new_codex_skill_draft)

        load_btn = QToolButton()
        load_btn.setText("불러")
        load_btn.clicked.connect(self._load_selected_codex_skill)

        copy_prompt_btn = QToolButton()
        copy_prompt_btn.setText("초안 복사")
        copy_prompt_btn.clicked.connect(self._copy_codex_skill_prompt_to_clipboard)

        delete_btn = QToolButton()
        delete_btn.setText("삭제")
        delete_btn.clicked.connect(self._delete_selected_codex_skill)

        more_btn = QToolButton()
        more_btn.setText("더보기")
        more_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        more_menu = QMenu(more_btn)

        def add_more_action(text: str, cb: Callable[[], None]) -> None:
            action = QAction(text, self)
            action.triggered.connect(lambda checked=False, callback=cb: callback())
            more_menu.addAction(action)

        for text, cb in [
            ("복제", self._duplicate_selected_codex_skill),
            ("클립보드에서 스킬 만들기", self._new_codex_skill_from_clipboard),
            ("현재 요청으로 스킬 만들기", self._new_codex_skill_from_current_request),
            ("선택 스킬 열기", self._open_selected_codex_skill_file),
            ("선택 경로 복사", self._copy_selected_codex_skill_path_to_clipboard),
            ("스킬 요청 복사", self._copy_codex_skill_call_prompt_to_clipboard),
        ]:
            add_more_action(text, cb)
        more_btn.setMenu(more_menu)

        for index, btn in enumerate(
            (save_btn, new_btn, load_btn, copy_prompt_btn, delete_btn, more_btn)
        ):
            btn.setMinimumHeight(30)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            actions.addWidget(btn, index // 3, index % 3)
        editor_layout.addLayout(actions)

        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(8)
        splitter.addWidget(list_panel)
        splitter.addWidget(editor_area)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([220, 620])
        self.codex_skill_editor_splitter = splitter
        main_layout.addWidget(splitter, stretch=1)

        self._refresh_codex_skill_list()
        self._new_codex_skill_draft()
        self._update_codex_skill_call_preview()

        return root

_publish_context(globals())

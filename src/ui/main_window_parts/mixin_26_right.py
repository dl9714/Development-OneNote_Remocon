# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


class MainWindowInitRightMixin:

    def _apply_workspace_button_icons(self) -> None:
        try:
            style = self.style()
            pixmap = QApplication.style().StandardPixmap
            self.refresh_button.setIcon(style.standardIcon(pixmap.SP_BrowserReload))
            self.connect_selected_list_button.setIcon(
                style.standardIcon(pixmap.SP_ArrowForward)
            )
            self.center_button.setIcon(style.standardIcon(pixmap.SP_ArrowRight))
        except Exception:
            pass

    def _build_right_workspace(
        self,
        initial_status,
        COLOR_ACCENT,
        COLOR_ACCENT_HOVER,
        COLOR_ACCENT_PRESSED,
        COLOR_STATUS_BAR,
        search_hint_font_pt,
    ) -> None:
        def _make_right_button_compact(button, *, min_width: int = 44) -> None:
            if not IS_WINDOWS:
                return
            try:
                button.setMinimumWidth(min_width)
                button.setSizePolicy(
                    QSizePolicy.Policy.Minimum,
                    QSizePolicy.Policy.Fixed,
                )
            except Exception:
                pass

        # 3. 오른쪽 패널: 위치정렬/코덱스 탭만 교체되고 1, 2패널은 고정된다.
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        connection_group = QGroupBox(_connection_group_title())
        connection_layout = QVBoxLayout(connection_group)
        if IS_WINDOWS:
            connection_layout.setContentsMargins(6, 8, 6, 6)
            connection_layout.setSpacing(5)

        list_header_layout = QHBoxLayout()
        if IS_WINDOWS:
            list_header_layout.setSpacing(4)
        list_hint_label = QLabel(
            "더블클릭/Enter 연결" if IS_WINDOWS else _onenote_list_hint_text()
        )
        if IS_WINDOWS:
            list_hint_label.setMinimumWidth(0)
            list_hint_label.setSizePolicy(
                QSizePolicy.Policy.Ignored,
                QSizePolicy.Policy.Fixed,
            )
        list_header_layout.addWidget(list_hint_label, alignment=Qt.AlignmentFlag.AlignLeft)
        list_header_layout.addStretch()

        self.refresh_button = QPushButton(" 새로고침")
        if IS_WINDOWS:
            self.refresh_button.setText("새로고침")
            _make_right_button_compact(self.refresh_button, min_width=64)
        self.refresh_button.clicked.connect(self.refresh_onenote_list)
        list_header_layout.addWidget(self.refresh_button)

        self.connect_selected_list_button = QPushButton("선택 연결")
        if IS_WINDOWS:
            self.connect_selected_list_button.setText("연결")
            _make_right_button_compact(self.connect_selected_list_button, min_width=44)
        self.connect_selected_list_button.clicked.connect(
            self._connect_selected_onenote_list_item
        )
        self.connect_selected_list_button.setEnabled(False)
        list_header_layout.addWidget(self.connect_selected_list_button)

        connection_layout.addLayout(list_header_layout)

        self.onenote_list_widget = QListWidget()
        self.onenote_list_widget.addItem("자동 재연결 시도 중...")
        self.onenote_list_widget.installEventFilter(self)
        self.onenote_list_widget.viewport().installEventFilter(self)
        self.onenote_list_widget.itemDoubleClicked.connect(
            self.connect_and_center_from_list_item
        )
        self.onenote_list_widget.itemActivated.connect(
            self.connect_and_center_from_list_item
        )
        self.onenote_list_widget.itemSelectionChanged.connect(
            self._sync_onenote_list_action_buttons
        )
        connection_layout.addWidget(self.onenote_list_widget)
        right_layout.addWidget(connection_group)

        actions_group = QGroupBox(_current_actions_group_title())
        actions_layout = QVBoxLayout(actions_group)

        self.center_button = QPushButton(_primary_restore_button_text())
        self.center_button.setStyleSheet(
            f"""
            QPushButton {{
                background-color: {COLOR_ACCENT};
                color: #111;
                font-weight: bold;
                padding: 8px 16px;
            }}
            QPushButton:hover {{ background-color: {COLOR_ACCENT_HOVER}; }}
            QPushButton:pressed {{ background-color: {COLOR_ACCENT_PRESSED}; }}
            QPushButton:disabled {{
                background-color: #555555;
                color: #999999;
                border: 1px solid #444444;
            }}
        """
        )
        self.center_button.clicked.connect(self.center_selected_item_action)
        self.center_button.setEnabled(False)
        actions_layout.addWidget(self.center_button)

        if IS_WINDOWS or IS_MACOS:
            mac_bulk_actions_layout = QVBoxLayout() if IS_WINDOWS else QHBoxLayout()
            if IS_WINDOWS:
                mac_bulk_actions_layout.setSpacing(5)

            self.open_all_notebooks_button = QPushButton(
                _open_unchecked_notebooks_button_label()
            )
            self.open_all_notebooks_button.setToolTip(_open_unchecked_notebooks_tip())
            if IS_WINDOWS:
                self.open_all_notebooks_button.setText("미체크 열기")
                _make_right_button_compact(self.open_all_notebooks_button, min_width=86)
            self.open_all_notebooks_button.clicked.connect(
                self._open_all_notebooks_from_connected_onenote
            )
            self.open_all_notebooks_button.setEnabled(False)
            mac_bulk_actions_layout.addWidget(self.open_all_notebooks_button)

            actions_layout.addLayout(mac_bulk_actions_layout)

        right_layout.addWidget(actions_group)

        search_group = QGroupBox(_search_group_title())
        search_group_layout = QVBoxLayout(search_group)
        search_group_layout.setSpacing(8)

        project_search_label = QLabel(_project_search_label_text())
        project_search_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        search_group_layout.addWidget(project_search_label)

        module_search_layout = QHBoxLayout()
        self.module_project_search_input = QLineEdit()
        self.module_project_search_input.setPlaceholderText(_project_search_placeholder_text())
        self.module_project_search_input.setClearButtonEnabled(True)
        self.module_project_search_input.textChanged.connect(
            self._schedule_project_buffer_search_highlight
        )
        self.btn_module_project_search_clear = QToolButton()
        self.btn_module_project_search_clear.setText("검색 지우기")
        if IS_WINDOWS:
            self.btn_module_project_search_clear.setText("지우기")
            _make_right_button_compact(self.btn_module_project_search_clear, min_width=52)
        self.btn_module_project_search_clear.clicked.connect(
            self.module_project_search_input.clear
        )
        module_search_layout.addStretch(1)
        module_search_layout.addWidget(self.module_project_search_input, stretch=4)
        module_search_layout.addWidget(self.btn_module_project_search_clear)
        module_search_layout.addStretch(1)
        search_group_layout.addLayout(module_search_layout)

        project_search_hint = QLabel(_project_search_hint_text())
        project_search_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        project_search_hint.setWordWrap(True)
        project_search_hint.setStyleSheet(
            f"color: #B8B8B8; font-size: {search_hint_font_pt};"
        )
        search_group_layout.addWidget(project_search_hint)
        right_layout.addWidget(search_group)

        right_layout.addStretch(1)

        workspace_panel = QWidget()
        workspace_layout = QVBoxLayout(workspace_panel)
        workspace_layout.setContentsMargins(0, 0, 0, 0)
        workspace_layout.setSpacing(0)

        self._workspace_splitter_profiles = {}
        self._active_workspace_splitter_mode = "remocon"
        self.remocon_workspace_tabs = QTabWidget()
        self.remocon_workspace_tabs.setObjectName("RemoconWorkspaceTabs")
        self._codex_remocon_tab_loaded = False
        self._codex_harness_tab_loaded = False
        self.remocon_workspace_tabs.addTab(right_panel, _remocon_workspace_tab_title())
        self.remocon_workspace_tabs.addTab(
            self._build_codex_tab_placeholder("remocon"), "원노트 리모컨"
        )
        self.remocon_workspace_tabs.addTab(
            self._build_codex_tab_placeholder("harness"), "원노트 하네스"
        )
        self.remocon_workspace_tabs.currentChanged.connect(
            self._on_remocon_workspace_tab_changed
        )
        workspace_layout.addWidget(self.remocon_workspace_tabs, stretch=1)
        self.main_splitter.addWidget(workspace_panel)

        self.connection_status_label = QLabel(initial_status)
        self.statusBar().addPermanentWidget(self.connection_status_label)
        self.version_status_label = QLabel(f"v{APP_VERSION} ({APP_BUILD_VERSION})")
        self.version_status_label.setToolTip("앱 버전 / 빌드")
        self.statusBar().addPermanentWidget(self.version_status_label)
        self.statusBar().setStyleSheet(f"background-color: {COLOR_STATUS_BAR};")
        QTimer.singleShot(0, self._apply_workspace_button_icons)

_publish_context(globals())

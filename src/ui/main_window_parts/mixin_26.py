# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts.main_window_style import main_window_stylesheet
from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin26:

    def init_ui(self, initial_status):
        self.setWindowTitle(_main_window_title())

        # --- 메뉴바 생성 ---
        menubar = self.menuBar()
        file_menu = menubar.addMenu("&파일")

        backup_action = QAction("백업하기...", self)
        backup_action.triggered.connect(self._backup_full_settings)
        file_menu.addAction(backup_action)

        restore_action = QAction("복원하기...", self)
        restore_action.triggered.connect(self._restore_full_settings)
        file_menu.addAction(restore_action)

        file_menu.addSeparator()

        shared_settings_action = QAction("공용 설정 JSON 위치 지정...", self)
        shared_settings_action.triggered.connect(self._choose_shared_settings_json)
        file_menu.addAction(shared_settings_action)

        show_settings_path_action = QAction("현재 설정 JSON 위치 보기", self)
        show_settings_path_action.triggered.connect(self._show_settings_json_path)
        file_menu.addAction(show_settings_path_action)

        open_settings_folder_action = QAction("설정 JSON 폴더 열기", self)
        open_settings_folder_action.triggered.connect(self._open_settings_json_folder)
        file_menu.addAction(open_settings_folder_action)

        clear_shared_settings_action = QAction("공용 설정 JSON 연결 해제", self)
        clear_shared_settings_action.triggered.connect(self._clear_shared_settings_json)
        file_menu.addAction(clear_shared_settings_action)

        settings_menu = menubar.addMenu("&환경설정")
        panel_width_action = QAction("패널 폭 조정...", self)
        panel_width_action.triggered.connect(self._open_environment_settings_dialog)
        settings_menu.addAction(panel_width_action)

        special_menu = menubar.addMenu("&특수 기능")
        self.open_all_notebooks_action = QAction(
            _open_unchecked_notebooks_button_label(),
            self,
        )
        self.open_all_notebooks_action.setStatusTip(_open_unchecked_notebooks_tip())
        self.open_all_notebooks_action.triggered.connect(
            self._open_all_notebooks_from_connected_onenote
        )
        self.open_all_notebooks_action.setEnabled(False)
        special_menu.addAction(self.open_all_notebooks_action)

        self.refresh_open_notebooks_action = QAction(
            "현재 열린 전자필기장 종합 새로고침", self
        )
        self.refresh_open_notebooks_action.setStatusTip(
            "macOS OneNote 사이드바의 열린 전자필기장 목록을 백그라운드로 수집해 종합에 반영합니다."
        )
        self.refresh_open_notebooks_action.triggered.connect(
            lambda: self._register_all_notebooks_from_current_onenote(force=True)
        )
        self.refresh_open_notebooks_action.setEnabled(False)
        special_menu.addAction(self.refresh_open_notebooks_action)

        help_menu = menubar.addMenu("&도움말")
        onenote_harness_help_action = QAction("원노트 하네스 도움말", self)
        onenote_harness_help_action.triggered.connect(self._show_onenote_harness_help)
        help_menu.addAction(onenote_harness_help_action)
        help_menu.addSeparator()

        app_info_action = QAction("앱 정보", self)
        app_info_action.setMenuRole(QAction.MenuRole.AboutRole)
        app_info_action.triggered.connect(self._show_app_info)
        help_menu.addAction(app_info_action)

        # --- 스타일시트 정의 (생략) ---
        COLOR_BACKGROUND = "#2E2E2E"
        COLOR_PRIMARY_TEXT = "#E0E0E0"
        COLOR_SECONDARY_TEXT = "#B0B0B0"
        COLOR_GROUPBOX_BG = "#3C3C3C"
        COLOR_ACCENT = "#A6D854"
        COLOR_ACCENT_HOVER = "#B8E966"
        COLOR_ACCENT_PRESSED = "#95C743"
        COLOR_SECONDARY_BUTTON = "#555555"
        COLOR_SECONDARY_BUTTON_HOVER = "#666666"
        COLOR_SECONDARY_BUTTON_PRESSED = "#444444"
        COLOR_LIST_BG = "#252525"
        COLOR_LIST_SELECTED = "#0078D7"
        COLOR_STATUS_BAR = "#252525"
        app_font_stack = _platform_ui_font_stack()
        base_font_pt = "12pt" if IS_MACOS else "11pt"
        status_font_pt = "11pt" if IS_MACOS else "10pt"
        side_label_font_pt = "9pt" if IS_MACOS else "8pt"
        search_hint_font_pt = "10pt" if IS_MACOS else "9pt"

        self.setStyleSheet(
            main_window_stylesheet(
                COLOR_BACKGROUND=COLOR_BACKGROUND,
                COLOR_PRIMARY_TEXT=COLOR_PRIMARY_TEXT,
                COLOR_SECONDARY_TEXT=COLOR_SECONDARY_TEXT,
                COLOR_GROUPBOX_BG=COLOR_GROUPBOX_BG,
                COLOR_ACCENT=COLOR_ACCENT,
                COLOR_SECONDARY_BUTTON=COLOR_SECONDARY_BUTTON,
                COLOR_SECONDARY_BUTTON_HOVER=COLOR_SECONDARY_BUTTON_HOVER,
                COLOR_SECONDARY_BUTTON_PRESSED=COLOR_SECONDARY_BUTTON_PRESSED,
                COLOR_LIST_BG=COLOR_LIST_BG,
                COLOR_LIST_SELECTED=COLOR_LIST_SELECTED,
                COLOR_STATUS_BAR=COLOR_STATUS_BAR,
                app_font_stack=app_font_stack,
                base_font_pt=base_font_pt,
                status_font_pt=status_font_pt,
                side_label_font_pt=side_label_font_pt,
            )
        )

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(10)

        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)  # self로 저장
        self.main_splitter.setChildrenCollapsible(False)
        main_layout.addWidget(self.main_splitter, stretch=1)

        self.left_splitter = QSplitter(Qt.Orientation.Horizontal)  # self로 저장
        self.left_splitter.setChildrenCollapsible(False)

        self._build_left_panels()

        self._build_right_workspace(
            initial_status,
            COLOR_ACCENT,
            COLOR_ACCENT_HOVER,
            COLOR_ACCENT_PRESSED,
            COLOR_STATUS_BAR,
            search_hint_font_pt,
        )

        self._restore_initial_splitter_states()

        # 초기 상태 업데이트
        self._update_move_button_state()
        QTimer.singleShot(0, self._sync_codex_target_from_current_fav_item)

_publish_context(globals())

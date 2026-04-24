# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin23:

    def _current_workspace_splitter_mode(self) -> str:
        tabs = getattr(self, "remocon_workspace_tabs", None)
        if tabs is None:
            return getattr(self, "_active_workspace_splitter_mode", "remocon")
        try:
            return self._workspace_mode_from_tab_index(tabs.currentIndex())
        except Exception:
            return getattr(self, "_active_workspace_splitter_mode", "remocon")

    def _splitter_total_size(self, splitter: Optional[QSplitter], fallback: int) -> int:
        if splitter is None:
            return fallback
        try:
            total = sum(max(0, int(size)) for size in splitter.sizes())
            if total > 0:
                return total
        except Exception:
            pass
        try:
            width = int(splitter.width())
            if width > 0:
                return width
        except Exception:
            pass
        return fallback

    def _capture_workspace_splitter_profile(self, mode: Optional[str] = None) -> None:
        mode = mode or self._current_workspace_splitter_mode()
        if mode not in ("remocon", "codex"):
            return
        profiles = getattr(self, "_workspace_splitter_profiles", {})
        try:
            profiles[mode] = {
                "main": list(self.main_splitter.sizes()),
                "left": list(self.left_splitter.sizes()),
                "codex": list(self.codex_splitter.sizes())
                if getattr(self, "codex_splitter", None) is not None
                else [],
            }
            self._workspace_splitter_profiles = profiles
        except Exception:
            pass

    def _restore_workspace_splitter_profile(self, mode: str) -> bool:
        profile = getattr(self, "_workspace_splitter_profiles", {}).get(mode)
        if not isinstance(profile, dict):
            return False
        try:
            main_sizes = profile.get("main") or []
            left_sizes = profile.get("left") or []
            codex_sizes = profile.get("codex") or []
            if len(main_sizes) >= 2 and sum(main_sizes) > 0:
                self.main_splitter.setSizes(main_sizes)
            if len(left_sizes) >= 2 and sum(left_sizes) > 0:
                self.left_splitter.setSizes(left_sizes)
            if (
                len(codex_sizes) >= 2
                and sum(codex_sizes) > 0
                and getattr(self, "codex_splitter", None) is not None
            ):
                self.codex_splitter.setSizes(codex_sizes)
            return bool(main_sizes or left_sizes or codex_sizes)
        except Exception:
            return False

    def _apply_workspace_splitter_preset(
        self,
        mode: Optional[str] = None,
        *,
        show_status: bool = True,
    ) -> None:
        mode = mode or self._current_workspace_splitter_mode()
        main_splitter = getattr(self, "main_splitter", None)
        left_splitter = getattr(self, "left_splitter", None)
        if main_splitter is None or left_splitter is None:
            return

        fallback_width = max(960, int(self.width()) - 24)
        total = self._splitter_total_size(main_splitter, fallback_width)
        if total <= 0:
            total = fallback_width

        if mode == "codex":
            left_width = min(max(int(total * 0.32), 320), 500)
            if total - left_width < 620:
                left_width = max(280, total - 620)
            status_name = "코덱스"
        elif mode == "balanced":
            left_width = max(300, total // 2)
            status_name = "균등"
        else:
            left_width = min(max(int(total * 0.42), 390), 610)
            if total - left_width < 420:
                left_width = max(300, total - 420)
            status_name = _remocon_workspace_tab_title()

        right_width = max(260, total - left_width)
        main_splitter.setSizes([left_width, right_width])

        if mode == "codex":
            first_width = min(max(int(left_width * 0.42), 145), 215)
        else:
            first_width = min(max(int(left_width * 0.38), 155), 245)
        second_width = max(180, left_width - first_width)
        left_splitter.setSizes([first_width, second_width])

        codex_splitter = getattr(self, "codex_splitter", None)
        if codex_splitter is not None:
            codex_total = self._splitter_total_size(
                codex_splitter,
                max(620, right_width),
            )
            nav_width = min(max(int(codex_total * 0.20), 176), 240)
            codex_splitter.setSizes([nav_width, max(360, codex_total - nav_width)])

        capture_mode = (
            mode
            if mode in ("remocon", "codex")
            else self._current_workspace_splitter_mode()
        )
        self._capture_workspace_splitter_profile(capture_mode)
        if show_status:
            try:
                self.connection_status_label.setText(f"패널 폭 적용: {status_name}")
            except Exception:
                pass

    def _select_workspace_splitter_preset(self, mode: str) -> None:
        tabs = getattr(self, "remocon_workspace_tabs", None)
        target_index = 1 if mode == "codex" else 0
        if tabs is not None and mode in ("remocon", "codex"):
            try:
                tabs.setCurrentIndex(target_index)
            except Exception:
                pass
        self._apply_workspace_splitter_preset(mode, show_status=True)

    def _save_workspace_splitter_layout_now(self) -> None:
        self._capture_workspace_splitter_profile()
        self._save_window_state()
        try:
            self.connection_status_label.setText("현재 패널 폭을 저장했습니다.")
        except Exception:
            pass

    def _on_remocon_workspace_tab_changed(self, index: int) -> None:
        self._ensure_remocon_workspace_tab_loaded(index)
        next_mode = self._workspace_mode_from_tab_index(index)
        self._active_workspace_splitter_mode = next_mode

    def _build_workspace_splitter_toolbar(self) -> QWidget:
        bar = QWidget()
        bar.setObjectName("WorkspaceSplitterToolbar")
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        label = QLabel("패널 폭")
        label.setStyleSheet("font-weight: bold; color: #D8DEE9;")
        layout.addWidget(label)

        fit_btn = QToolButton()
        fit_btn.setText("현재 탭 맞춤")
        fit_btn.clicked.connect(
            lambda: self._apply_workspace_splitter_preset(
                self._current_workspace_splitter_mode(),
                show_status=True,
            )
        )

        remocon_btn = QToolButton()
        remocon_btn.setText(f"{_remocon_workspace_tab_title()} 폭")
        remocon_btn.clicked.connect(
            lambda: self._select_workspace_splitter_preset("remocon")
        )

        codex_btn = QToolButton()
        codex_btn.setText("코덱스 폭")
        codex_btn.clicked.connect(
            lambda: self._select_workspace_splitter_preset("codex")
        )

        balanced_btn = QToolButton()
        balanced_btn.setText("균등")
        balanced_btn.clicked.connect(
            lambda: self._apply_workspace_splitter_preset(
                "balanced",
                show_status=True,
            )
        )

        save_btn = QToolButton()
        save_btn.setText("현재 저장")
        save_btn.clicked.connect(self._save_workspace_splitter_layout_now)

        for btn in (fit_btn, remocon_btn, codex_btn, balanced_btn, save_btn):
            btn.setMinimumHeight(28)
            layout.addWidget(btn)

        layout.addStretch(1)
        return bar

    def _open_environment_settings_dialog(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("환경설정")
        dialog.resize(560, 160)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        panel_group = QGroupBox("패널 폭")
        panel_layout = QVBoxLayout(panel_group)
        panel_layout.setContentsMargins(12, 14, 12, 12)
        panel_layout.setSpacing(10)

        desc = QLabel("1, 2, 3번째 패널 폭을 현재 작업에 맞게 조정합니다.")
        desc.setWordWrap(True)
        panel_layout.addWidget(desc)
        panel_layout.addWidget(self._build_workspace_splitter_toolbar())
        layout.addWidget(panel_group)

        button_row = QHBoxLayout()
        button_row.addStretch(1)
        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(dialog.accept)
        button_row.addWidget(close_btn)
        layout.addLayout(button_row)

        dialog.exec()

_publish_context(globals())

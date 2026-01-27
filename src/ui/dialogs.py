# -*- coding: utf-8 -*-
"""
다이얼로그 모듈

애플리케이션에서 사용하는 다이얼로그 클래스들을 정의합니다.
"""

from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QListWidget
from src.workers.thread_workers import WindowListWorker
from src.core.window_manager import WindowManager


class OtherWindowSelectionDialog(QDialog):
    """OneNote 외의 다른 창을 선택하는 다이얼로그"""

    def __init__(self, my_pid: int, parent=None):
        super().__init__(parent)
        self.my_pid = my_pid
        self.window_manager = WindowManager()
        self.setWindowTitle("연결할 창을 더블클릭하세요.")
        self.setGeometry(400, 400, 500, 420)

        self._setup_ui()
        self._start_scan()

    def _setup_ui(self):
        """UI를 설정합니다."""
        self.layout = QVBoxLayout(self)

        self.tip_label = QLabel("창 목록을 검색 중입니다...")
        self.layout.addWidget(self.tip_label)

        self.other_list_widget = QListWidget()
        self.layout.addWidget(self.other_list_widget)
        self.other_list_widget.hide()

        self.windows_info = []
        self.selected_info = None

        self.other_list_widget.itemDoubleClicked.connect(self._on_window_selected)

    def _start_scan(self):
        """창 스캔을 시작합니다."""
        self.worker = WindowListWorker()
        self.worker.done.connect(self._on_windows_list_ready)
        self.worker.start()

    def _on_windows_list_ready(self, results):
        """창 목록 스캔이 완료되었을 때 호출됩니다."""
        self.tip_label.hide()

        if not results:
            self.tip_label.setText("표시할 창이 없습니다. 다시 시도해 주세요.")
            self.tip_label.show()
            return

        # OneNote가 아닌 창만 필터링
        for r in results:
            pid = r.get("pid")
            if pid == self.my_pid:
                continue
            if not self.window_manager.is_onenote_window(r):
                self.windows_info.append(r)

        self.windows_info.sort(key=lambda r: r.get("title", ""))

        if self.windows_info:
            items = [
                f'{r["title"]}  [{r["class_name"]}] (0x{r["handle"]:X})'
                for r in self.windows_info
            ]
            self.other_list_widget.addItems(items)
            self.other_list_widget.show()
        else:
            self.tip_label.setText("OneNote를 제외한 다른 창이 없습니다.")
            self.tip_label.show()

    def _on_window_selected(self, item):
        """창이 선택되었을 때 호출됩니다."""
        row = self.other_list_widget.currentRow()
        if 0 <= row < len(self.windows_info):
            self.selected_info = self.windows_info[row]
            self.accept()

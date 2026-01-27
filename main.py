#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
OneNote Remocon - 전자필기장 스크롤 리모컨

메인 진입점입니다.
"""

import time

_T0 = time.perf_counter()
def _boot(msg: str):
    dt = (time.perf_counter() - _T0) * 1000.0
    print(f"[BOOT0] {dt:8.1f} ms | {msg}")

_boot("process start")

import sys
_boot("import sys done")
import os

# 프로젝트 루트를 Python 경로에 추가
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QIcon
from src.ui.main_window import OneNoteScrollRemoconApp, ROLE_TYPE
from src.constants import APP_ICON_PATH
from src.utils import resource_path

_boot("import MainWindow done")


def _set_windows_appusermodelid(app_id: str) -> None:
    """
    Windows 작업표시줄/Alt+Tab 아이콘이 '파이썬 기본 아이콘'으로 뜨는 문제를 줄이기 위한 설정.
    (PyInstaller/Qt 조합에서 특히 중요)
    """
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def _apply_app_icon(app: QApplication) -> None:
    icon_path = resource_path(APP_ICON_PATH)
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))


def main():
    _boot("enter main()")
    """애플리케이션을 실행합니다."""
    app = QApplication(sys.argv)
    _boot("QApplication() created")

    # 0) 작업표시줄/Alt+Tab 아이콘 안정화 (Windows)
    if sys.platform.startswith("win"):
        _set_windows_appusermodelid("OneNoteRemocon.OneNote_Remocon")

    # 1) QApplication 레벨 아이콘 (모든 윈도우에 기본 적용)
    _apply_app_icon(app)

    window = OneNoteScrollRemoconApp()
    _boot("MainWindow() created")
    window.show()
    _boot("MainWindow.show() called")

    # 즐겨찾기 더블클릭 동작 설정
    try:
        window.fav_tree.itemDoubleClicked.disconnect()
    except TypeError:
        pass

    def _toggle_group_and_activate_section(item, col):
        node_type = item.data(0, ROLE_TYPE)
        name = item.text(0)
        print(f"[DBG][FAV][DBLCLK][MAIN] type={node_type} name='{name}' childCount={item.childCount()}")
        # ✅ 그룹만 토글, 나머지는 전부 '활성화 로직'으로 넘긴다
        if node_type == "group":
            item.setExpanded(not item.isExpanded())
            return
        window._on_fav_item_double_clicked(item, col)

    window.fav_tree.itemDoubleClicked.connect(_toggle_group_and_activate_section)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()

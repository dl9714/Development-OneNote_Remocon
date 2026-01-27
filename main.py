#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
OneNote Remocon - 전자필기장 스크롤 리모컨

메인 진입점입니다.
"""

import sys
import os

# 프로젝트 루트를 Python 경로에 추가
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt6.QtWidgets import QApplication

# main_window 모듈에서 필요한 것들 import
from src.ui.main_window import OneNoteScrollRemoconApp, ROLE_TYPE


def main():
    """애플리케이션을 실행합니다."""
    app = QApplication(sys.argv)

    window = OneNoteScrollRemoconApp()
    window.show()

    # 즐겨찾기 더블클릭 동작 설정
    try:
        window.fav_tree.itemDoubleClicked.disconnect()
    except TypeError:
        pass

    def _toggle_group_and_activate_section(item, col):
        node_type = item.data(0, ROLE_TYPE)
        if node_type != "section":
            item.setExpanded(not item.isExpanded())
        else:
            window._on_fav_item_double_clicked(item, col)

    window.fav_tree.itemDoubleClicked.connect(_toggle_group_and_activate_section)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()

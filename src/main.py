# -*- coding: utf-8 -*-
"""
OneNote Remocon 메인 진입점

리팩토링된 모듈들을 사용하는 메인 애플리케이션 진입점입니다.
"""

import sys
import os

# 프로젝트 루트를 Python 경로에 추가
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QIcon

# 리팩토링된 모듈들 임포트
from src.constants import APP_ICON_PATH
from src.utils import resource_path

# 기존 메인 앱 클래스를 임포트 (점진적 리팩토링을 위해)
# 나중에는 리팩토링된 버전으로 교체될 것입니다.
from OneNote_Remocon import OneNoteScrollRemoconApp, ROLE_TYPE


def main():
    """메인 애플리케이션을 실행합니다."""
    app = QApplication(sys.argv)

    # 아이콘 설정
    icon_path = resource_path(APP_ICON_PATH)
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    # 메인 윈도우 생성 및 표시
    window = OneNoteScrollRemoconApp()
    window.show()

    # 즐겨찾기 트리 더블클릭 이벤트 재연결
    try:
        window.fav_tree.itemDoubleClicked.disconnect()
    except TypeError:
        pass

    def _toggle_group_and_activate_section(item, col):
        """그룹은 확장/축소하고, 섹션은 활성화합니다."""
        node_type = item.data(0, ROLE_TYPE)
        if node_type != "section":
            item.setExpanded(not item.isExpanded())
        else:
            window._on_fav_item_double_clicked(item, col)

    window.fav_tree.itemDoubleClicked.connect(_toggle_group_and_activate_section)

    # 애플리케이션 실행
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

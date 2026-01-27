# -*- coding: utf-8 -*-
"""
커스텀 위젯 모듈

애플리케이션에서 사용하는 커스텀 위젯들을 정의합니다.
"""

from PyQt6.QtWidgets import QTreeWidget, QAbstractItemView
from PyQt6.QtCore import pyqtSignal, Qt
from src.constants import ROLE_TYPE


class FavoritesTree(QTreeWidget):
    """
    즐겨찾기 트리 위젯

    드래그 앤 드롭, 키보드 단축키를 지원하는 커스텀 트리 위젯입니다.
    """

    # 커스텀 시그널
    structureChanged = pyqtSignal()
    deleteRequested = pyqtSignal()
    renameRequested = pyqtSignal()
    copyRequested = pyqtSignal()
    pasteRequested = pyqtSignal()

    def __init__(self, parent=None):
        """즐겨찾기 트리를 초기화합니다."""
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        """UI를 설정합니다."""
        self.setHeaderHidden(True)
        self.setColumnCount(1)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setIndentation(16)
        self.setAnimated(True)
        self.setExpandsOnDoubleClick(True)

    def dropEvent(self, event):
        """
        드롭 이벤트를 처리합니다.

        섹션 항목이 다른 섹션 항목 위로 드롭되면 형제 항목으로 변경합니다.
        """
        source_item = self.currentItem()
        target_item = self.itemAt(event.position().toPoint())

        if not source_item:
            event.ignore()
            return

        super().dropEvent(event)

        # 섹션을 섹션 위로 드롭한 경우 형제로 만들기
        if target_item and source_item.parent() == target_item:
            source_type = source_item.data(0, ROLE_TYPE)
            target_type = target_item.data(0, ROLE_TYPE)

            if source_type == "section" and target_type == "section":
                moved_item = target_item.takeChild(
                    target_item.indexOfChild(source_item)
                )

                if moved_item:
                    parent_of_target = target_item.parent()
                    if not parent_of_target:
                        parent_of_target = self.invisibleRootItem()
                    target_index = parent_of_target.indexOfChild(target_item)
                    parent_of_target.insertChild(target_index + 1, moved_item)
                    self.setCurrentItem(moved_item)

        self.structureChanged.emit()

    def keyPressEvent(self, event):
        """
        키보드 이벤트를 처리합니다.

        지원하는 단축키:
        - Delete: 항목 삭제
        - F2: 항목 이름 변경
        - Ctrl+C: 항목 복사
        - Ctrl+V: 항목 붙여넣기
        """
        # Delete 키
        if event.key() == Qt.Key.Key_Delete:
            self.deleteRequested.emit()
            event.accept()
            return

        # F2 키
        if event.key() == Qt.Key.Key_F2:
            self.renameRequested.emit()
            event.accept()
            return

        # Ctrl+C / Ctrl+V
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.key() == Qt.Key.Key_C:
                self.copyRequested.emit()
                event.accept()
                return
            if event.key() == Qt.Key.Key_V:
                self.pasteRequested.emit()
                event.accept()
                return

        super().keyPressEvent(event)

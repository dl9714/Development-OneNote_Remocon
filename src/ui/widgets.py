# -*- coding: utf-8 -*-
"""
커스텀 위젯 모듈

애플리케이션에서 사용하는 커스텀 위젯들을 정의합니다.
"""

from PyQt6.QtWidgets import QTreeWidget, QAbstractItemView
from PyQt6.QtCore import pyqtSignal, Qt
import traceback
from src.constants import ROLE_TYPE
from typing import Optional


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
    cutRequested = pyqtSignal()
    undoRequested = pyqtSignal()
    redoRequested = pyqtSignal()

    def __init__(self, parent=None):
        """즐겨찾기 트리를 초기화합니다."""
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        """UI를 설정합니다."""
        self.setHeaderHidden(True)
        self.setColumnCount(1)
        # ✅ 더블클릭 시 '이름 편집'으로 들어가는 기본 동작 차단
        #    (이름 변경은 F2 / 메뉴로만)
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        # ✅ 2패널(모듈영역) 다중 선택 지원 (CTRL/SHIFT)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        # ✅ 키 이벤트(Del/F2/Ctrl+C/V)가 확실히 들어오게 포커스 강화
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
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

            # 섹션(또는 버퍼)이 같은 타입 위로 드롭되면 형제로 이동
            if source_type == target_type and source_type in ["section", "buffer"]:
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
        try:
            # ✅ 디버그: 어떤 트리가 키를 먹는지 추적
            mods_obj = event.modifiers()
            # PyQt6: modifiers가 enum/flags라 int()가 바로 안 되는 케이스가 있어 .value로 정수화
            mods_val = mods_obj.value if hasattr(mods_obj, "value") else mods_obj
            print(
                f"[DBG][TREE][KEY] cls={self.__class__.__name__} key={event.key()} mods={int(mods_val)} hasFocus={self.hasFocus()}"
            )

            # Delete 키
            if event.key() == Qt.Key.Key_Delete:
                print(f"[DBG][TREE][DEL] emit deleteRequested from {self.__class__.__name__}")
                self.deleteRequested.emit()
                event.accept()
                return
        except Exception:
            print("[ERR][TREE][KEY] exception")
            traceback.print_exc()

        # F2 키
        if event.key() == Qt.Key.Key_F2:
            self.renameRequested.emit()
            event.accept()
            return

        # Ctrl 조합 단축키 (편집 중이면 기본 편집 동작 우선)
        # 다른 modifier(Shift 등)와 같이 눌려도 Ctrl을 감지하도록 비트 플래그로 체크
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            # 편집 중(Ctrl+Z 등)은 기본 편집 동작 우선
            if self.state() == QAbstractItemView.State.EditingState:
                super().keyPressEvent(event)
                return

            # Undo / Redo
            if event.key() == Qt.Key.Key_Z:
                if event.modifiers() & Qt.KeyboardModifier.ShiftModifier:
                    print("[DBG][TREE][UNDO] emit redoRequested (Ctrl+Shift+Z)")
                    self.redoRequested.emit()
                else:
                    print("[DBG][TREE][UNDO] emit undoRequested (Ctrl+Z)")
                    self.undoRequested.emit()
                event.accept()
                return
            if event.key() == Qt.Key.Key_Y:
                print("[DBG][TREE][UNDO] emit redoRequested (Ctrl+Y)")
                self.redoRequested.emit()
                event.accept()
                return

            # Cut / Copy / Paste
            if event.key() == Qt.Key.Key_X:
                self.cutRequested.emit()
                event.accept()
                return
            if event.key() == Qt.Key.Key_C:
                self.copyRequested.emit()
                event.accept()
                return
            if event.key() == Qt.Key.Key_V:
                self.pasteRequested.emit()
                event.accept()
                return

        super().keyPressEvent(event)


class BufferTree(FavoritesTree):
    """
    즐겨찾기 버퍼 목록을 관리하는 트리 위젯 (그룹 기능 포함)
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        # 기본적인 설정은 FavoritesTree와 동일

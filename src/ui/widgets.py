# -*- coding: utf-8 -*-
import os
"""
커스텀 위젯 모듈

애플리케이션에서 사용하는 커스텀 위젯들을 정의합니다.
"""

from PyQt6.QtWidgets import (
    QTreeWidget,
    QAbstractItemView,
    QLineEdit,
    QStyledItemDelegate,
    QStyleOptionViewItem,
)
from PyQt6.QtCore import pyqtSignal, Qt, QSize
from PyQt6.QtGui import QColor, QPainter, QPen
import traceback
from src.constants import ROLE_TYPE, ROLE_OPEN_NOTEBOOK
from typing import Optional


TREE_WIDGET_KEY_DEBUG = os.environ.get("ONENOTE_REMOCON_DEBUG_KEYS") == "1"


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
        self.setUniformRowHeights(True)
        self.setAnimated(False)
        self.setExpandsOnDoubleClick(True)

    def dropEvent(self, event):
        """
        드롭 이벤트를 처리합니다.

        섹션 항목이 다른 섹션 항목 위로 드롭되면 형제 항목으로 변경합니다.
        """
        source_item = self.currentItem()
        target_item = self.itemAt(event.position().toPoint())
        drop_position = self.dropIndicatorPosition()

        if not source_item:
            event.ignore()
            return

        root_item = self.invisibleRootItem()
        source_parent_before = source_item.parent() or root_item
        source_row_before = source_parent_before.indexOfChild(source_item)

        super().dropEvent(event)

        # 섹션을 섹션 위로 드롭한 경우 형제로 만들기
        if target_item and source_item.parent() == target_item:
            source_type = source_item.data(0, ROLE_TYPE)
            target_type = target_item.data(0, ROLE_TYPE)

            # 전자필기장/섹션/버퍼가 같은 타입 위로 드롭되면 형제로 이동
            if source_type == target_type and source_type in {
                "section",
                "buffer",
                "notebook",
            }:
                moved_item = target_item.takeChild(
                    target_item.indexOfChild(source_item)
                )

                if moved_item:
                    parent_of_target = target_item.parent()
                    if not parent_of_target:
                        parent_of_target = root_item
                    target_index = parent_of_target.indexOfChild(target_item)
                    insert_index = target_index
                    if (
                        drop_position
                        != QAbstractItemView.DropIndicatorPosition.AboveItem
                    ):
                        insert_index = target_index + 1
                    parent_of_target.insertChild(insert_index, moved_item)
                    self.setCurrentItem(moved_item)

        source_parent_after = source_item.parent() or root_item
        source_row_after = source_parent_after.indexOfChild(source_item)
        if (
            source_parent_after is not source_parent_before
            or source_row_after != source_row_before
        ):
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
            if TREE_WIDGET_KEY_DEBUG:
                print(
                    f"[DBG][TREE][KEY] cls={self.__class__.__name__} key={event.key()} mods={int(mods_val)} hasFocus={self.hasFocus()}"
                )

            # Delete 키
            if event.key() == Qt.Key.Key_Delete:
                if TREE_WIDGET_KEY_DEBUG:
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
                    if TREE_WIDGET_KEY_DEBUG:
                        print("[DBG][TREE][UNDO] emit redoRequested (Ctrl+Shift+Z)")
                    self.redoRequested.emit()
                else:
                    if TREE_WIDGET_KEY_DEBUG:
                        print("[DBG][TREE][UNDO] emit undoRequested (Ctrl+Z)")
                    self.undoRequested.emit()
                event.accept()
                return
            if event.key() == Qt.Key.Key_Y:
                if TREE_WIDGET_KEY_DEBUG:
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


class TreeNameEditDelegate(QStyledItemDelegate):
    """
    트리 이름 변경용 인라인 에디터 높이를 안정적으로 맞춰주는 델리게이트.

    전역 QLineEdit padding이 트리 인라인 편집기에도 그대로 적용되면
    기본 셀 높이보다 편집기가 살짝 커져 텍스트 상/하단이 미세하게 잘릴 수 있다.
    여기서는 이름 변경 에디터에만 조금 더 타이트한 내부 여백과 최소 높이를 적용해
    프로젝트/모듈 트리의 rename UX를 자연스럽게 맞춘다.
    """

    _HORIZONTAL_TEXT_MARGIN = 8
    _VERTICAL_TEXT_MARGIN = 2
    _EDITOR_EXTRA_HEIGHT = 2
    _OPEN_NOTEBOOK_MARK_WIDTH = 18

    def paint(self, painter, option, index):
        node_type = index.data(ROLE_TYPE)
        if node_type == "notebook":
            shifted_option = QStyleOptionViewItem(option)
            shifted_option.rect = option.rect.adjusted(
                self._OPEN_NOTEBOOK_MARK_WIDTH,
                0,
                0,
                0,
            )
            super().paint(painter, shifted_option, index)

            if bool(index.data(ROLE_OPEN_NOTEBOOK)):
                painter.save()
                try:
                    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
                    pen = QPen(QColor("#B9FF5A"))
                    pen.setWidth(2)
                    pen.setCapStyle(Qt.PenCapStyle.RoundCap)
                    pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
                    painter.setPen(pen)

                    left = option.rect.left() + 3
                    center_y = option.rect.center().y()
                    painter.drawLine(left + 1, center_y, left + 6, center_y + 5)
                    painter.drawLine(left + 6, center_y + 5, left + 14, center_y - 6)
                finally:
                    painter.restore()
            return

        super().paint(painter, option, index)

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setObjectName("TreeItemRenameEditor")
        editor.setFrame(False)
        editor.setTextMargins(
            self._HORIZONTAL_TEXT_MARGIN,
            self._VERTICAL_TEXT_MARGIN,
            self._HORIZONTAL_TEXT_MARGIN,
            self._VERTICAL_TEXT_MARGIN,
        )

        fm = editor.fontMetrics()
        min_height = fm.height() + (self._VERTICAL_TEXT_MARGIN * 2) + self._EDITOR_EXTRA_HEIGHT
        editor.setMinimumHeight(min_height)
        editor.setStyleSheet(
            """
            QLineEdit#TreeItemRenameEditor {
                padding: 0px;
                border: 1px solid #0078D7;
                border-radius: 4px;
                background: palette(base);
            }
            """
        )
        return editor

    def updateEditorGeometry(self, editor, option, index):
        rect = option.rect.adjusted(0, -1, 0, 1)
        if index.data(ROLE_TYPE) == "notebook":
            rect = rect.adjusted(self._OPEN_NOTEBOOK_MARK_WIDTH, 0, 0, 0)
        editor_height = max(rect.height(), editor.minimumHeight())

        if editor_height > rect.height():
            extra = editor_height - rect.height()
            top_extra = extra // 2
            bottom_extra = extra - top_extra
            rect = rect.adjusted(0, -top_extra, 0, bottom_extra)

        editor.setGeometry(rect)

    def sizeHint(self, option, index):
        base = super().sizeHint(option, index)
        min_height = option.fontMetrics.height() + (self._VERTICAL_TEXT_MARGIN * 2) + self._EDITOR_EXTRA_HEIGHT
        return QSize(base.width(), max(base.height(), min_height))

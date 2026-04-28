# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


class MainWindowInitLeftMixin:

    def _ensure_tree_icons(self) -> None:
        if getattr(self, "_tree_icons_ready", False):
            return
        self._tree_icons_ready = True
        try:
            style = self.style()
            pixmap = QApplication.style().StandardPixmap
            self._icon_file = style.standardIcon(pixmap.SP_FileIcon)
            self._icon_dir = style.standardIcon(pixmap.SP_DirIcon)
            self._icon_agg = style.standardIcon(pixmap.SP_ComputerIcon)
            self._icon_open_notebook = _make_open_notebook_check_icon()
        except Exception:
            self._icon_file = None
            self._icon_dir = None
            self._icon_agg = None
            self._icon_open_notebook = None

    def _apply_loaded_tree_icons(self) -> None:
        self._ensure_tree_icons()
        for tree_name in ("buffer_tree", "fav_tree"):
            tree = getattr(self, tree_name, None)
            if tree is None:
                continue
            was_blocked = tree.blockSignals(True)
            was_updates_enabled = tree.updatesEnabled()
            tree.setUpdatesEnabled(False)
            try:
                root = tree.invisibleRootItem()
                stack = [root.child(i) for i in range(root.childCount() - 1, -1, -1)]
                while stack:
                    item = stack.pop()
                    node_type = item.data(0, ROLE_TYPE)
                    payload = item.data(0, ROLE_DATA) or {}
                    if tree_name == "buffer_tree":
                        if node_type == "group":
                            icon = self._icon_dir
                        elif payload.get("virtual") == "aggregate":
                            icon = self._icon_agg
                        else:
                            icon = self._icon_file
                    else:
                        icon = (
                            self._icon_file
                            if node_type in ("section", "notebook")
                            else self._icon_dir
                        )
                    if icon is not None:
                        item.setIcon(0, icon)
                    for j in range(item.childCount() - 1, -1, -1):
                        stack.append(item.child(j))
            finally:
                tree.setUpdatesEnabled(was_updates_enabled)
                tree.blockSignals(was_blocked)

    def _build_left_panels(self) -> None:
        def _make_toolbar_button_compact(button, *, fixed_width: Optional[int] = None) -> None:
            try:
                if fixed_width is not None:
                    button.setFixedWidth(fixed_width)
                    return
                if not IS_WINDOWS:
                    return
                button.setMinimumWidth(38)
                button.setSizePolicy(
                    QSizePolicy.Policy.Minimum,
                    QSizePolicy.Policy.Fixed,
                )
            except Exception:
                pass

        # 1. 즐겨찾기 버퍼 관리 패널 (가장 왼쪽)
        buffer_panel = QWidget()
        if IS_WINDOWS:
            buffer_panel.setMinimumWidth(80)
            buffer_panel.setSizePolicy(
                QSizePolicy.Policy.Ignored,
                QSizePolicy.Policy.Preferred,
            )
        buffer_layout = QVBoxLayout(buffer_panel)
        buffer_layout.setContentsMargins(0, 0, 0, 0)
        buffer_layout.setSpacing(8)

        buffer_group = QGroupBox(_buffer_group_title())
        if IS_WINDOWS:
            buffer_group.setMinimumWidth(0)
            buffer_group.setSizePolicy(
                QSizePolicy.Policy.Ignored,
                QSizePolicy.Policy.Preferred,
            )
        buffer_group_layout = QVBoxLayout(buffer_group)
        if IS_WINDOWS:
            buffer_group_layout.setContentsMargins(6, 8, 6, 6)
            buffer_group_layout.setSpacing(5)

        # 즐겨찾기 버퍼 상단 툴바: 추가, 이름변경
        buffer_toolbar_top_layout = QHBoxLayout()
        if IS_WINDOWS:
            buffer_toolbar_top_layout.setSpacing(4)
        self.btn_add_buffer_group = QToolButton()
        self.btn_add_buffer_group.setText(_buffer_group_add_label())
        if IS_WINDOWS:
            self.btn_add_buffer_group.setText("그룹")
        self.btn_add_buffer_group.clicked.connect(self._add_buffer_group)
        _make_toolbar_button_compact(self.btn_add_buffer_group)

        self.btn_add_buffer = QToolButton()
        self.btn_add_buffer.setText(_buffer_item_add_label())
        if IS_WINDOWS:
            self.btn_add_buffer.setText("버퍼")
        self.btn_add_buffer.clicked.connect(self._add_buffer)
        _make_toolbar_button_compact(self.btn_add_buffer)

        self.btn_rename_buffer = QToolButton()
        self.btn_rename_buffer.setText(_rename_button_label())
        if IS_WINDOWS:
            self.btn_rename_buffer.setText("이름")
        self.btn_rename_buffer.clicked.connect(self._rename_buffer)
        _make_toolbar_button_compact(self.btn_rename_buffer)

        self.btn_register_all_notebooks = QToolButton()
        self.btn_register_all_notebooks.setText(_register_all_notebooks_button_label())
        self.btn_register_all_notebooks.setToolTip(
            "현재 열린 OneNote 전자필기장 목록을 다시 읽고 연두색 열림 체크와 분류 상태를 한 번에 갱신합니다."
        )
        self.btn_register_all_notebooks.clicked.connect(self._register_all_notebooks_from_current_onenote)
        self.btn_register_all_notebooks.setEnabled(False)  # 종합 버퍼에서만 활성화
        self.btn_register_all_notebooks.setVisible(False)  # 종합 버퍼에서만 표시
        _make_toolbar_button_compact(self.btn_register_all_notebooks)

        buffer_toolbar_top_layout.addWidget(self.btn_add_buffer_group)
        buffer_toolbar_top_layout.addWidget(self.btn_add_buffer)
        if not IS_WINDOWS:
            buffer_toolbar_top_layout.addWidget(self.btn_rename_buffer)
        buffer_toolbar_top_layout.addStretch(1)
        buffer_group_layout.addLayout(buffer_toolbar_top_layout)

        # QListWidget -> BufferTree로 교체
        self.buffer_tree = BufferTree()
        if IS_WINDOWS:
            self.buffer_tree.setMinimumWidth(0)
            self.buffer_tree.setSizePolicy(
                QSizePolicy.Policy.Ignored,
                QSizePolicy.Policy.Expanding,
            )
        self._tree_name_edit_delegate = TreeNameEditDelegate(self)
        self.buffer_tree.setItemDelegate(self._tree_name_edit_delegate)
        # PERF: 큰 트리에서 초기 렌더링 성능 개선
        try:
            self.buffer_tree.setUniformRowHeights(True)
            self.buffer_tree.setAnimated(False)
        except Exception:
            pass

        self.buffer_tree.itemClicked.connect(self._on_buffer_tree_item_clicked)
        self.buffer_tree.itemDoubleClicked.connect(self._on_buffer_tree_double_clicked)
        # ✅ 2패널(모듈/섹션)이 즉시 갱신되도록 하되, 부팅 중에는 무시
        self.buffer_tree.itemSelectionChanged.connect(
            lambda: getattr(self, "_boot_loading", False) or self._on_buffer_tree_selection_changed()
        )
        self.buffer_tree.structureChanged.connect(self._request_buffer_structure_save)
        self.buffer_tree.renameRequested.connect(self._rename_buffer)
        self.buffer_tree.deleteRequested.connect(self._delete_buffer)

        self.buffer_tree.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self.buffer_tree.customContextMenuRequested.connect(
            self._on_buffer_context_menu
        )

        buffer_group_layout.addWidget(self.buffer_tree)

        # 즐겨찾기 버퍼 하단 툴바: 삭제, 위로, 아래로
        self.btn_delete_buffer = QToolButton()
        self.btn_delete_buffer.setText("삭제")
        _make_toolbar_button_compact(
            self.btn_delete_buffer,
            fixed_width=44 if IS_WINDOWS else 52,
        )
        self.btn_delete_buffer.clicked.connect(self._delete_buffer)

        self.btn_buffer_move_up = QToolButton()
        self.btn_buffer_move_up.setText("▲")
        self.btn_buffer_move_up.setToolTip("위로")
        _make_toolbar_button_compact(
            self.btn_buffer_move_up,
            fixed_width=26 if IS_WINDOWS else 32,
        )
        self.btn_buffer_move_up.clicked.connect(self._move_buffer_up)

        self.btn_buffer_move_down = QToolButton()
        self.btn_buffer_move_down.setText("▼")
        self.btn_buffer_move_down.setToolTip("아래로")
        _make_toolbar_button_compact(
            self.btn_buffer_move_down,
            fixed_width=26 if IS_WINDOWS else 32,
        )
        self.btn_buffer_move_down.clicked.connect(self._move_buffer_down)

        buffer_toolbar_bottom_layout = QVBoxLayout() if IS_WINDOWS else QHBoxLayout()
        if IS_WINDOWS:
            buffer_toolbar_bottom_layout.setSpacing(3)
            buffer_edit_layout = QHBoxLayout()
            buffer_edit_layout.setSpacing(4)
            buffer_edit_layout.addWidget(self.btn_delete_buffer)
            buffer_edit_layout.addWidget(self.btn_rename_buffer)
            buffer_edit_layout.addStretch(1)
            buffer_move_layout = QHBoxLayout()
            buffer_move_layout.setSpacing(4)
            buffer_move_layout.addStretch(1)
            buffer_move_layout.addWidget(self.btn_buffer_move_up)
            buffer_move_layout.addWidget(self.btn_buffer_move_down)
            buffer_toolbar_bottom_layout.addLayout(buffer_edit_layout)
            buffer_toolbar_bottom_layout.addLayout(buffer_move_layout)
        else:
            buffer_toolbar_bottom_layout.addWidget(self.btn_delete_buffer)
            buffer_toolbar_bottom_layout.addStretch(1)
            buffer_toolbar_bottom_layout.addWidget(self.btn_buffer_move_up)
            buffer_toolbar_bottom_layout.addWidget(self.btn_buffer_move_down)
        buffer_group_layout.addLayout(buffer_toolbar_bottom_layout)

        buffer_layout.addWidget(buffer_group)
        self.left_splitter.addWidget(buffer_panel)

        # 2. 즐겨찾기 관리 패널 (중앙)
        favorites_panel = QWidget()
        if IS_WINDOWS:
            favorites_panel.setMinimumWidth(104)
            favorites_panel.setSizePolicy(
                QSizePolicy.Policy.Ignored,
                QSizePolicy.Policy.Preferred,
            )
        left_layout = QVBoxLayout(favorites_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        fav_group = QGroupBox(_favorites_group_title())
        if IS_WINDOWS:
            fav_group.setMinimumWidth(0)
            fav_group.setSizePolicy(
                QSizePolicy.Policy.Ignored,
                QSizePolicy.Policy.Preferred,
            )
        fav_layout = QVBoxLayout(fav_group)
        if IS_WINDOWS:
            fav_layout.setContentsMargins(6, 8, 6, 6)
            fav_layout.setSpacing(5)

        # 툴바 - 1행: 그룹추가, 현재 전자필기장 추가, 이름 바꾸기
        tb1_layout = QHBoxLayout()
        if IS_WINDOWS:
            tb1_layout.setSpacing(4)
        self.btn_add_group = QToolButton()
        self.btn_add_group.setText("그룹 추가")
        if IS_WINDOWS:
            self.btn_add_group.setText("그룹")
        self.btn_add_group.clicked.connect(self._add_group)
        _make_toolbar_button_compact(self.btn_add_group)
        self.btn_add_section_current = QToolButton()
        self.btn_add_section_current.setText(_current_add_button_label())
        if IS_WINDOWS:
            self.btn_add_section_current.setText("현재")
        self.btn_add_section_current.clicked.connect(self._add_section_from_current)
        _make_toolbar_button_compact(self.btn_add_section_current)
        self.btn_activate_favorite = QToolButton()
        self.btn_activate_favorite.setText(_favorite_activate_button_label())
        if IS_WINDOWS:
            self.btn_activate_favorite.setText("실행")
        self.btn_activate_favorite.clicked.connect(self._activate_current_favorite_item)
        self.btn_activate_favorite.setEnabled(False)
        _make_toolbar_button_compact(self.btn_activate_favorite)
        self.btn_rename = QToolButton()
        self.btn_rename.setText(_rename_button_label())
        if IS_WINDOWS:
            self.btn_rename.setText("이름")
        self.btn_rename.clicked.connect(self._rename_favorite_item)
        _make_toolbar_button_compact(self.btn_rename)
        tb1_layout.addWidget(self.btn_add_section_current)
        tb1_layout.addWidget(self.btn_activate_favorite)
        if not IS_WINDOWS:
            tb1_layout.addWidget(self.btn_rename)
        tb1_layout.addStretch(1)

        # 툴바 - 2행: 그룹 추가, 그룹 펼치기/접기 드롭다운
        tb2_layout = QHBoxLayout()
        if IS_WINDOWS:
            tb2_layout.setSpacing(4)
        self.btn_group_expand_collapse = QToolButton()
        self.btn_group_expand_collapse.setText("그룹 펼치기/접기")
        if IS_WINDOWS:
            self.btn_group_expand_collapse.setText("펼침/접기")
        self.btn_group_expand_collapse.setToolTip("그룹 펼치기 또는 그룹 접기를 선택합니다.")
        self.btn_group_expand_collapse.setPopupMode(
            QToolButton.ToolButtonPopupMode.InstantPopup
        )
        self.menu_group_expand_collapse = QMenu(self.btn_group_expand_collapse)
        self.action_expand_all_groups = QAction("그룹 펼치기", self)
        self.action_collapse_all_groups = QAction("그룹 접기", self)
        self.menu_group_expand_collapse.addAction(self.action_expand_all_groups)
        self.menu_group_expand_collapse.addAction(self.action_collapse_all_groups)
        self.btn_group_expand_collapse.setMenu(self.menu_group_expand_collapse)
        _make_toolbar_button_compact(self.btn_group_expand_collapse)
        tb2_layout.addWidget(self.btn_add_group)
        tb2_layout.addStretch(1)
        tb2_layout.addWidget(self.btn_group_expand_collapse)

        tb3_layout = QHBoxLayout()
        if IS_WINDOWS:
            tb3_layout.setSpacing(4)
        tb3_layout.addStretch(1)
        tb3_layout.addWidget(self.btn_register_all_notebooks)
        tb3_layout.addStretch(1)

        fav_layout.addLayout(tb1_layout)
        fav_layout.addLayout(tb2_layout)
        fav_layout.addLayout(tb3_layout)

        self.fav_tree = FavoritesTree()
        if IS_WINDOWS:
            self.fav_tree.setMinimumWidth(0)
            self.fav_tree.setSizePolicy(
                QSizePolicy.Policy.Ignored,
                QSizePolicy.Policy.Expanding,
            )
        self.fav_tree.setItemDelegate(self._tree_name_edit_delegate)
        # PERF: 큰 트리에서 초기 렌더링 성능 개선
        try:
            self.fav_tree.setUniformRowHeights(True)
            self.fav_tree.setAnimated(False)
        except Exception:
            pass

        self._tree_icons_ready = False
        self._icon_file = None
        self._icon_dir = None
        self._icon_agg = None
        self._icon_open_notebook = None

        self.action_expand_all_groups.triggered.connect(
            lambda: self._expand_fav_groups_always(reason="toolbar")
        )
        self.action_collapse_all_groups.triggered.connect(
            lambda: self._collapse_fav_groups_always(reason="toolbar")
        )
        self.fav_tree.itemDoubleClicked.connect(self._on_fav_item_double_clicked)
        self.fav_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.fav_tree.customContextMenuRequested.connect(self._on_fav_context_menu)
        self.fav_tree.structureChanged.connect(self._request_favorites_save)
        self.fav_tree.itemChanged.connect(self._request_favorites_save)
        # ✅ 키보드 매핑 (Del, F2, Ctrl+C/V) 직접 연결
        self.fav_tree.deleteRequested.connect(self._delete_favorite_item)
        self.fav_tree.renameRequested.connect(self._rename_favorite_item)
        self.fav_tree.copyRequested.connect(self._copy_favorite_item)
        self.fav_tree.pasteRequested.connect(self._paste_favorite_item)
        self.fav_tree.cutRequested.connect(self._cut_favorite_item)
        self.fav_tree.undoRequested.connect(self._undo_favorite_tree)
        self.fav_tree.redoRequested.connect(self._redo_favorite_tree)
        fav_layout.addWidget(self.fav_tree)

        # 즐겨찾기 하단 툴바: 삭제, 위로, 아래로 (삭제 버튼 재배치)
        move_buttons_layout = QHBoxLayout()
        if IS_WINDOWS:
            move_buttons_layout.setSpacing(4)

        # 삭제 버튼 (tb2에서 이동)
        self.btn_delete = QToolButton()
        self.btn_delete.setText("삭제")
        _make_toolbar_button_compact(
            self.btn_delete,
            fixed_width=44 if IS_WINDOWS else 52,
        )
        self.btn_delete.clicked.connect(self._delete_favorite_item)

        self.btn_move_up = QToolButton()
        self.btn_move_up.setText("▲")
        self.btn_move_up.setToolTip("위로")
        _make_toolbar_button_compact(
            self.btn_move_up,
            fixed_width=26 if IS_WINDOWS else 32,
        )
        self.btn_move_up.clicked.connect(self._move_item_up)
        self.btn_move_down = QToolButton()
        self.btn_move_down.setText("▼")
        self.btn_move_down.setToolTip("아래로")
        _make_toolbar_button_compact(
            self.btn_move_down,
            fixed_width=26 if IS_WINDOWS else 32,
        )
        self.btn_move_down.clicked.connect(self._move_item_down)

        if IS_WINDOWS:
            favorite_edit_layout = QHBoxLayout()
            favorite_edit_layout.setSpacing(4)
            favorite_edit_layout.addWidget(self.btn_delete)
            favorite_edit_layout.addWidget(self.btn_rename)
            favorite_edit_layout.addStretch(1)
            move_buttons_layout.addStretch(1)
            move_buttons_layout.addWidget(self.btn_move_up)
            move_buttons_layout.addWidget(self.btn_move_down)
            fav_layout.addLayout(favorite_edit_layout)
            fav_layout.addLayout(move_buttons_layout)
        else:
            move_buttons_layout.addWidget(self.btn_delete)
            move_buttons_layout.addStretch(1)
            move_buttons_layout.addWidget(self.btn_move_up)
            move_buttons_layout.addWidget(self.btn_move_down)
            fav_layout.addLayout(move_buttons_layout)

        self.fav_tree.itemSelectionChanged.connect(self._update_move_button_state)
        self.fav_tree.itemSelectionChanged.connect(self._sync_favorite_action_buttons)
        self.fav_tree.itemSelectionChanged.connect(
            self._sync_codex_target_from_current_fav_item
        )
        left_layout.addWidget(fav_group, stretch=1)

        self.left_splitter.addWidget(favorites_panel)
        self.main_splitter.addWidget(self.left_splitter)

_publish_context(globals())

# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin39:

    def _delete_buffer(self):
        print("[DBG][BUF][DEL] _delete_buffer: ENTER")
        try:
            item = self.buffer_tree.currentItem()
            print(f"[DBG][BUF][DEL] currentItem={item}")
            if not item:
                print("[DBG][BUF][DEL] no currentItem -> RETURN")
                return

            node_type = item.data(0, ROLE_TYPE)
            payload = item.data(0, ROLE_DATA) or {}
            name = item.text(0) or "(no-name)"
            deleting_id = payload.get("id")
            locked = bool(payload.get("locked"))

            parent = item.parent() or self.buffer_tree.invisibleRootItem()
            idx = parent.indexOfChild(item)
            print(f"[DBG][BUF][DEL] node_type={node_type} name='{name}' id={deleting_id} locked={locked} parent={parent} idx={idx}")

            if locked:
                print("[DBG][BUF][DEL] locked item -> blocked")
                QMessageBox.information(self, "삭제 불가", "이 항목은 고정 항목이라 삭제할 수 없습니다.")
                return

            deleting_active = bool(self.active_buffer_id and deleting_id == self.active_buffer_id)
            print(f"[DBG][BUF][DEL] deleting_active={deleting_active} active_buffer_id={self.active_buffer_id}")

            # ✅ 확인창
            if node_type == "group":
                child_cnt = item.childCount()
                msg = f"그룹 '{name}' 을(를) 삭제할까요?\n\n하위 항목 {child_cnt}개도 함께 삭제됩니다."
            else:
                msg = f"버퍼 '{name}' 을(를) 삭제할까요?"

            reply = QMessageBox.question(
                self,
                "삭제 확인",
                msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            print(f"[DBG][BUF][DEL] confirm reply={reply}")
            if reply != QMessageBox.StandardButton.Yes:
                print("[DBG][BUF][DEL] user cancelled -> RETURN")
                return

            # ✅ 활성 버퍼 삭제면 중앙 트리 저장(유실 방지)
            if deleting_active:
                print("[DBG][BUF][DEL] deleting active buffer -> _save_favorites()")
                try:
                    self._save_favorites()
                except Exception:
                    print("[ERR][BUF][DEL] _save_favorites failed (ignored)")
                    traceback.print_exc()

            # ✅ 실제 트리에서 제거
            taken = parent.takeChild(idx)
            print(f"[DBG][BUF][DEL] takeChild result={taken}")
            del taken

            # ✅ 구조 저장
            print("[DBG][BUF][DEL] _save_buffer_structure()")
            try:
                self._save_buffer_structure()
            except Exception:
                print("[ERR][BUF][DEL] _save_buffer_structure failed")
                traceback.print_exc()

            # ✅ 활성 버퍼였다면 다른 버퍼로 자동 전환
            if deleting_active:
                print("[DBG][BUF][DEL] deleted active -> reset active and auto-select next buffer")
                self.active_buffer_id = None
                self.active_buffer_item = None
                self.active_buffer_node = None
                self._active_buffer_settings_node = None
                self.settings["active_buffer_id"] = None

                found_item = self._first_buffer_item
                print(f"[DBG][BUF][DEL] next buffer found_item={found_item}")
                if found_item:
                    self.buffer_tree.setCurrentItem(found_item)
                    self._on_buffer_tree_item_clicked(found_item, 0)
                else:
                    print("[DBG][BUF][DEL] no buffer remains -> clear fav_tree")
                    try:
                        self.fav_tree.clear()
                    except Exception:
                        traceback.print_exc()

            try:
                self._update_buffer_move_button_state()
            except Exception:
                pass

            print("[DBG][BUF][DEL] DONE")
        except Exception:
            print("[ERR][BUF][DEL] exception in _delete_buffer")
            traceback.print_exc()

    # ----------------- 15-2. 즐겨찾기 버퍼 순서 변경 로직 (수정) -----------------
    def _update_buffer_move_button_state(self):
        """버퍼 트리 이동 버튼 상태 업데이트"""
        item = self.buffer_tree.currentItem()
        if not item:
            self.btn_buffer_move_up.setEnabled(False)
            self.btn_buffer_move_down.setEnabled(False)
            return

        parent = item.parent() or self.buffer_tree.invisibleRootItem()
        index = parent.indexOfChild(item)

        self.btn_buffer_move_up.setEnabled(index > 0)
        self.btn_buffer_move_down.setEnabled(index < parent.childCount() - 1)

    def _move_buffer_up(self):
        item = self.buffer_tree.currentItem()
        if not item: return

        parent = item.parent() or self.buffer_tree.invisibleRootItem()
        index = parent.indexOfChild(item)
        if index > 0:
            taken = parent.takeChild(index)
            parent.insertChild(index - 1, taken)
            self.buffer_tree.setCurrentItem(taken)
            self._save_buffer_structure()
            self._update_buffer_move_button_state()

    def _move_buffer_down(self):
        item = self.buffer_tree.currentItem()
        if not item: return

        parent = item.parent() or self.buffer_tree.invisibleRootItem()
        index = parent.indexOfChild(item)
        if index < parent.childCount() - 1:
            taken = parent.takeChild(index)
            parent.insertChild(index + 1, taken)
            self.buffer_tree.setCurrentItem(taken)
            self._save_buffer_structure()
            self._update_buffer_move_button_state()

    def _on_buffer_context_menu(self, pos):
        """버퍼 트리 컨텍스트 메뉴"""
        item = self.buffer_tree.currentItem()
        menu = QMenu(self)

        act_add_grp = QAction("그룹 추가", self)
        act_add_grp.triggered.connect(self._add_buffer_group)
        menu.addAction(act_add_grp)

        act_add_buf = QAction("버퍼 추가", self)
        act_add_buf.triggered.connect(self._add_buffer)
        menu.addAction(act_add_buf)

        if item:
            menu.addSeparator()
            act_rename = QAction("이름 변경 (F2)", self)
            act_rename.triggered.connect(self._rename_buffer)
            menu.addAction(act_rename)

            act_del = QAction("삭제", self)
            act_del.triggered.connect(self._delete_buffer)
            menu.addAction(act_del)

        menu.exec(self.buffer_tree.viewport().mapToGlobal(pos))


    # ----------------- 16-1. 즐겨찾기 복사/붙여넣기 로직 -----------------
    def _copy_favorite_item(self):
        """선택된 즐겨찾기 항목(다중 가능)을 복사합니다."""
        items = self._selected_fav_items_top()
        if not items:
            return
        payload = [self._serialize_fav_item(it) for it in items]
        # 단일이면 dict로, 다중이면 list로 저장
        self.clipboard_data = payload[0] if len(payload) == 1 else payload
        self.connection_status_label.setText(
            f"{len(items)}개 항목 복사됨."
        )

    def _paste_favorite_item(self):
        """클립보드에 있는 즐겨찾기 항목을 붙여넣습니다."""
        if not self.clipboard_data:
            QMessageBox.warning(
                self, "붙여넣기 오류", "클립보드에 복사된 항목이 없습니다."
            )
            return

        # ✅ 붙여넣기 대상 정규화: notebook/section 선택 상태에서 붙여넣으면
        #    항목(노트북/섹션) 안에 항목이 들어가 트리가 꼬입니다.
        #    따라서 선택이 notebook/section이면 자동으로 group 레벨로 올려 붙여넣습니다.
        parent = self._normalize_fav_paste_parent(self._current_fav_item())

        def _deep_copy_node(node: Dict[str, Any]) -> Dict[str, Any]:
            new_node = node.copy()
            new_node["id"] = str(uuid.uuid4())
            # new_node["name"] = f"복사본 - {new_node['name']}" # 이 줄을 제거하거나 주석 처리
            if "children" in new_node:
                new_node["children"] = [
                    _deep_copy_node(child) for child in new_node["children"]
                ]
            return new_node

        try:
            nodes = self.clipboard_data
            if isinstance(nodes, dict):
                nodes = [nodes]
            if not isinstance(nodes, list):
                nodes = []
            new_items = []

            # ✅ 다중 붙여넣기를 '한 번의 Undo'로 묶기
            with self._fav_bulk_edit(reason=f"paste:{len(nodes)}"):
                for node in nodes:
                    copied_node = _deep_copy_node(node)
                    new_item = self._append_fav_node(parent, copied_node)
                    new_items.append(new_item)
                if new_items:
                    self.fav_tree.setCurrentItem(new_items[-1])

            self.connection_status_label.setText(f"{len(new_items)}개 항목 붙여넣기 완료.")

        except Exception as e:
            QMessageBox.critical(
                self, "붙여넣기 오류", f"항목을 붙여넣는 중 오류가 발생했습니다: {e}"
            )

    def _selected_fav_items_top(self) -> List[QTreeWidgetItem]:
        """
        다중 선택 시 '상위 선택만' 반환합니다.
        (부모와 자식을 동시에 선택했을 때 중복 복사/붙여넣기 방지)
        """
        items = self.fav_tree.selectedItems()
        if not items:
            return []
        # QTreeWidgetItem은 unhashable이므로 id()로 membership set 구성
        selected_ids = {id(it) for it in items}
        top_items: List[QTreeWidgetItem] = []
        for it in items:
            p = it.parent()
            skip = False
            while p is not None:
                if id(p) in selected_ids:
                    skip = True
                    break
                p = p.parent()
            if not skip:
                top_items.append(it)
        return top_items

    def _normalize_fav_paste_parent(self, item: Optional[QTreeWidgetItem]) -> QTreeWidgetItem:
        """
        붙여넣기 대상 정규화:
        - group 선택: 그대로 group 안에 붙여넣기
        - notebook/section 선택: 자동으로 상위 group 레벨에 붙여넣기 (항목-항목 중첩 방지)
        - 그 외/None: 루트
        """
        root = self.fav_tree.invisibleRootItem()
        if not item:
            return root
        try:
            t = item.data(0, ROLE_TYPE)
        except Exception:
            return root
        if t == "group":
            return item
        # notebook/section 등은 group까지 올라간다
        p = item
        while p is not None:
            try:
                if p.data(0, ROLE_TYPE) == "group":
                    return p
            except Exception:
                break
            p = p.parent()
        return root

    def _fav_capture_center_tree_snapshot(self) -> str:
        """
        현재 중앙 FavoritesTree의 상태를 JSON 스냅샷으로 캡쳐한다.
        Undo 그룹(base/final)을 _fav_last_snapshot에 의존하면
        버퍼 전환/리빌드 스킵/hash 최적화/중간 save 타이밍에 의해 base==final로 잡혀
        Ctrl+Z가 '안 먹는' 상태가 생길 수 있어서, 트리에서 직접 캡쳐로 고정한다.
        """
        try:
            data = []
            root = self.fav_tree.invisibleRootItem()
            for i in range(root.childCount()):
                data.append(self._serialize_fav_item(root.child(i)))
            return json.dumps(data, sort_keys=True, ensure_ascii=False)
        except Exception:
            return "[]"

    def _fav_record_snapshot(self, new_snapshot: str):
        """FavoritesTree 변경사항을 Undo/Redo 스택에 기록합니다."""
        # (A) 로드/undo apply 중에는 기록하지 않는다.
        if getattr(self, "_fav_undo_suspended", False):
            self._fav_last_snapshot = new_snapshot
            return

        # (B) 최초 스냅샷
        if self._fav_last_snapshot is None:
            self._fav_last_snapshot = new_snapshot
            return

        # (C) 동일하면 skip
        if new_snapshot == self._fav_last_snapshot:
            return

        # (D) 다중 붙여넣기/다중 삭제 같은 "벌크 변경"에서는
        #     itemChanged가 여러 번 발생하며 _save_favorites()가 연속 호출된다.
        #     이때 매번 undo 스택에 쌓이면 Ctrl+Z가 "한 개씩" 되돌아가서 답답해진다.
        #     => 트랜잭션(depth>0)에서는 _fav_last_snapshot만 갱신하고,
        #        최종 커밋은 _fav_end_undo_group()에서 1회만 수행한다.
        if getattr(self, "_fav_undo_batch_depth", 0) > 0:
            self._fav_last_snapshot = new_snapshot
            self._fav_undo_batch_final_snapshot = new_snapshot
            return

        # (E) 일반 단일 변경
        self._fav_undo_stack.append(self._fav_last_snapshot)
        if len(self._fav_undo_stack) > self._fav_undo_max:
            self._fav_undo_stack = self._fav_undo_stack[-self._fav_undo_max:]
        self._fav_redo_stack.clear()
        self._fav_last_snapshot = new_snapshot

    def _fav_begin_undo_group(self, *, reason: str = "") -> None:
        """여러 변경을 한 번의 Undo step으로 묶기 시작."""
        if self._fav_undo_batch_depth == 0:
            # ✅ base는 _fav_last_snapshot 대신 '현재 트리'에서 직접 캡쳐 (base==final 문제 원천 차단)
            base = self._fav_capture_center_tree_snapshot()
            self._fav_undo_batch_base_snapshot = base
            self._fav_undo_batch_final_snapshot = None
            self._fav_undo_batch_reason = reason or ""
            # base를 last_snapshot에도 맞춰둬야 이후 비교가 흔들리지 않는다.
            self._fav_last_snapshot = base
        self._fav_undo_batch_depth += 1
        try:
            if reason:
                print(f"[DBG][FAV][UNDO_GRP] begin depth={self._fav_undo_batch_depth} reason={reason} base_len={len(self._fav_undo_batch_base_snapshot or '')}")
        except Exception:
            pass

_publish_context(globals())

# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin40:

    def _fav_end_undo_group(self) -> None:
        """Undo group 종료: 변경이 있었으면 undo 스택에 1회만 커밋."""
        if self._fav_undo_batch_depth <= 0:
            return
        self._fav_undo_batch_depth -= 1
        if self._fav_undo_batch_depth != 0:
            return

        base = self._fav_undo_batch_base_snapshot or ""
        # ✅ final도 last_snapshot 의존 X: 트리에서 직접 캡쳐
        final = self._fav_undo_batch_final_snapshot
        if final is None:
            final = self._fav_capture_center_tree_snapshot()
        # last_snapshot을 final로 동기화 (Undo/Redo 비교 흔들림 방지)
        self._fav_last_snapshot = final

        changed = (final != base)
        if changed:
            self._fav_undo_stack.append(base)
            if len(self._fav_undo_stack) > self._fav_undo_max:
                self._fav_undo_stack = self._fav_undo_stack[-self._fav_undo_max:]
            self._fav_redo_stack.clear()

        try:
            r = self._fav_undo_batch_reason
            print(
                f"[DBG][FAV][UNDO_GRP] end changed={int(changed)} undo={len(self._fav_undo_stack)} redo={len(self._fav_redo_stack)} reason={r} base_len={len(base)} final_len={len(final)}"
            )
        except Exception:
            pass

        self._fav_undo_batch_base_snapshot = None
        self._fav_undo_batch_final_snapshot = None
        self._fav_undo_batch_reason = ""

    @contextmanager
    def _fav_bulk_edit(self, *, reason: str = ""):
        """
        FavoritesTree를 벌크로 수정할 때 사용.
        - Qt itemChanged 연쇄 save를 막기 위해 fav_tree signals를 잠깐 막고
        - Undo/Redo는 begin/end로 한 번의 step으로 묶는다.
        """
        self._fav_begin_undo_group(reason=reason)
        was_updates_enabled = self.fav_tree.updatesEnabled()
        self.fav_tree.blockSignals(True)
        self.fav_tree.setUpdatesEnabled(False)
        try:
            yield
            # 벌크 작업이 끝나면 딱 1번만 저장(=스냅샷 갱신)
            try:
                self._save_favorites()
            except Exception:
                pass
        finally:
            self.fav_tree.setUpdatesEnabled(was_updates_enabled)
            self.fav_tree.blockSignals(False)
            try:
                if was_updates_enabled:
                    self.fav_tree.viewport().update()
            except Exception:
                pass
            self._fav_end_undo_group()

    def _fav_apply_snapshot(self, snapshot: str):
        """스냅샷(JSON 문자열)을 중앙 즐겨찾기 트리에 적용합니다."""
        try:
            data = json.loads(snapshot) if snapshot else []
            if not isinstance(data, list):
                data = []
        except Exception:
            print("[ERR][FAV][UNDO] invalid snapshot")
            traceback.print_exc()
            return
        self._fav_undo_suspended = True
        try:
            # ✅ Undo/Redo는 무조건 리빌드되어야 한다.
            # 해시 스킵 최적화가 undo 적용을 막는 케이스를 원천 차단.
            self._last_center_payload_hash = None
            self._last_center_payload_snapshot = None
            self._last_center_payload_source_id = 0
            self._fav_last_persisted_hash = None
            self._load_favorites_into_center_tree(data)
            self._save_favorites()
        finally:
            self._fav_undo_suspended = False

    def _undo_favorite_tree(self):
        try:
            print(f"[DBG][FAV][UNDO] called undo={len(self._fav_undo_stack)} redo={len(self._fav_redo_stack)} last_len={len(self._fav_last_snapshot or '')}")
        except Exception:
            pass
        if not self._fav_undo_stack:
            try:
                self.connection_status_label.setText("되돌릴 작업이 없습니다.")
            except Exception:
                pass
            return
        cur = self._fav_last_snapshot
        if cur is None:
            try:
                cur = json.dumps(self.active_buffer_node.get("data", []), sort_keys=True, ensure_ascii=False)
            except Exception:
                cur = ""
        self._fav_redo_stack.append(cur or "")
        snap = self._fav_undo_stack.pop()
        self._fav_apply_snapshot(snap)
        try:
            self.connection_status_label.setText("되돌리기 완료 (Ctrl+Z)")
        except Exception:
            pass

    def _redo_favorite_tree(self):
        try:
            print(f"[DBG][FAV][REDO] called undo={len(self._fav_undo_stack)} redo={len(self._fav_redo_stack)} last_len={len(self._fav_last_snapshot or '')}")
        except Exception:
            pass
        if not self._fav_redo_stack:
            try:
                self.connection_status_label.setText("다시 실행할 작업이 없습니다.")
            except Exception:
                pass
            return
        cur = self._fav_last_snapshot
        if cur is None:
            try:
                cur = json.dumps(self.active_buffer_node.get("data", []), sort_keys=True, ensure_ascii=False)
            except Exception:
                cur = ""
        self._fav_undo_stack.append(cur or "")
        snap = self._fav_redo_stack.pop()
        self._fav_apply_snapshot(snap)
        try:
            self.connection_status_label.setText("다시 실행 완료 (Ctrl+Shift+Z)")
        except Exception:
            pass

    def _cut_favorite_item(self):
        """선택 항목 잘라내기(Ctrl+X): 복사 + 삭제."""
        items = self._selected_fav_items_top()
        if not items:
            cur = self._current_fav_item()
            if cur:
                items = [cur]
        if not items:
            return
        payload_nodes = [self._serialize_fav_item(it) for it in items]
        self.clipboard_data = payload_nodes[0] if len(payload_nodes) == 1 else payload_nodes

        # ✅ 다중 잘라내기를 '한 번의 Undo'로 묶기
        with self._fav_bulk_edit(reason=f"cut:{len(items)}"):
            # 실제 삭제 (부모 기준으로 takeChild)
            for it in items:
                parent = it.parent() or self.fav_tree.invisibleRootItem()
                idx = parent.indexOfChild(it)
                if idx >= 0:
                    parent.takeChild(idx)
        try:
            self.connection_status_label.setText(f"{len(items)}개 항목 잘라내기 완료.")
        except Exception:
            pass

    # ----------------- 16. 즐겨찾기 조작 -----------------
    def _current_fav_item(self) -> Optional[QTreeWidgetItem]:
        items = self.fav_tree.selectedItems()
        return items[0] if items else None

    def _sync_favorite_action_buttons(self):
        item = self._current_fav_item()
        node_type = item.data(0, ROLE_TYPE) if item is not None else None
        try:
            self.btn_activate_favorite.setEnabled(node_type in ("section", "notebook"))
        except Exception:
            pass
        try:
            self.btn_rename.setEnabled(item is not None)
        except Exception:
            pass

    def _activate_current_favorite_item(self):
        item = self._current_fav_item()
        if item is None:
            return
        node_type = item.data(0, ROLE_TYPE)
        if node_type not in ("section", "notebook"):
            return
        started_at = time.perf_counter()
        self._sync_codex_target_from_fav_item(item)
        self._activate_favorite_section(item, started_at=started_at)

    def _move_item_up(self):
        item = self._current_fav_item()
        if not item:
            return

        parent = item.parent() or self.fav_tree.invisibleRootItem()
        index = parent.indexOfChild(item)

        if index > 0:
            is_expanded = item.isExpanded()
            taken_item = parent.takeChild(index)
            parent.insertChild(index - 1, taken_item)
            taken_item.setExpanded(is_expanded)
            self.fav_tree.setCurrentItem(taken_item)
            self._save_favorites()
            self._update_move_button_state()

    def _move_item_down(self):
        item = self._current_fav_item()
        if not item:
            return

        parent = item.parent() or self.fav_tree.invisibleRootItem()
        index = parent.indexOfChild(item)

        if index < parent.childCount() - 1:
            is_expanded = item.isExpanded()
            taken_item = parent.takeChild(index)
            parent.insertChild(index + 1, taken_item)
            taken_item.setExpanded(is_expanded)
            self.fav_tree.setCurrentItem(taken_item)
            self._save_favorites()
            self._update_move_button_state()

    def _update_move_button_state(self):
        item = self._current_fav_item()

        if not item:
            self.btn_move_up.setEnabled(False)
            self.btn_move_down.setEnabled(False)
            return

        parent = item.parent() or self.fav_tree.invisibleRootItem()
        index = parent.indexOfChild(item)

        self.btn_move_up.setEnabled(index > 0)
        self.btn_move_down.setEnabled(index < parent.childCount() - 1)

    def _add_group(self):
        parent = self._current_fav_item()
        if parent and parent.data(0, ROLE_TYPE) == "section":
            parent = parent.parent()
        parent = parent or self.fav_tree.invisibleRootItem()
        node = {"type": "group", "name": "새 그룹", "children": []}
        item = self._append_fav_node(parent, node)
        self.fav_tree.editItem(item, 0)
        self._save_favorites()

_publish_context(globals())

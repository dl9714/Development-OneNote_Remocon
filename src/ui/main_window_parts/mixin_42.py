# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin42:

    def _add_section_from_current(self):
        if not self.onenote_window:
            QMessageBox.information(self, "안내", "먼저 연결된 창이 있어야 합니다.")
            return

        title = ""
        notebook_text = ""
        page_text = ""
        try:
            title = self.onenote_window.window_text()
        except Exception:
            pass

        section_text = None
        if IS_MACOS:
            context = self._mac_selected_outline_context(self.onenote_window)
            notebook_text = str(context.get("notebook") or "").strip()
            section_text = str(context.get("section") or "").strip() or None
            page_text = str(context.get("page") or "").strip()
        else:
            try:
                tc = self.tree_control or _find_tree_or_list(self.onenote_window)
                if tc:
                    sel = get_selected_tree_item_fast(tc)
                    if sel:
                        section_text = sel.window_text()
            except Exception:
                pass

        if IS_MACOS and section_text and page_text:
            default_name = f"{section_text} · {page_text}"
        else:
            default_name = section_text or page_text or notebook_text or title or (
                "새 보기" if IS_MACOS else "새 섹션"
            )
        name, ok = QInputDialog.getText(
            self,
            "보기 바로가기 추가" if IS_MACOS else "섹션 즐겨찾기 추가",
            "표시 이름:",
            text=default_name,
        )
        if not ok or not name.strip():
            return

        try:
            sig = build_window_signature(self.onenote_window)
        except Exception:
            sig = {}

        target = {
            "sig": sig,
            "notebook_text": notebook_text,
            "section_text": section_text,
            "page_text": page_text,
        }
        node = {"type": "section", "name": name.strip(), "target": target}

        parent = self._current_fav_item()
        if parent and parent.data(0, ROLE_TYPE) == "section":
            parent = parent.parent()
        parent = parent or self.fav_tree.invisibleRootItem()
        self._append_fav_node(parent, node)
        self._save_favorites()

    def _add_section_from_other_window(self):
        dialog = OtherWindowSelectionDialog(self.my_pid, self)
        if not dialog.exec():
            return
        info = dialog.selected_info
        if not info:
            return

        default_name = (info.get("title") or "새 섹션").strip() or "새 섹션"
        name, ok = QInputDialog.getText(
            self,
            "보기 바로가기 추가" if IS_MACOS else "섹션 즐겨찾기 추가",
            "표시 이름:",
            text=default_name,
        )
        if not ok or not name.strip():
            return

        try:
            ensure_pywinauto()
            win = resolve_window_target(info)
            if win is None:
                raise ElementNotFoundError
            sig = build_window_signature(win)
        except Exception:
            sig = {
                "handle": info.get("handle"),
                "pid": info.get("pid"),
                "class_name": info.get("class_name"),
                "title": info.get("title"),
            }
        target = {"sig": sig, "section_text": None}
        node = {"type": "section", "name": name.strip(), "target": target}

        parent = self._current_fav_item()
        if parent and parent.data(0, ROLE_TYPE) == "section":
            parent = parent.parent()
        parent = parent or self.fav_tree.invisibleRootItem()
        self._append_fav_node(parent, node)
        self._save_favorites()

    def _rename_favorite_item(self):
        item = self._current_fav_item()
        if not item:
            return
        self.fav_tree.editItem(item, 0)

    def _delete_favorite_item(self):
        print("[DBG][FAV][DEL] _delete_favorite_item: ENTER")
        try:
            # ✅ 다중선택 삭제: 상위 선택만 남김(부모/자식 중복 선택 방지)
            targets = self._selected_fav_items_top()
            if not targets:
                item = self._current_fav_item()
                if item:
                    targets = [item]
            print(f"[DBG][FAV][DEL] targets_count={len(targets)}")
            if not targets:
                return

            # ✅ 확인 메시지(한 번만)
            names = [t.text(0) for t in targets[:5]]
            more = "" if len(targets) <= 5 else f" 외 {len(targets)-5}개"
            msg = f"선택한 {len(targets)}개 항목을 삭제할까요?\n- " + "\n- ".join(names) + more
            ret = QMessageBox.question(
                self,
                "삭제 확인",
                msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ret != QMessageBox.StandardButton.Yes:
                return

            # ✅ 안전한 삭제 순서: 깊은 항목부터(자식 먼저)
            def _depth(it: QTreeWidgetItem) -> int:
                d = 0
                p = it.parent()
                while p:
                    d += 1
                    p = p.parent()
                return d
            targets.sort(key=_depth, reverse=True)

            # ✅ 다중 삭제를 '한 번의 Undo'로 묶기
            with self._fav_bulk_edit(reason=f"delete:{len(targets)}"):
                for it in targets:
                    parent = it.parent() or self.fav_tree.invisibleRootItem()
                    idx = parent.indexOfChild(it)
                    print(f"[DBG][FAV][DEL] remove name='{it.text(0)}' depth={_depth(it)} idx={idx}")
                    parent.takeChild(idx)
            print("[DBG][FAV][DEL] DONE multi")
        except Exception:
            print("[ERR][FAV][DEL] exception")
            traceback.print_exc()

    def _on_fav_context_menu(self, pos):
        item = self._current_fav_item()
        menu = QMenu(self)

        act_add_group = QAction("그룹 추가", self)
        act_add_group.triggered.connect(self._add_group)
        menu.addAction(act_add_group)

        act_add_curr = QAction(_current_add_button_label(), self)
        act_add_curr.triggered.connect(self._add_section_from_current)
        menu.addAction(act_add_curr)

        act_add_other = QAction("다른 창 추가", self)
        act_add_other.triggered.connect(self._add_section_from_other_window)
        menu.addAction(act_add_other)

        # 복사/붙여넣기 메뉴
        menu.addSeparator()

        act_copy = QAction("복사 (Ctrl+C)", self)
        act_copy.triggered.connect(self._copy_favorite_item)
        act_copy.setEnabled(item is not None)
        menu.addAction(act_copy)

        act_paste = QAction("붙여넣기 (Ctrl+V)", self)
        act_paste.triggered.connect(self._paste_favorite_item)
        act_paste.setEnabled(self.clipboard_data is not None)
        menu.addAction(act_paste)

        if item:
            try:
                item_type = item.data(0, ROLE_TYPE)
            except Exception:
                item_type = None
            if item_type in ("section", "notebook"):
                menu.addSeparator()
                act_send_to_codex = QAction("코덱스 작업 위치로 보내기", self)
                act_send_to_codex.triggered.connect(
                    lambda checked=False, fav_item=item: self._sync_codex_target_from_fav_item(
                        fav_item,
                        switch_to_codex=True,
                    )
                )
                menu.addAction(act_send_to_codex)

            menu.addSeparator()
            act_rename = QAction("이름바꾸기", self)
            act_rename.triggered.connect(self._rename_favorite_item)
            menu.addAction(act_rename)

            act_delete = QAction("삭제", self)
            act_delete.triggered.connect(self._delete_favorite_item)
            menu.addAction(act_delete)

        menu.exec(self.fav_tree.viewport().mapToGlobal(pos))

    def _on_fav_item_double_clicked(self, item: QTreeWidgetItem, col: int):
        if not item:
            return
        started_at = time.perf_counter()
        node_type = item.data(0, ROLE_TYPE)
        print(f"[DBG][FAV][DBLCLK] type={node_type} text='{item.text(0)}'")
        # ✅ notebook 타입도 더블클릭 동작해야 함
        if node_type not in ("section", "notebook"):
            return
        self._sync_codex_target_from_fav_item(item)
        self._activate_favorite_section(item, started_at=started_at)

    # _activate_favorite_notebook 제거됨 (기능 통합)

    def _cancel_pending_center_after_activate(self):
        try:
            if self._pending_center_timer is not None:
                self._pending_center_timer.stop()
                self._pending_center_timer.deleteLater()
        except Exception:
            pass
        self._pending_center_timer = None
        worker = self._center_worker
        self._center_worker = None
        if worker is not None:
            try:
                if worker.isRunning():
                    worker.requestInterruption()
                    worker.wait(500)
            except Exception:
                pass

    def _cancel_pending_favorite_activation(self):
        worker = self._favorite_activation_worker
        self._favorite_activation_worker = None
        if worker is not None:
            try:
                if worker.isRunning():
                    worker.requestInterruption()
                    worker.wait(500)
            except Exception:
                pass

    def _retain_qthread_until_finished(self, worker: Optional[QThread], attr_name: str):
        if worker is None:
            return
        self._retained_qthreads.append(worker)

        def _cleanup():
            try:
                if getattr(self, attr_name, None) is worker:
                    setattr(self, attr_name, None)
            except Exception:
                pass
            try:
                self._retained_qthreads.remove(worker)
            except ValueError:
                pass
            try:
                worker.deleteLater()
            except Exception:
                pass

        worker.finished.connect(_cleanup)

    def _mark_favorite_item_stale(
        self,
        item: Optional[QTreeWidgetItem],
        fallback_name: str,
    ) -> str:
        current_name = ""
        if item is not None:
            try:
                current_name = item.text(0) or ""
            except Exception:
                current_name = ""

        base_name = current_name or fallback_name or ""
        stale_prefixes = ("(구) ", "(old) ")
        if not base_name:
            return ""
        for prefix in stale_prefixes:
            if base_name.startswith(prefix):
                return base_name[len(prefix):]
        return base_name

    def _restore_favorite_item_from_stale(
        self,
        item: Optional[QTreeWidgetItem],
        fallback_name: str,
    ) -> Dict[str, Any]:
        current_name = ""
        if item is not None:
            try:
                current_name = item.text(0) or ""
            except Exception:
                current_name = ""

        result = {
            "display_name": current_name or fallback_name or "",
            "changed": False,
            "was_stale": False,
        }
        if item is None:
            return result

        stale_prefixes = ("(구) ", "(old) ")
        restored_name = current_name or fallback_name or ""
        for prefix in stale_prefixes:
            if restored_name.startswith(prefix):
                restored_name = restored_name[len(prefix):]
                result["was_stale"] = True
                break

        if restored_name and restored_name != current_name:
            try:
                item.setText(0, restored_name)
                self._save_favorites()
                result["changed"] = True
            except Exception:
                pass
        result["display_name"] = restored_name or current_name or fallback_name or ""
        return result

_publish_context(globals())

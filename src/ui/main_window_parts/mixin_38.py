# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin38:

    def _sorted_copy_nodes_by_name(self, nodes: Any) -> List:
        """
        종합(aggregate) 버퍼에서만 사용하는 표시용 정렬.
        - 원본 nodes를 변형하지 않기 위해 deepcopy 후 정렬
        - group이 있으면 children도 재귀 정렬
        """
        if not isinstance(nodes, list):
            return nodes if isinstance(nodes, list) else (nodes or [])
        try:
            copied = copy.deepcopy(nodes)
        except Exception:
            copied = list(nodes)

        def _disp_name(n: Any) -> str:
            if not isinstance(n, dict):
                return ""
            # group / buffer / section / notebook 모두 name 우선
            name = n.get("name")
            if name:
                return name
            t = n.get("target") or {}
            return t.get("section_text") or t.get("notebook_text") or ""

        def _rec(lst: List) -> List:
            # children 먼저 정렬
            for it in lst:
                if isinstance(it, dict) and isinstance(it.get("children"), list):
                    it["children"] = _rec(it["children"])
            try:
                lst.sort(key=lambda n: _name_sort_key(_disp_name(n)))
            except Exception:
                pass
            return lst

        return _rec(copied)

    # ----------------- FavoritesTree Undo/Redo helpers -----------------
    def _fav_reset_undo_context_from_data(self, data, *, reason: str = "") -> None:
        """
        2패널(모듈영역) Undo/Redo가 이상해지는 핵심 원인은
        '버퍼 전환 후에도 이전 버퍼의 _fav_last_snapshot / undo stack이 유지'되어
        Ctrl+Z가 다른 버퍼 스냅샷을 불러오는 케이스가 생기는 것입니다.

        버퍼/그룹 선택으로 중앙 트리를 로드한 직후,
        로드된 데이터 기준으로 undo/redo 컨텍스트를 초기화합니다.

        - _fav_last_snapshot: 현재 상태(초기 스냅샷)
        - _fav_undo_stack / _fav_redo_stack: 비움

        이렇게 해야 첫 변경 저장 시 "초기 스냅샷"이 undo에 들어가며,
        Ctrl+Z가 현재 버퍼 내부에서만 정상 동작합니다.
        """
        try:
            data_list = data if isinstance(data, list) else []
            snap = None
            if (
                isinstance(data, list)
                and id(data) == getattr(self, "_last_center_payload_source_id", 0)
            ):
                snap = getattr(self, "_last_center_payload_snapshot", None)
            if snap is None:
                snap = json.dumps(data_list, sort_keys=True, ensure_ascii=False)
        except Exception:
            snap = ""
        try:
            self._fav_undo_stack.clear()
            self._fav_redo_stack.clear()
        except Exception:
            self._fav_undo_stack = []
            self._fav_redo_stack = []
        self._fav_last_snapshot = snap
        try:
            self._fav_last_persisted_hash = None
            if snap == getattr(self, "_last_center_payload_snapshot", None):
                self._fav_last_persisted_hash = getattr(
                    self, "_last_center_payload_hash", None
                )
            if not self._fav_last_persisted_hash:
                self._fav_last_persisted_hash = hashlib.md5(
                    snap.encode("utf-8")
                ).hexdigest()
        except Exception:
            self._fav_last_persisted_hash = None
        try:
            tag = f" reason={reason}" if reason else ""
            self._dbg_hot(f"[DBG][FAV][UNDO_CTX] reset{tag} undo=0 redo=0 snap_len={len(snap)}")
        except Exception:
            pass

    # ----------------- 15-3. 버퍼 트리 이벤트 핸들러 -----------------
    def _on_buffer_tree_item_clicked(self, item, col):
        """버퍼 트리 항목 클릭 시 처리"""
        if not item:
            return
        node_type = item.data(0, ROLE_TYPE)
        payload = item.data(0, ROLE_DATA) or {}

        if node_type == "buffer":
            # 버퍼 전환 직전: 현재 중앙 트리 내용을 "이전 버퍼"에 반드시 저장
            # (그렇지 않으면 버퍼를 다시 클릭했을 때 섹션/그룹이 사라지는 현상 발생)
            flushed_current_buffer = False
            try:
                flushed_current_buffer = self._flush_pending_favorites_save()
            except Exception:
                pass
            if (
                self.active_buffer_id
                and payload
                and payload.get("id") != self.active_buffer_id
                and not flushed_current_buffer
            ):
                self._save_favorites()

            self.active_buffer_id = payload.get("id")
            self.active_buffer_node = payload  # Dict payload(스냅샷)
            self.active_buffer_item = item
            self._active_buffer_settings_node = _find_buffer_node_by_id(
                self.settings.get("favorites_buffers", []),
                self.active_buffer_id,
            )
            self.settings["active_buffer_id"] = self.active_buffer_id

            try:
                # ✅ 종합 버퍼: 전자필기장(노트북) 저장용
                if payload.get("id") == AGG_BUFFER_ID:
                    saved = payload.get("data", []) or []
                    self._dbg_node_type_counts(saved, "AGG_SAVED")

                    # 1) 종합 버퍼에 notebook/group 저장이 있으면: 그걸 기준으로 분류 표시한다.
                    if self._nodes_have_any_type(saved, {"notebook", "group"}):
                        self._dbg_hot(f"[DBG][AGG] load SAVED data (len={len(saved)})")
                        data_to_load = self._build_aggregate_categorized_display_nodes(saved)
                        self._load_favorites_into_center_tree(data_to_load)
                    else:
                        # 2) 저장된게 없을 때만: 기존 종합 계산 fallback
                        agg_data = self._build_aggregate_buffer()
                        self._dbg_node_type_counts(agg_data, "AGG_BUILT")
                        self._dbg_hot(f"[DBG][AGG] load BUILT aggregate (len={len(agg_data)})")
                        data_to_load = self._build_aggregate_categorized_display_nodes(agg_data)
                        self._load_favorites_into_center_tree(data_to_load)

                    # ✅ 버퍼 전환 직후: Undo/Redo 컨텍스트를 현재 버퍼 데이터로 리셋
                    self._fav_reset_undo_context_from_data(data_to_load, reason="buffer_switch:AGG")
                    self.btn_add_section_current.setEnabled(False)
                    self.btn_add_group.setEnabled(True)  # 원하면 그룹도 허용
                    if hasattr(self, "btn_register_all_notebooks"):
                        self.btn_register_all_notebooks.setEnabled(True)
                        self.btn_register_all_notebooks.setVisible(True)
                else:
                    data_to_load = payload.get("data", []) or []
                    self._load_favorites_into_center_tree(data_to_load)

                    # ✅ 버퍼 전환 직후: Undo/Redo 컨텍스트를 현재 버퍼 데이터로 리셋
                    self._fav_reset_undo_context_from_data(data_to_load, reason="buffer_switch")
                    self.btn_add_section_current.setEnabled(True)
                    if hasattr(self, "btn_register_all_notebooks"):
                        self.btn_register_all_notebooks.setEnabled(False)
                        self.btn_register_all_notebooks.setVisible(False)
                    self.btn_add_group.setEnabled(True)
                self._last_loaded_center_buffer_id = self.active_buffer_id
            finally:
                try:
                    self.buffer_tree.setUpdatesEnabled(True)
                    self.fav_tree.setUpdatesEnabled(True)
                    self.buffer_tree.viewport().update()
                    self.fav_tree.viewport().update()
                except Exception:
                    pass
        else:
            # 그룹 선택 시
            # 현재 버퍼 내용이 남아있을 수 있으므로 먼저 저장
            flushed_current_buffer = False
            try:
                flushed_current_buffer = self._flush_pending_favorites_save()
            except Exception:
                pass
            if self.active_buffer_id and not flushed_current_buffer:
                self._save_favorites()
            if hasattr(self, "btn_register_all_notebooks"):
                self.btn_register_all_notebooks.setEnabled(False)
                self.btn_register_all_notebooks.setVisible(False)
            self.btn_add_section_current.setEnabled(False)
            self.btn_add_group.setEnabled(False)
            self.active_buffer_id = None
            self.active_buffer_item = None
            self._active_buffer_settings_node = None
            self._last_loaded_center_buffer_id = None
            self._load_favorites_into_center_tree([])

            # ✅ 버퍼가 아닌(그룹/빈) 상태에서도 Undo 컨텍스트를 리셋 (이전 버퍼 스냅샷 혼입 방지)
            self._fav_reset_undo_context_from_data([], reason="buffer_switch:group_or_none")
        self._update_buffer_move_button_state()

    def _on_buffer_tree_selection_changed(self):
        """
        1패널에서 클릭/키보드 이동 등으로 "선택"만 바뀐 경우에도
        2패널(모듈/섹션)이 즉시 갱신되도록 한다.
        """
        if getattr(self, "_buf_sel_guard", False):
            return
        item = self.buffer_tree.currentItem()
        if not item:
            return

        node_type = item.data(0, ROLE_TYPE)
        payload = item.data(0, ROLE_DATA) or {}

        # 이미 활성 버퍼면 스킵(불필요 리로드 방지)
        if node_type == "buffer":
            cur_id = payload.get("id")
            if cur_id and self.active_buffer_id == cur_id:
                return

        self._buf_sel_guard = True
        try:
            # 기존 클릭 로직 재사용
            self._on_buffer_tree_item_clicked(item, 0)
        finally:
            self._buf_sel_guard = False

    def _on_buffer_tree_double_clicked(self, item, col):
        node_type = item.data(0, ROLE_TYPE)
        if node_type == "group":
            # 그룹이면 확장/축소 (기본 동작)
            pass
        else:
            # 버퍼면 이름 편집
            pass

    def _add_buffer_group(self):
        """새 버퍼 그룹 추가"""
        parent = self.buffer_tree.currentItem()
        # 버퍼가 선택되어 있으면 그 부모(그룹 또는 루트)에 추가
        if parent and parent.data(0, ROLE_TYPE) == "buffer":
            parent = parent.parent()

        parent = parent or self.buffer_tree.invisibleRootItem()

        node = {"type": "group", "name": "새 그룹", "children": []}
        item = self._append_buffer_node(parent, node)
        self.buffer_tree.setCurrentItem(item)
        self.buffer_tree.editItem(item, 0)
        self._save_buffer_structure()

    def _add_buffer(self):
        """새 버퍼 추가"""
        parent = self.buffer_tree.currentItem()
        # 버퍼가 선택되어 있으면 그 부모에 추가
        if parent and parent.data(0, ROLE_TYPE) == "buffer":
            parent = parent.parent()

        parent = parent or self.buffer_tree.invisibleRootItem()

        node = {"type": "buffer", "name": "새 버퍼", "data": []}
        item = self._append_buffer_node(parent, node)
        self.buffer_tree.setCurrentItem(item)
        self.buffer_tree.editItem(item, 0)
        # 새 버퍼가 생성되면 클릭 이벤트 강제 호출하여 활성화
        self._on_buffer_tree_item_clicked(item, 0)
        self._save_buffer_structure()

    def _rename_buffer(self):
        item = self.buffer_tree.currentItem()
        if item:
            self.buffer_tree.editItem(item, 0)

_publish_context(globals())

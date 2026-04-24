# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin36:

    def _refresh_project_buffer_search_highlights(self) -> None:
        search_input = getattr(self, "module_project_search_input", None)
        if search_input is None:
            return
        query = _normalize_project_search_key(search_input.text())
        if not query:
            if (
                self._buffer_search_highlighted_by_id
                or self._module_search_highlighted_by_id
            ):
                self._clear_project_buffer_search_highlights()
            return
        self._buffer_search_last_applied_key = ""
        self._highlight_project_buffers_from_module_search(
            search_input.text(),
            update_status=False,
            scroll=False,
            precomputed_query=query,
        )

    def _load_favorites_into_center_tree(self, node_data: List):
        """즐겨찾기 데이터를 중앙 트리에 로드합니다."""
        # ✅ 동일 데이터면 rebuild 스킵 (클릭 렉 제거 핵심)
        payload_raw = None
        source_id = id(node_data) if isinstance(node_data, list) else 0
        try:
            payload_raw = json.dumps(node_data, sort_keys=True, ensure_ascii=False)
        except Exception:
            payload_raw = None

        if (
            payload_raw is not None
            and payload_raw == getattr(self, "_last_center_payload_snapshot", None)
        ):
            self._last_center_payload_source_id = source_id
            return

        # 로딩 중 clear/append 과정에서 structureChanged/itemChanged가 발생하면
        # 선택 버퍼가 바뀌는 타이밍에 "빈 데이터"가 저장되는 문제가 발생할 수 있다.
        # (재현: 버퍼 A에서 섹션 추가 → 버퍼 B 클릭 → 다시 A 클릭 시 A가 빈 목록으로 덮임)
        self.fav_tree.blockSignals(True)
        was_updates_enabled = self.fav_tree.updatesEnabled()
        self.fav_tree.setUpdatesEnabled(False)
        try:
            self.fav_tree.clear()
            self._module_search_index = []
            self._module_search_last_match_records = []
            self._module_search_highlighted_by_id = {}
            self._module_search_match_count = 0
            t_build0 = time.perf_counter()
            for node in node_data:
                self._append_fav_node(self.fav_tree.invisibleRootItem(), node)

            build_ms = (time.perf_counter() - t_build0) * 1000.0
            total_nodes = -1
            if self._debug_perf_logs or self._debug_hotpaths:
                total_nodes = self._count_nodes_recursive(node_data)
            self._dbg_perf(
                f"[BOOT][PERF][FAV_REBUILD] total_nodes={total_nodes} build_ms={build_ms:.1f}"
            )

            # ✅ 2패널은 '그룹이 항상 펼쳐진 상태'가 기본 UX
            #    - 예전에는 최초 1회만 펼쳤는데, 그 이후엔 항상 접힌 상태로 복원되어 사용성이 나빠짐
            #    - expandAll() 대신 '자식이 있는 노드만 expandItem' 방식으로 그룹만 펼쳐 성능도 방어
            self._expand_fav_groups_always(total_nodes=total_nodes, reason="rebuild")
            self._rebuild_module_search_index()
            self._last_center_payload_hash = None
            self._last_center_payload_snapshot = payload_raw
            self._last_center_payload_source_id = source_id
        finally:
            self.fav_tree.setUpdatesEnabled(was_updates_enabled)
            self.fav_tree.blockSignals(False)
            if was_updates_enabled:
                self.fav_tree.viewport().update()
        self._refresh_project_buffer_search_highlights()

    def _request_favorites_save(self, *_args):
        """
        FavoritesTree 변경 신호 폭주를 짧게 묶어 1회 저장으로 합칩니다.
        - 이름변경(itemChanged 연속)
        - 드래그/구조변경(structureChanged + itemChanged 연쇄)
        - 붙여넣기 직후의 다중 signal
        """
        if not self.active_buffer_node:
            return
        if getattr(self, "_fav_undo_suspended", False):
            return
        self._module_search_index = []
        self._module_search_last_match_records = []
        self._buffer_search_last_applied_key = ""
        self._fav_save_pending = True
        self._fav_save_timer.start(self._fav_save_interval_ms)

    def _flush_pending_favorites_save(self):
        if not getattr(self, "_fav_save_pending", False):
            return False
        self._fav_save_pending = False
        self._save_favorites()
        return True

    def _save_favorites(self):
        """현재 활성화된 중앙 트리의 내용을 버퍼 트리의 해당 노드 데이터에 반영하고 저장합니다."""
        if not self.active_buffer_node:
            return
        self._invalidate_aggregate_cache(
            invalidate_classified_keys=self.active_buffer_id != AGG_BUFFER_ID
        )
        self._module_search_index = []
        self._module_search_last_match_records = []
        self._buffer_search_index = []
        self._buffer_search_last_match_records = []
        self._buffer_search_last_applied_key = ""

        try:
            if getattr(self, "_fav_save_timer", None) is not None and self._fav_save_timer.isActive():
                self._fav_save_timer.stop()
        except Exception:
            pass
        self._fav_save_pending = False

        # ✅ 종합 버퍼도 이제 '노트북 저장'을 위해 저장 허용

        try:
            data = []
            root = self.fav_tree.invisibleRootItem()
            for i in range(root.childCount()):
                data.append(self._serialize_fav_item(root.child(i)))

            try:
                snap = json.dumps(data, sort_keys=True, ensure_ascii=False)
            except Exception:
                snap = "[]"
            if (
                snap == getattr(self, "_fav_last_snapshot", None)
                and getattr(self, "_fav_last_persisted_hash", None) is None
            ):
                self._last_center_payload_hash = None
                self._last_center_payload_snapshot = snap
                self._last_center_payload_source_id = id(data)
                return
            try:
                current_hash = None
                if snap == getattr(self, "_last_center_payload_snapshot", None):
                    current_hash = getattr(self, "_last_center_payload_hash", None)
                if not current_hash:
                    current_hash = hashlib.md5(snap.encode("utf-8")).hexdigest()
            except Exception:
                current_hash = None

            # ✅ IMPORTANT:
            # 중앙 트리는 _load_favorites_into_center_tree()에서만 해시를 갱신한다는 가정이 있었는데,
            # 실제로는 paste/cut/delete 등으로 트리를 "직접" 수정한다.
            # 그러면 _last_center_payload_hash가 과거 상태(예: 7개)의 해시로 남아,
            # Undo가 과거 스냅샷(7개)을 로드하려 할 때 "해시가 같다"로 판단되어 리빌드가 스킵된다.
            # => Ctrl+Z가 안 먹는 것처럼 보이는 핵심 원인.
            self._last_center_payload_hash = current_hash
            self._last_center_payload_snapshot = snap
            self._last_center_payload_source_id = id(data)

            # ✅ 같은 상태를 signal 연쇄로 다시 저장하려는 경우 early-return
            #    - rename 편집 종료 / itemChanged 연속 호출 / 구조변경 후 중복 저장 방어
            #    - Undo/Redo 스냅샷은 동일 상태면 원래도 skip이므로 여기서 바로 끊어도 안전
            if current_hash is not None and current_hash == getattr(self, "_fav_last_persisted_hash", None):
                return

            # --- Undo/Redo: FavoritesTree 변경 스냅샷 기록 ---
            try:
                self._fav_record_snapshot(snap)
            except Exception:
                pass

            # 메모리 상의 active_buffer_node 데이터 업데이트
            if self.active_buffer_node is not None:
                self.active_buffer_node["data"] = data
                self._dbg_hot(f"[DBG][FAV][SAVE] Updated active_buffer_node data: count={len(data)}")

            # PyQt의 item.data()로 얻은 dict는 "수정해도 item 내부에 반영되지" 않는 경우가 있다.
            # 따라서 활성 버퍼의 QTreeWidgetItem에도 동일 데이터를 강제 주입한다.
            if self.active_buffer_item is None and self.active_buffer_id:
                self.active_buffer_item = self._buffer_item_index.get(self.active_buffer_id)
            if self.active_buffer_item is None and self.active_buffer_id:
                # 예외 상황 대비: ID로 다시 찾아서 연결
                iterator = QTreeWidgetItemIterator(self.buffer_tree)
                while iterator.value():
                    it = iterator.value()
                    p = it.data(0, ROLE_DATA) or {}
                    if p.get("id") == self.active_buffer_id:
                        self.active_buffer_item = it
                        break
                    iterator += 1

            if self.active_buffer_item is not None:
                p = self.active_buffer_item.data(0, ROLE_DATA) or {}
                p["data"] = data
                self.active_buffer_item.setData(0, ROLE_DATA, p)
                self._dbg_hot(f"[DBG][FAV][SAVE] Updated active_buffer_item payload: id={p.get('id')}")

            self._fav_last_persisted_hash = current_hash
            # 그리고 전체 버퍼 구조 저장
            settings_buffer = self._active_buffer_settings_node
            if settings_buffer is None:
                settings_buffer = _find_buffer_node_by_id(
                    self.settings.get("favorites_buffers", []),
                    self.active_buffer_id,
                )
                self._active_buffer_settings_node = settings_buffer
            if settings_buffer is not None:
                settings_buffer["data"] = data
                self.settings["active_buffer_id"] = self.active_buffer_id
                self._save_settings_to_file()
                self._dbg_hot("[DBG][FAV][SAVE] Active buffer data persisted without full structure rebuild")
            else:
                self._save_buffer_structure()
                self._dbg_hot("[DBG][FAV][SAVE] Fallback full buffer structure persist")

            if (
                self.active_buffer_id == AGG_BUFFER_ID
                and not getattr(self, "_aggregate_reclassify_in_progress", False)
            ):
                self._refresh_active_aggregate_classification_from_saved_data(
                    current_nodes=data,
                    persist=True,
                    show_status=False,
                )

        except Exception as e:
            print(f"[ERROR] 즐겨찾기 저장 실패: {e}")

    def _request_buffer_structure_save(self, *_args):
        if getattr(self, "_boot_loading", False):
            return
        self._buffer_save_timer.start(self._buffer_save_interval_ms)

    def _flush_pending_buffer_structure_save(self) -> None:
        try:
            if self._buffer_save_timer.isActive():
                self._buffer_save_timer.stop()
                self._save_buffer_structure()
        except Exception:
            pass

    def _save_buffer_structure(self):
        """버퍼 트리의 구조(그룹/버퍼)를 settings에 저장합니다."""
        self._invalidate_aggregate_cache()
        try:
            if self._buffer_save_timer.isActive():
                self._buffer_save_timer.stop()
        except Exception:
            pass
        self._buffer_item_index = {}
        self._first_buffer_item = None
        self._buffer_search_index = []
        self._buffer_search_last_match_records = []
        root = self.buffer_tree.invisibleRootItem()
        structure = []
        for i in range(root.childCount()):
            structure.append(self._serialize_buffer_item(root.child(i), rebuild_index=True))

        try:
            structure_sig = hashlib.md5(
                json.dumps(structure, sort_keys=True, ensure_ascii=False).encode("utf-8")
            ).hexdigest()
        except Exception:
            structure_sig = None

        if structure_sig is not None and structure_sig == getattr(self, "_last_saved_buffer_structure_sig", None):
            self.settings["favorites_buffers"] = structure
            self._active_buffer_settings_node = _find_buffer_node_by_id(
                self.settings.get("favorites_buffers", []),
                self.active_buffer_id,
            )
            self._refresh_project_buffer_search_highlights()
            return

        self.settings["favorites_buffers"] = structure
        # ✅ 저장 직전에 구조 강제 보정(순서/락/종합 유지)
        _ensure_default_and_aggregate_inplace(self.settings)
        self._last_saved_buffer_structure_sig = structure_sig
        self._active_buffer_settings_node = _find_buffer_node_by_id(
            self.settings.get("favorites_buffers", []),
            self.active_buffer_id,
        )
        self._save_settings_to_file()
        self._refresh_project_buffer_search_highlights()

    def _serialize_buffer_item(
        self,
        item: QTreeWidgetItem,
        *,
        rebuild_index: bool = False,
    ) -> Dict:
        node_type = item.data(0, ROLE_TYPE)
        payload = item.data(0, ROLE_DATA) or {}
        if rebuild_index:
            item_id = payload.get("id")
            if item_id:
                self._buffer_item_index[item_id] = item
                if node_type == "buffer" and self._first_buffer_item is None:
                    self._first_buffer_item = item
            if node_type == "buffer":
                search_key = self._project_buffer_search_key_for_item(item)
                if search_key:
                    parents = []
                    parent_item = item.parent()
                    while parent_item is not None:
                        parents.append(parent_item)
                        parent_item = parent_item.parent()
                    self._buffer_search_index.append(
                        {"item": item, "key": search_key, "parents": tuple(parents)}
                    )

        node = {
            "type": node_type,
            "id": payload.get("id"),
            "name": item.text(0)
        }

        if node_type == "group":
            if payload.get("locked"):
                node["locked"] = True
            children = []
            for i in range(item.childCount()):
                children.append(
                    self._serialize_buffer_item(
                        item.child(i),
                        rebuild_index=rebuild_index,
                    )
                )
            node["children"] = children
        else:
            # 버퍼인 경우, 현재 메모리 상의 데이터를 유지하거나
            # 활성 상태라면 현재 중앙 트리에서 가져와야 함.
            # payload['data']는 로드 시점의 스냅샷일 수 있으므로 주의.
            # 여기서는 payload['data']를 그대로 쓰고,
            # 활성 버퍼가 변경될 때마다 payload['data']를 갱신해두는 방식을 사용.
            if payload.get("virtual"):
                node["virtual"] = payload.get("virtual")
            if payload.get("locked"):
                node["locked"] = True

            node["data"] = payload.get("data", [])
            # [DBG] 종합 버퍼 저장 스캔
            if node.get("id") == AGG_BUFFER_ID:
                self._dbg_hot(f"[DBG][SSOT][SERIALIZE] Aggregate data count={len(node['data'])}")

        return node

    def _request_settings_save(self):
        self._settings_save_pending = True
        self._settings_save_timer.start(self._settings_save_interval_ms)

    def _flush_pending_settings_save(self):
        if self._settings_save_in_progress:
            return False
        timer_active = self._settings_save_timer.isActive()
        if not self._settings_save_pending and not timer_active:
            return False
        if timer_active:
            self._settings_save_timer.stop()
        self._settings_save_pending = False
        self._settings_save_in_progress = True
        try:
            return save_settings(self.settings)
        finally:
            self._settings_save_in_progress = False

_publish_context(globals())

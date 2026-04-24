# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin34:

    def _finish_boot_sequence(self):
        """부팅 완료 단계에서 마지막 상태(활성 버퍼 데이터)를 강제 복원합니다."""
        print("[BOOT] Starting final boot sequence...")
        t0 = time.perf_counter()
        try:
            # 활성 버퍼 다시 확인
            active_id = self.settings.get("active_buffer_id")
            found_data = []
            buf_name = "None"
            if active_id and getattr(self, "_last_loaded_center_buffer_id", None) == active_id:
                print(f"[BOOT][PERF] final restore skipped; active buffer already loaded: {active_id}")
                return

            found_item = self._buffer_item_index.get(active_id) if active_id else None
            if found_item is not None:
                payload = found_item.data(0, ROLE_DATA) or {}
                buf_name = found_item.text(0)
                if active_id == AGG_BUFFER_ID:
                    saved = payload.get("data", []) or []
                    source = saved if self._nodes_have_any_type(saved, {"notebook", "group"}) else self._build_aggregate_buffer()
                    found_data = self._build_aggregate_categorized_display_nodes(source)
                else:
                    found_data = payload.get("data", [])
            else:
                iterator = QTreeWidgetItemIterator(self.buffer_tree)
                while iterator.value():
                    item = iterator.value()
                    payload = item.data(0, ROLE_DATA) or {}
                    if payload.get("id") == active_id:
                        buf_name = item.text(0)
                        if active_id == AGG_BUFFER_ID:
                            saved = payload.get("data", []) or []
                            source = saved if self._nodes_have_any_type(saved, {"notebook", "group"}) else self._build_aggregate_buffer()
                            found_data = self._build_aggregate_categorized_display_nodes(source)
                        else:
                            found_data = payload.get("data", [])
                        break
                    iterator += 1

            # 강제 리빌드
            self._rebuild_modules_from_buffer(buf_name, found_data)
        except Exception as e:
            print(f"[BOOT][RESTORE][BUF_REBUILD_FAIL] {e}")
            traceback.print_exc()
        finally:
            self._boot_loading = False
            total_ms = (time.perf_counter() - t0) * 1000.0
            try:
                total_nodes = self._count_nodes_recursive(found_data)
            except Exception:
                total_nodes = len(found_data) if isinstance(found_data, list) else 0
            print(f"[BOOT][PERF] final_restore_ms={total_ms:.1f} total_nodes={total_nodes}")
            print(f"[BOOT] Final boot sequence finished. (Active: {buf_name})")

    def _rebuild_modules_from_buffer(self, buffer_name: str, nodes: list):
        """
        저장된 favorites_buffers 기준으로
        2패널(모듈/전자필기장 영역)을 복원합니다.
        """
        print(f"[BOOT][BUF_RESTORE] buffer='{buffer_name}' count={len(nodes)}")

        if nodes:
            self._load_favorites_into_center_tree(nodes)
        else:
            self._clear_module_panel()
        self._last_loaded_center_buffer_id = (
            self.settings.get("active_buffer_id") if buffer_name != "None" else None
        )

        # 활성 상태바 업데이트
        if buffer_name != "None":
            self.connection_status_label.setText(f"준비됨 (활성 버퍼: {buffer_name})")

    def _clear_module_panel(self):
        """중앙 모듈(즐겨찾기) 패널을 완전히 비웁니다."""
        was_updates_enabled = True
        try:
            self.fav_tree.blockSignals(True)
            was_updates_enabled = self.fav_tree.updatesEnabled()
            self.fav_tree.setUpdatesEnabled(False)
            self.fav_tree.clear()
            self._last_loaded_center_buffer_id = None
            self._last_center_payload_snapshot = None
            self._last_center_payload_source_id = 0
            self._last_center_payload_hash = None # 해시 캐시 초기화
            self._module_search_index = []
            self._module_search_last_match_records = []
            self._module_search_highlighted_by_id = {}
            self._module_search_match_count = 0
        except Exception:
            pass
        finally:
            try:
                self.fav_tree.setUpdatesEnabled(was_updates_enabled)
                if was_updates_enabled:
                    self.fav_tree.viewport().update()
            except Exception:
                pass
            self.fav_tree.blockSignals(False)

    # ----------------- 15. 즐겨찾기 로드/세이브 (계층형 버퍼 시스템 적용) -----------------
    def _load_buffers_and_favorites(self):
        """설정에서 버퍼 트리를 로드합니다."""
        # ✅ 로드 전에 강제 보정 (UI에서 깨져도 복구)
        _ensure_default_and_aggregate_inplace(self.settings)
        self._invalidate_aggregate_cache()

        self.buffer_tree.blockSignals(True)
        self.buffer_tree.clear()
        self._buffer_item_index = {}
        self._first_buffer_item = None
        self._buffer_search_index = []
        self._buffer_search_last_match_records = []
        self._buffer_search_highlighted_by_id = {}
        self._buffer_search_match_count = 0
        self._buffer_search_last_applied_key = ""
        self._buffer_search_last_first_match_id = 0
        self._active_buffer_settings_node = None

        buffers_data = self.settings.get("favorites_buffers", [])

        # ✅ 레거시 설정(딕트/버퍼 없는 리스트) 자동 마이그레이션
        tmp = {
            "favorites_buffers": buffers_data,
            "active_buffer_id": self.settings.get("active_buffer_id"),
            "active_buffer": self.settings.get("active_buffer"),
        }
        if _migrate_favorites_buffers_inplace(tmp):
            self.settings.update(tmp)
            buffers_data = self.settings.get("favorites_buffers", [])
            try:
                self._save_settings_to_file()
            except Exception:
                pass

        # 방어: 그래도 dict면 비워서 크래시 방지
        if isinstance(buffers_data, dict):
            buffers_data = []

        self._boot_loading = True
        try:
            for node in buffers_data:
                self._append_buffer_node(self.buffer_tree.invisibleRootItem(), node)

            try:
                # 시작 시 프로젝트 영역은 항상 전체 펼침 상태로 보여준다.
                # 기존 expandToDepth(1)은 그룹 아래 하위 폴더가 접힌 채 남아서
                # 사용자가 앱을 켤 때마다 다시 열어야 했다.
                self._expand_buffer_groups_always(reason="startup")
            except Exception:
                pass
            self.buffer_tree.blockSignals(False)

            # 활성 버퍼 복원
            active_id = self.settings.get("active_buffer_id")
            found_item = self._buffer_item_index.get(active_id) if active_id else None

            if active_id and not found_item:
                # 트리를 순회하며 ID 찾기
                iterator = QTreeWidgetItemIterator(self.buffer_tree)
                while iterator.value():
                    item = iterator.value()
                    payload = item.data(0, ROLE_DATA)
                    if payload and payload.get("id") == active_id:
                        found_item = item
                        break
                    iterator += 1

            # 못 찾았으면 첫 번째 버퍼 선택
            if not found_item:
                found_item = self._first_buffer_item
            if not found_item:
                iterator = QTreeWidgetItemIterator(self.buffer_tree)
                while iterator.value():
                    item = iterator.value()
                    if item.data(0, ROLE_TYPE) == "buffer":
                        found_item = item
                        break
                    iterator += 1

            if found_item:
                self.buffer_tree.setCurrentItem(found_item)
                self._on_buffer_tree_item_clicked(found_item, 0)
        finally:
            self._boot_loading = False
            try:
                self.buffer_tree.blockSignals(False)
                self.buffer_tree.viewport().update()
            except Exception:
                pass
            self._refresh_project_buffer_search_highlights()

    def _append_buffer_node(self, parent: QTreeWidgetItem, node: Dict[str, Any]) -> QTreeWidgetItem:
        item = QTreeWidgetItem(parent)
        node_type = node.get("type", "buffer")
        name = node.get("name", "이름 없음")
        item.setText(0, name)
        item.setData(0, ROLE_TYPE, node_type)

        # 데이터(즐겨찾기 목록)는 트리에 직접 저장하지 않고,
        # 구조 변경 시 settings에서 다시 읽거나 관리함.
        # 여기서는 ID와 데이터 참조를 위해 payload 저장
        payload = {
            "id": node.get("id", str(uuid.uuid4())),
            "data": node.get("data", []),  # 버퍼인 경우 데이터
            "virtual": node.get("virtual"),  # 종합(aggregate) 등
            "locked": bool(node.get("locked", False)),
        }

        if node_type == "group":
            icon = getattr(self, "_icon_dir", None) or self.style().standardIcon(QApplication.style().StandardPixmap.SP_DirIcon)
            item.setIcon(0, icon)
            for child in node.get("children", []):
                self._append_buffer_node(item, child)
        else:
            # ✅ 종합(가상) 버퍼는 전용 아이콘(컴퓨터)로 표시
            if payload.get("virtual") == "aggregate":
                icon = getattr(self, "_icon_agg", None) or self.style().standardIcon(QApplication.style().StandardPixmap.SP_ComputerIcon)
                item.setIcon(0, icon)
            else:
                icon = getattr(self, "_icon_file", None) or self.style().standardIcon(QApplication.style().StandardPixmap.SP_FileIcon)
                item.setIcon(0, icon)

        item.setData(0, ROLE_DATA, payload)
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

        # ✅ locked 노드는 편집/이동/드롭 막기 (Default 그룹, 종합)
        if payload.get("locked"):
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable & ~Qt.ItemFlag.ItemIsDragEnabled & ~Qt.ItemFlag.ItemIsDropEnabled)
        else:
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsDragEnabled | Qt.ItemFlag.ItemIsDropEnabled)

        return item

    def _expand_fav_groups_always(self, *, total_nodes: int = -1, reason: str = "") -> None:
        """2패널(중앙 트리)에서 그룹 노드를 기본으로 펼쳐둡니다.

        주의: fav_tree.expandAll()은 노드가 많을 때 렉을 유발할 수 있어,
        '자식이 있는 노드만 expandItem()' 하는 방식으로 그룹만 펼칩니다.
        """
        changed = False
        was_updates_enabled = True
        try:
            root = self.fav_tree.invisibleRootItem()
            stack = [root.child(i) for i in range(root.childCount() - 1, -1, -1)]
            expanded = 0
            was_updates_enabled = self.fav_tree.updatesEnabled()
            self.fav_tree.setUpdatesEnabled(False)
            while stack:
                it = stack.pop()
                # 자식이 있으면 '그룹성 노드'로 간주
                if it.childCount() > 0:
                    if not it.isExpanded():
                        self.fav_tree.expandItem(it)
                        expanded += 1
                        changed = True
                    for j in range(it.childCount()):
                        stack.append(it.child(j))
            tag = f" reason={reason}" if reason else ""
            self._dbg_hot(f"[DBG][FAV][EXPAND_GROUPS]{tag} expanded={expanded} total_nodes={total_nodes}")
        except Exception as e:
            try:
                tag = f" reason={reason}" if reason else ""
                self._dbg_hot(f"[DBG][FAV][EXPAND_GROUPS][FAIL]{tag} {e}")
            except Exception:
                pass
        finally:
            try:
                self.fav_tree.setUpdatesEnabled(was_updates_enabled)
                if changed and was_updates_enabled:
                    self.fav_tree.viewport().update()
            except Exception:
                pass

    def _collapse_fav_groups_always(self, *, reason: str = "") -> None:
        changed = False
        was_updates_enabled = True
        try:
            root = self.fav_tree.invisibleRootItem()
            stack = [root.child(i) for i in range(root.childCount() - 1, -1, -1)]
            collapsed = 0
            was_updates_enabled = self.fav_tree.updatesEnabled()
            self.fav_tree.setUpdatesEnabled(False)
            while stack:
                it = stack.pop()
                if it.childCount() > 0:
                    for j in range(it.childCount() - 1, -1, -1):
                        stack.append(it.child(j))
                    if it.isExpanded():
                        self.fav_tree.collapseItem(it)
                        collapsed += 1
                        changed = True
            tag = f" reason={reason}" if reason else ""
            self._dbg_hot(f"[DBG][FAV][COLLAPSE_GROUPS]{tag} collapsed={collapsed}")
        except Exception as e:
            try:
                tag = f" reason={reason}" if reason else ""
                self._dbg_hot(f"[DBG][FAV][COLLAPSE_GROUPS][FAIL]{tag} {e}")
            except Exception:
                pass
        finally:
            try:
                self.fav_tree.setUpdatesEnabled(was_updates_enabled)
                if changed and was_updates_enabled:
                    self.fav_tree.viewport().update()
            except Exception:
                pass

    def _project_search_text_from_target(self, target: Any) -> List[str]:
        if not isinstance(target, dict):
            return []
        parts = []
        for key in ("notebook_text", "section_text", "display_text", "path"):
            value = target.get(key)
            if isinstance(value, str) and value.strip():
                parts.append(value.strip())
        return parts

    def _project_search_text_from_nodes(
        self, nodes: Any, max_parts: int = 1200
    ) -> List[str]:
        if not isinstance(nodes, list):
            return []
        parts = []
        stack = list(reversed(nodes))
        while stack and len(parts) < max_parts:
            node = stack.pop()
            if not isinstance(node, dict):
                continue
            name = node.get("name")
            if isinstance(name, str) and name.strip():
                parts.append(name.strip())
            parts.extend(self._project_search_text_from_target(node.get("target")))
            children = node.get("children")
            if isinstance(children, list):
                stack.extend(reversed(children))
            data = node.get("data")
            if isinstance(data, list):
                stack.extend(reversed(data))
        return parts

_publish_context(globals())

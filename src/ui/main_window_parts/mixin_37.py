# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin37:

    def _save_settings_to_file(self, immediate: bool = False):
        """현재 self.settings 객체를 파일에 저장합니다."""
        if immediate:
            self._settings_save_pending = True
            self._flush_pending_settings_save()
        else:
            self._request_settings_save()


    def _backup_full_settings(self):
        """전체 설정을 백업합니다."""
        last_dir = self.settings.get("last_backup_dir", os.getcwd())
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        default_filename = f"OneNote_Remocon_Backup_{timestamp}.json"

        default_path = os.path.join(last_dir, default_filename)

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "설정 백업",
            default_path,
            "JSON Files (*.json);;All Files (*)",
        )

        if file_path:
            try:
                # 현재 메모리 설정을 파일로 강제 저장 (최신 상태 반영)
                self._save_window_state() # 창 위치 등 업데이트
                self._flush_pending_buffer_structure_save()
                self._save_favorites()    # 즐겨찾기 업데이트

                # 백업 디렉토리 기억
                self.settings["last_backup_dir"] = os.path.dirname(file_path)

                # _write_json을 사용하여 안전하게 저장
                _write_json(file_path, self.settings)

                # 설정 파일에도 last_backup_dir 반영하여 저장
                self._save_settings_to_file(immediate=True)

                QMessageBox.information(
                    self, "백업 완료", f"성공적으로 백업되었습니다.\n\n경로: {file_path}"
                )
            except Exception as e:
                QMessageBox.critical(
                    self, "백업 실패", f"백업 중 오류가 발생했습니다:\n{e}"
                )

    def _reload_settings_after_path_change(self) -> None:
        self.settings = load_settings()
        try:
            self._load_buffers_and_favorites()
            self._update_move_button_state()
        except Exception as e:
            print(f"[WARN] 설정 경로 변경 후 UI 재로드 실패: {e}")
        try:
            self.connection_status_label.setText(_settings_path_mode_label())
        except Exception:
            pass

    def _choose_shared_settings_json(self):
        """프로젝트 실행과 EXE 실행이 같이 쓸 공용 설정 JSON을 지정합니다."""
        current_path = _get_settings_file_path()
        default_dir = os.path.dirname(current_path) if current_path else os.getcwd()
        default_path = os.path.join(default_dir, SETTINGS_FILE)

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "공용 설정 JSON 위치 지정",
            default_path,
            "JSON Files (*.json);;All Files (*)",
        )
        if not file_path:
            return

        file_path = _expand_external_settings_path(file_path)
        try:
            self._save_window_state()
            self._flush_pending_buffer_structure_save()
            self._save_favorites()

            use_existing = False
            if os.path.exists(file_path):
                answer = QMessageBox.question(
                    self,
                    "공용 설정 JSON 선택",
                    (
                        "선택한 JSON 파일이 이미 있습니다.\n\n"
                        "예: 기존 JSON을 불러와 사용합니다.\n"
                        "아니오: 현재 앱 설정으로 해당 JSON을 덮어씁니다.\n"
                        "취소: 변경하지 않습니다."
                    ),
                    (
                        QMessageBox.StandardButton.Yes
                        | QMessageBox.StandardButton.No
                        | QMessageBox.StandardButton.Cancel
                    ),
                    QMessageBox.StandardButton.Yes,
                )
                if answer == QMessageBox.StandardButton.Cancel:
                    return
                use_existing = answer == QMessageBox.StandardButton.Yes

            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            if not use_existing:
                payload = dict(self.settings)
                payload.pop("favorites", None)
                _ensure_default_and_aggregate_inplace(payload)
                _write_json(file_path, payload)

            _set_external_settings_file_path(file_path)
            _SETTINGS_OBJECT_CACHE.clear()
            _JSON_TEXT_CACHE.clear()
            self._reload_settings_after_path_change()
            QMessageBox.information(
                self,
                "공용 설정 JSON 연결 완료",
                (
                    "이제 프로젝트 실행과 EXE 실행이 같은 설정 JSON을 사용합니다.\n\n"
                    f"경로: {file_path}\n\n"
                    "다른 PC에서는 OneDrive가 동기화된 로컬 경로를 한 번 지정하면 됩니다."
                ),
            )
        except Exception as e:
            QMessageBox.critical(self, "공용 설정 JSON 연결 실패", str(e))

    def _show_settings_json_path(self):
        QMessageBox.information(
            self,
            "현재 설정 JSON 위치",
            (
                f"{_settings_path_mode_label()}\n\n"
                f"공용 경로 포인터: {_settings_path_config_file()}\n"
                f"환경변수 우선순위: {SETTINGS_PATH_ENV}"
            ),
        )

    def _open_settings_json_folder(self):
        path = _get_settings_file_path()
        folder = os.path.dirname(path)
        try:
            os.makedirs(folder, exist_ok=True)
            open_path_in_system(folder)
        except Exception as e:
            QMessageBox.warning(self, "폴더 열기 실패", str(e))

    def _clear_shared_settings_json(self):
        if not _get_external_settings_file_path():
            QMessageBox.information(self, "안내", "현재 공용 설정 JSON 연결이 없습니다.")
            return
        answer = QMessageBox.question(
            self,
            "공용 설정 JSON 연결 해제",
            (
                "공용 설정 JSON 연결을 해제하고 기본 위치의 설정 JSON을 사용합니다.\n"
                "현재 공용 JSON 파일 자체는 삭제하지 않습니다.\n\n"
                "계속하시겠습니까?"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            self._save_settings_to_file(immediate=True)
            _clear_external_settings_file_path()
            _SETTINGS_OBJECT_CACHE.clear()
            _JSON_TEXT_CACHE.clear()
            self._reload_settings_after_path_change()
            QMessageBox.information(
                self,
                "연결 해제 완료",
                f"이제 기본 설정 JSON을 사용합니다.\n\n{_get_settings_file_path()}",
            )
        except Exception as e:
            QMessageBox.critical(self, "연결 해제 실패", str(e))

    def _restore_full_settings(self):
        """설정 파일을 복원합니다."""
        last_dir = self.settings.get("last_backup_dir", os.getcwd())

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "설정 복원",
            last_dir,
            "JSON Files (*.json);;All Files (*)",
        )

        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    restored_data = json.load(f)

                # 마이그레이션 적용 (구버전 백업일 경우 대비)
                if _migrate_favorites_buffers_inplace(restored_data):
                    print("[INFO] 복원 데이터 마이그레이션 수행됨")

                confirm = QMessageBox.question(
                    self,
                    "복원 확인",
                    "설정을 복원하면 현재 설정이 덮어씌워집니다.\n계속하시겠습니까?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )

                if confirm != QMessageBox.StandardButton.Yes:
                    return

                # 복원 디렉토리 기억
                restored_data["last_backup_dir"] = os.path.dirname(file_path)

                # 현재 설정 교체
                self.settings = DEFAULT_SETTINGS.copy()
                self.settings.update(restored_data)

                # 파일에 즉시 반영
                self._save_settings_to_file(immediate=True)

                # UI 리로드
                # 1. 버퍼/즐겨찾기 트리 갱신
                self._load_buffers_and_favorites()

                # 2. 스플리터 위치 등은 재시작 시 적용되거나 지금 강제 적용 가능하나
                # 여기서는 데이터 갱신에 집중

                QMessageBox.information(
                    self, "복원 완료", "설정이 복원되었습니다."
                )

            except Exception as e:
                QMessageBox.critical(
                    self, "복원 실패", f"복원 중 오류가 발생했습니다:\n{e}"
                )

    def _serialize_fav_item(self, item: QTreeWidgetItem) -> Dict[str, Any]:
        node_type = item.data(0, ROLE_TYPE)
        payload = item.data(0, ROLE_DATA) or {}
        node = {
            "type": node_type,
            "id": payload.get("id") or str(uuid.uuid4()),
            "name": item.text(0),
        }
        if node_type in ("section", "notebook"):
            node["target"] = payload.get("target", {})
        if node_type == "notebook" and bool(payload.get("is_open")):
            node["is_open"] = True
        children = []
        for i in range(item.childCount()):
            children.append(self._serialize_fav_item(item.child(i)))
        if children:
            node["children"] = children
        return node

    def _append_fav_node(
        self, parent: QTreeWidgetItem, node: Dict[str, Any]
    ) -> QTreeWidgetItem:
        item = QTreeWidgetItem(parent)
        node_type = node.get("type", "group")
        raw_name = str(node.get("name", "이름 없음") or "이름 없음")
        name = (
            _strip_stale_favorite_prefix(raw_name)
            if node_type in ("section", "notebook")
            else raw_name
        )
        item.setText(0, name)
        item.setData(0, ROLE_TYPE, node_type)
        payload = {"id": node.get("id") or str(uuid.uuid4())}
        if node_type in ("section", "notebook"):
            target = node.get("target", {})
            payload["target"] = target
            is_open_notebook = node_type == "notebook" and bool(
                node.get("is_open")
                or node.get("open")
                or (target or {}).get("is_open")
            )
            if is_open_notebook:
                item.setData(0, ROLE_OPEN_NOTEBOOK, True)
                payload["is_open"] = True
                item.setToolTip(0, "현재 OneNote에 열려 있는 전자필기장")
            icon = getattr(self, "_icon_file", None)
            if icon is not None:
                item.setIcon(0, icon)
        else:
            icon = getattr(self, "_icon_dir", None)
            if icon is not None:
                item.setIcon(0, icon)
        item.setData(0, ROLE_DATA, payload)
        flags = getattr(self, "_fav_tree_item_flags", None)
        if flags is None:
            flags = (
                item.flags()
                | Qt.ItemFlag.ItemIsEditable
                | Qt.ItemFlag.ItemIsDragEnabled
                | Qt.ItemFlag.ItemIsDropEnabled
                | Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
            )
            self._fav_tree_item_flags = flags
        item.setFlags(flags)
        for ch in node.get("children", []):
            self._append_fav_node(item, ch)
        return item

    def _dbg_node_type_counts(self, nodes, tag=""):
        try:
            if not self._debug_hotpaths:
                return
            cnt = {"notebook": 0, "section": 0, "page": 0, "group": 0, "buffer": 0, "unknown": 0}
            def rec(arr):
                if not isinstance(arr, list):
                    return
                for n in arr:
                    if not isinstance(n, dict): continue
                    ty = n.get("type", "unknown")
                    cnt[ty] = cnt.get(ty or "unknown", 0) + 1
                    rec(n.get("children", []))
                    rec(n.get("data", []))
            rec(nodes)
            self._dbg_hot(f"[DBG][NODE_TYPES]{'['+tag+']' if tag else ''} total={len(nodes) if isinstance(nodes, list) else 'NA'} {cnt}")
        except Exception as e:
            print(f"[DBG][NODE_TYPES][FAIL] {e}")

    def _nodes_have_type(self, nodes, ty: str) -> bool:
        if not isinstance(nodes, list):
            return False
        for n in nodes:
            if isinstance(n, dict) and n.get("type") == ty:
                return True
        return False

    def _nodes_have_any_type(self, nodes, types) -> bool:
        return isinstance(nodes, list) and any(isinstance(n, dict) and n.get("type") in types for n in nodes)

    def _count_nodes_recursive(self, nodes) -> int:
        """node(list[dict])의 총 노드 수(자식 포함)를 빠르게 계산한다."""
        if not isinstance(nodes, list):
            return 0
        total = 0
        stack = list(nodes)
        while stack:
            n = stack.pop()
            if not isinstance(n, dict):
                continue
            total += 1
            ch = n.get("children")
            if isinstance(ch, list) and ch:
                stack.extend(ch)
        return total

_publish_context(globals())

# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin14:

    def _write_codex_targets(self, targets: List[Dict[str, str]]) -> None:
        path = self._codex_targets_path()
        payload = {
            "version": 1,
            "targets": targets,
        }
        self._write_json_file_atomic(path, payload)
        try:
            self._codex_targets_cache = (
                os.path.getmtime(path),
                [dict(t) for t in targets if isinstance(t, dict)],
            )
        except Exception:
            self._codex_targets_cache = None

    def _selected_codex_target_profile(self) -> Dict[str, str]:
        combo = getattr(self, "codex_target_combo", None)
        if combo is not None:
            data = combo.currentData()
            if isinstance(data, dict):
                return data
        try:
            profile = self._codex_target_from_fields()
            if profile.get("path") or profile.get("notebook"):
                return profile
        except Exception:
            pass
        return self._default_codex_targets()[0]

    def _codex_target_from_fields(self) -> Dict[str, str]:
        def _text(attr: str) -> str:
            widget = getattr(self, attr, None)
            if widget is None:
                return ""
            try:
                return widget.text().strip()
            except Exception:
                return ""

        return {
            "name": _text("codex_target_name_input") or "새 대상",
            "path": _text("codex_target_path_input"),
            "notebook": _text("codex_target_notebook_input"),
            "section_group": _text("codex_target_group_input"),
            "section": _text("codex_target_section_input"),
            "section_group_id": _text("codex_target_group_id_input"),
            "section_id": _text("codex_target_section_id_input"),
        }

    def _populate_codex_target_fields(self, profile: Dict[str, str]) -> None:
        mapping = {
            "codex_target_name_input": profile.get("name", ""),
            "codex_target_path_input": profile.get("path", ""),
            "codex_target_notebook_input": profile.get("notebook", ""),
            "codex_target_group_input": profile.get("section_group", ""),
            "codex_target_section_input": profile.get("section", ""),
            "codex_target_group_id_input": profile.get("section_group_id", ""),
            "codex_target_section_id_input": profile.get("section_id", ""),
        }
        for attr, value in mapping.items():
            widget = getattr(self, attr, None)
            if widget is not None:
                if widget.text() == value:
                    continue
                widget.blockSignals(True)
                widget.setText(value)
                widget.blockSignals(False)

    def _refresh_codex_target_combo(self, selected_name: Optional[str] = None) -> None:
        combo = getattr(self, "codex_target_combo", None)
        if combo is None:
            return
        current_name = selected_name or combo.currentText()
        combo.blockSignals(True)
        combo.clear()
        for profile in self._load_codex_targets():
            combo.addItem(profile.get("name", "대상"), profile)
        idx = combo.findText(current_name)
        if idx >= 0:
            combo.setCurrentIndex(idx)
        combo.blockSignals(False)
        self._on_codex_target_selected()

    def _on_codex_target_selected(self) -> None:
        profile = self._selected_codex_target_profile()
        self._populate_codex_target_fields(profile)
        self._apply_codex_target_to_request()

    def _apply_codex_target_to_request(self) -> None:
        target_input = getattr(self, "codex_request_target_input", None)
        if target_input is None:
            return
        profile = self._codex_target_from_fields()
        target_text = (
            profile.get("path")
            or "생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리"
        )
        if target_input.text() != target_text:
            target_input.setText(target_text)
        self._schedule_codex_codegen_previews()

    def _codex_target_profile_from_fav_item(
        self, item: Optional[QTreeWidgetItem]
    ) -> Dict[str, str]:
        if item is None:
            return {}

        try:
            node_type = item.data(0, ROLE_TYPE)
        except Exception:
            return {}

        if node_type not in ("section", "notebook"):
            return {}

        def _clean(value: Any) -> str:
            if value is None:
                return ""
            return str(value).strip()

        try:
            payload = item.data(0, ROLE_DATA) or {}
        except Exception:
            payload = {}
        if not isinstance(payload, dict):
            payload = {}

        target = payload.get("target") or {}
        if not isinstance(target, dict):
            target = {}

        display_name = _clean(item.text(0))
        notebook = _clean(
            target.get("notebook")
            or target.get("notebook_text")
            or target.get("notebookName")
        )
        section_group = _clean(
            target.get("section_group")
            or target.get("section_group_text")
            or target.get("sectionGroup")
            or target.get("sectionGroupName")
        )
        section = _clean(
            target.get("section")
            or target.get("section_text")
            or target.get("sectionName")
        )

        ancestors: List[tuple[str, str]] = []
        parent = item.parent()
        while parent is not None:
            try:
                parent_type = _clean(parent.data(0, ROLE_TYPE))
                parent_name = _clean(parent.text(0))
            except Exception:
                parent_type = ""
                parent_name = ""
            if parent_name:
                ancestors.append((parent_type, parent_name))
            parent = parent.parent()
        ancestors.reverse()

        if node_type == "notebook":
            notebook = notebook or display_name
            section_group = ""
            section = ""
        else:
            section = section or display_name
            if not notebook:
                for ancestor_type, ancestor_name in ancestors:
                    if ancestor_type == "notebook":
                        notebook = ancestor_name
                        break
            if not section_group:
                group_names = [
                    ancestor_name
                    for ancestor_type, ancestor_name in ancestors
                    if ancestor_type == "group"
                ]
                if group_names:
                    section_group = group_names[-1]

        path = _clean(target.get("path") or target.get("hierarchy_path"))
        if not path:
            path_parts: List[str] = []
            for candidate in (notebook, section_group, section):
                if candidate and candidate not in path_parts:
                    path_parts.append(candidate)
            if not path_parts:
                path_parts = [name for _, name in ancestors]
                if display_name:
                    path_parts.append(display_name)
            path = " > ".join(path_parts)

        if node_type == "notebook":
            name = notebook or display_name or path or "전자필기장"
        else:
            name_parts = [part for part in (notebook, section) if part]
            name = " - ".join(name_parts) or display_name or path or "섹션"

        section_group_id = _clean(
            target.get("section_group_id")
            or target.get("sectionGroupId")
            or target.get("section_group_id_text")
        )
        section_id = _clean(target.get("section_id") or target.get("sectionId"))
        if node_type == "section" and not section_id:
            section_id = _clean(target.get("id"))

        return {
            "name": name,
            "path": path,
            "notebook": notebook,
            "section_group": section_group,
            "section": section,
            "section_group_id": section_group_id,
            "section_id": section_id,
        }

    def _sync_codex_target_from_fav_item(
        self,
        item: Optional[QTreeWidgetItem],
        *,
        switch_to_codex: bool = False,
    ) -> bool:
        if not switch_to_codex and not any(
            getattr(self, attr, None)
            for attr in (
                "codex_target_name_input",
                "codex_request_target_input",
                "codex_status_summary_preview",
            )
        ):
            return True

        target_input = getattr(self, "codex_request_target_input", None)
        try:
            item_text = item.text(0) if item is not None else ""
        except Exception:
            item_text = ""
        if not switch_to_codex:
            cached = getattr(self, "_last_codex_fav_target_sync", None)
            if (
                isinstance(cached, tuple)
                and len(cached) == 4
                and cached[0] == id(item)
                and cached[1] == item_text
                and (target_input is None or target_input.text() == cached[2])
                and self._codex_target_from_fields() == cached[3]
            ):
                return True

        profile = self._codex_target_profile_from_fav_item(item)
        if not profile:
            return False

        target_text = (
            profile.get("path")
            or "생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리"
        )
        cache_key = (id(item), item_text, target_text, profile)
        request_matches = target_input is None or target_input.text() == target_text
        if (
            not switch_to_codex
            and self._codex_target_from_fields() == profile
            and request_matches
        ):
            self._last_codex_fav_target_sync = cache_key
            return True

        self._populate_codex_target_fields(profile)
        self._apply_codex_target_to_request()
        try:
            self._update_codex_status_summary()
        except Exception:
            pass

        self._last_codex_fav_target_sync = cache_key
        if switch_to_codex:
            self._scroll_codex_to_widget("codex_target_group_widget")

        try:
            self.connection_status_label.setText(
                f"코덱스 작업 위치 반영: {profile.get('path') or profile.get('name')}"
            )
        except Exception:
            pass
        return True

    def _sync_codex_target_from_current_fav_item(self) -> None:
        try:
            item = self.fav_tree.currentItem()
        except Exception:
            item = None
        self._sync_codex_target_from_fav_item(item)

    def _save_codex_target_profile(self) -> None:
        profile = self._codex_target_from_fields()
        targets = self._load_codex_targets()
        replaced = False
        for i, old in enumerate(targets):
            if old.get("name") == profile.get("name"):
                targets[i] = profile
                replaced = True
                break
        if not replaced:
            targets.append(profile)
        try:
            self._write_codex_targets(targets)
            self._refresh_codex_target_combo(profile.get("name"))
            try:
                self.connection_status_label.setText(
                    f"코덱스 작업 위치 저장 완료: {profile.get('name')}"
                )
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "작업 위치 저장 실패", str(e))

    def _new_codex_target_profile(self) -> None:
        combo = getattr(self, "codex_target_combo", None)
        if combo is not None:
            combo.blockSignals(True)
            combo.setCurrentIndex(-1)
            combo.blockSignals(False)
        self._populate_codex_target_fields(
            {
                "name": "새 대상",
                "path": "",
                "notebook": "",
                "section_group": "",
                "section": "",
                "section_group_id": "",
                "section_id": "",
            }
        )
        self._apply_codex_target_to_request()

_publish_context(globals())

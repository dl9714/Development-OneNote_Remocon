# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin33:

    def _nodes_contain_notebook(self, nodes: Any) -> bool:
        if not isinstance(nodes, list):
            return False
        stack = list(nodes)
        while stack:
            node = stack.pop()
            if not isinstance(node, dict):
                continue
            if node.get("type") == "notebook":
                return True
            children = node.get("children")
            if isinstance(children, list):
                stack.extend(children)
            data = node.get("data")
            if isinstance(data, list):
                stack.extend(data)
        return False

    def _collect_classified_aggregate_notebook_keys(self) -> Set[str]:
        if getattr(self, "_aggregate_classified_keys_cache_valid", False):
            return set(getattr(self, "_aggregate_classified_keys_cache", set()))

        keys: Set[str] = set()

        def _walk_fav_nodes(nodes: Any) -> None:
            if not isinstance(nodes, list):
                return
            for node in nodes:
                if not isinstance(node, dict):
                    continue
                if node.get("type") == "notebook":
                    keys.update(self._aggregate_notebook_keys_from_node(node))
                children = node.get("children")
                if isinstance(children, list):
                    _walk_fav_nodes(children)

        def _walk_buffers(nodes: Any) -> None:
            if not isinstance(nodes, list):
                return
            for node in nodes:
                if not isinstance(node, dict):
                    continue
                if node.get("type") == "buffer":
                    if node.get("id") == AGG_BUFFER_ID:
                        continue
                    _walk_fav_nodes(node.get("data") or [])
                elif node.get("type") == "group":
                    _walk_buffers(node.get("children") or [])

        _walk_buffers(self.settings.get("favorites_buffers", []))
        self._aggregate_classified_keys_cache = set(keys)
        self._aggregate_classified_keys_cache_valid = True
        return keys

    def _build_aggregate_categorized_display_nodes(
        self, source_nodes: Any
    ) -> List[Dict[str, Any]]:
        def _notebook_is_open(node: Dict[str, Any]) -> bool:
            target = node.get("target") or {}
            return bool(node.get("is_open") or node.get("open") or target.get("is_open"))

        def _sort_open_first_key(node: Dict[str, Any]) -> tuple:
            return (
                0 if _notebook_is_open(node) else 1,
                _name_sort_key((node or {}).get("name", "")),
            )

        def _group_label(base_name: str, children: List[Dict[str, Any]]) -> str:
            if not children:
                return f"{base_name} (열림 0/0)"
            open_count = sum(1 for child in children if _notebook_is_open(child))
            return f"{base_name} (열림 {open_count}/{len(children)})"

        source_id = id(source_nodes)
        if (
            source_id == getattr(self, "_aggregate_display_cache_source_id", 0)
            and getattr(self, "_aggregate_display_cache_kind", None) == "categorized"
            and isinstance(getattr(self, "_aggregate_display_cache", None), list)
        ):
            return self._aggregate_display_cache

        classified_keys = self._collect_classified_aggregate_notebook_keys()
        cache_sig = ("categorized", tuple(sorted(classified_keys)))

        notebooks = self._collect_notebook_nodes_from_nodes(source_nodes, include_keys=True)
        if not notebooks:
            notebooks = self._collect_notebook_nodes_from_nodes(
                _collect_all_sections_dedup(self.settings),
                include_keys=True,
            )

        unclassified: List[Dict[str, Any]] = []
        classified: List[Dict[str, Any]] = []
        for notebook, notebook_keys in notebooks:
            if notebook_keys and notebook_keys & classified_keys:
                classified.append(notebook)
            else:
                unclassified.append(notebook)

        try:
            unclassified.sort(key=_sort_open_first_key)
            classified.sort(key=_sort_open_first_key)
        except Exception:
            pass

        result = [
            {
                "type": "group",
                "id": AGG_UNCLASSIFIED_GROUP_ID,
                "name": _group_label(AGG_UNCLASSIFIED_GROUP_NAME, unclassified),
                "children": unclassified,
            },
            {
                "type": "group",
                "id": AGG_CLASSIFIED_GROUP_ID,
                "name": _group_label(AGG_CLASSIFIED_GROUP_NAME, classified),
                "children": classified,
            },
        ]
        self._aggregate_display_cache_sig = cache_sig
        self._aggregate_display_cache_kind = "categorized"
        self._aggregate_display_cache = result
        self._aggregate_display_cache_source_id = source_id
        return result

    def _aggregate_classification_signature_from_nodes(self, nodes: Any) -> tuple:
        if not isinstance(nodes, list):
            return tuple()

        def _collect_keys(children: Any) -> List[str]:
            keys: List[str] = []
            if not isinstance(children, list):
                return keys
            stack = list(children)
            while stack:
                child = stack.pop()
                if not isinstance(child, dict):
                    continue
                keys.extend(self._aggregate_notebook_keys_from_node(child))
                nested = child.get("children")
                if isinstance(nested, list):
                    stack.extend(nested)
            return keys

        groups: Dict[str, List[str]] = {
            AGG_UNCLASSIFIED_GROUP_ID: [],
            AGG_CLASSIFIED_GROUP_ID: [],
        }
        fallback_keys: List[str] = []

        for node in nodes:
            if not isinstance(node, dict):
                continue
            node_type = node.get("type")
            node_id = node.get("id")
            if node_type == "group" and node_id in groups:
                groups[node_id].extend(_collect_keys(node.get("children") or []))
            elif node_type == "notebook":
                fallback_keys.extend(self._aggregate_notebook_keys_from_node(node))

        if fallback_keys and not (groups[AGG_UNCLASSIFIED_GROUP_ID] or groups[AGG_CLASSIFIED_GROUP_ID]):
            return (("flat", tuple(sorted(set(fallback_keys)))),)

        return (
            (
                AGG_UNCLASSIFIED_GROUP_ID,
                tuple(sorted(set(groups[AGG_UNCLASSIFIED_GROUP_ID]))),
            ),
            (
                AGG_CLASSIFIED_GROUP_ID,
                tuple(sorted(set(groups[AGG_CLASSIFIED_GROUP_ID]))),
            ),
        )

    def _serialize_current_module_tree(self) -> List[Dict[str, Any]]:
        nodes: List[Dict[str, Any]] = []
        try:
            root = self.fav_tree.invisibleRootItem()
            for i in range(root.childCount()):
                nodes.append(self._serialize_fav_item(root.child(i)))
        except Exception:
            pass
        return nodes

    def _aggregate_source_nodes_for_fast_classification(
        self, current_nodes: Optional[List[Dict[str, Any]]] = None
    ) -> List[Dict[str, Any]]:
        if getattr(self, "active_buffer_id", None) == AGG_BUFFER_ID:
            if current_nodes is None:
                current_nodes = self._serialize_current_module_tree()
            if self._nodes_contain_notebook(current_nodes):
                return current_nodes

        agg_node = _find_buffer_node_by_id(
            self.settings.get("favorites_buffers", []),
            AGG_BUFFER_ID,
        )
        saved = (agg_node or {}).get("data", []) if isinstance(agg_node, dict) else []
        if self._nodes_contain_notebook(saved):
            return saved
        return _collect_all_sections_dedup(self.settings)

    def _persist_active_aggregate_data(self, data: List[Dict[str, Any]]) -> None:
        new_sig = self._calc_nodes_signature(data)
        if getattr(self, "active_buffer_id", None) == AGG_BUFFER_ID and self.active_buffer_node is not None:
            self.active_buffer_node["data"] = data
        if getattr(self, "active_buffer_id", None) == AGG_BUFFER_ID and self.active_buffer_item is not None:
            payload = self.active_buffer_item.data(0, ROLE_DATA) or {}
            payload["data"] = data
            self.active_buffer_item.setData(0, ROLE_DATA, payload)
        settings_node = _find_buffer_node_by_id(
            self.settings.get("favorites_buffers", []),
            AGG_BUFFER_ID,
        )
        if isinstance(settings_node, dict):
            old_sig = self._calc_nodes_signature(settings_node.get("data", []))
            settings_node["data"] = data
            self._active_buffer_settings_node = settings_node
            if not new_sig or new_sig != old_sig:
                self._open_all_candidate_count_dirty = True
                self._save_settings_to_file()

    def _refresh_active_aggregate_classification_from_saved_data(
        self,
        *,
        current_nodes: Optional[List[Dict[str, Any]]] = None,
        persist: bool = True,
        show_status: bool = False,
    ) -> bool:
        if getattr(self, "active_buffer_id", None) != AGG_BUFFER_ID:
            if show_status:
                QMessageBox.information(self, "안내", "종합 버퍼에서만 사용할 수 있습니다.")
            return False
        if getattr(self, "_aggregate_reclassify_in_progress", False):
            return False

        if current_nodes is None:
            current_nodes = self._serialize_current_module_tree()
        source_nodes = self._aggregate_source_nodes_for_fast_classification(current_nodes)
        categorized = self._build_aggregate_categorized_display_nodes(source_nodes)
        current_sig = self._aggregate_classification_signature_from_nodes(current_nodes)
        next_sig = self._aggregate_classification_signature_from_nodes(categorized)

        if current_sig and current_sig == next_sig:
            if show_status:
                try:
                    self.connection_status_label.setText(
                        "종합 분류 상태가 이미 최신입니다."
                    )
                except Exception:
                    pass
            return False

        self._aggregate_reclassify_in_progress = True
        try:
            self._load_favorites_into_center_tree(categorized)
            self._fav_reset_undo_context_from_data(
                categorized,
                reason="aggregate_fast_reclassify",
            )
            if persist:
                self._persist_active_aggregate_data(categorized)
            if show_status:
                try:
                    unclassified_count = len(categorized[0].get("children") or [])
                    classified_count = len(categorized[1].get("children") or [])
                    self.connection_status_label.setText(
                        "종합 분류 새로고침 완료: "
                        f"분류 안 됨 {unclassified_count}개, 분류됨 {classified_count}개"
                    )
                except Exception:
                    pass
            return True
        finally:
            self._aggregate_reclassify_in_progress = False

    def _refresh_active_aggregate_classification_action(self) -> None:
        self._refresh_active_aggregate_classification_from_saved_data(
            persist=True,
            show_status=True,
        )

    def _collect_open_all_notebook_candidates(self) -> List[Dict[str, Any]]:
        def _coerce_last_accessed_at(value: Any) -> int:
            try:
                return max(0, int(value or 0))
            except Exception:
                return 0

        self._last_open_all_candidate_scope = "SETTINGS_FALLBACK"
        seen: Set[str] = set()
        records: List[Dict[str, Any]] = []

        def _append_record(name: str, target: Dict[str, Any], source: str) -> None:
            clean_name = str(name or "").strip()
            key = _normalize_notebook_name_key(clean_name)
            if not key or key in seen:
                return
            seen.add(key)
            records.append(
                {
                    "name": clean_name,
                    "id": str(target.get("notebook_id") or "").strip(),
                    "path": str(target.get("path") or "").strip(),
                    "url": str(target.get("url") or target.get("notebook_url") or "").strip(),
                    "last_accessed_at": _coerce_last_accessed_at(
                        target.get("last_accessed_at")
                    ),
                    "source": source,
                }
            )

        try:
            source_nodes = self._aggregate_source_nodes_for_fast_classification()
            categorized = self._build_aggregate_categorized_display_nodes(source_nodes)
            aggregate_nodes = self._collect_notebook_nodes_from_nodes(categorized)
            for node in aggregate_nodes:
                if bool(node.get("is_open") or (node.get("target") or {}).get("is_open")):
                    continue
                target = dict(node.get("target") or {})
                notebook_name = (
                    str(target.get("notebook_text") or "").strip()
                    or str(node.get("name") or "").strip()
                )
                if notebook_name:
                    _append_record(notebook_name, target, "AGG_UNCHECKED")
            if aggregate_nodes:
                self._last_open_all_candidate_scope = "AGG_UNCHECKED"
                return records
        except Exception as e:
            print(f"[WARN][OPEN_ALL][CANDIDATES][AGG] {e}")

        if records:
            return records
        self._last_open_all_candidate_scope = "SETTINGS_FALLBACK"
        return _collect_known_notebook_name_records(self.settings)

    def _build_aggregate_buffer(self):
        """모든 섹션을 수집하여 종합 버퍼 데이터를 생성합니다."""
        if getattr(self, "_aggregate_cache_valid", False):
            return self._aggregate_cache
        data = _collect_all_sections_dedup(self.settings)
        self._aggregate_cache = data
        self._aggregate_cache_valid = True
        return data

_publish_context(globals())

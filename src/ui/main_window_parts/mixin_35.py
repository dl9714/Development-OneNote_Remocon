# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin35:

    def _project_search_text_from_fav_tree(self) -> List[str]:
        parts = []
        try:
            root = self.fav_tree.invisibleRootItem()
            stack = [root.child(i) for i in range(root.childCount() - 1, -1, -1)]
            while stack and len(parts) < 1200:
                item = stack.pop()
                text = item.text(0)
                if text:
                    parts.append(text)
                payload = item.data(0, ROLE_DATA) or {}
                parts.extend(self._project_search_text_from_target(payload.get("target")))
                for i in range(item.childCount() - 1, -1, -1):
                    stack.append(item.child(i))
        except Exception:
            pass
        return parts

    def _project_buffer_search_key_for_item(self, item: QTreeWidgetItem) -> str:
        if item is None:
            return ""
        parts = [item.text(0)]
        payload = item.data(0, ROLE_DATA) or {}
        item_id = payload.get("id")
        if item_id and item_id == getattr(self, "active_buffer_id", None):
            parts.extend(self._project_search_text_from_fav_tree())
        else:
            parts.extend(self._project_search_text_from_nodes(payload.get("data")))
        return _normalize_project_search_key(" ".join(p for p in parts if p))

    def _expand_buffer_groups_always(self, *, reason: str = "") -> None:
        changed = False
        was_updates_enabled = True
        try:
            root = self.buffer_tree.invisibleRootItem()
            stack = [root.child(i) for i in range(root.childCount() - 1, -1, -1)]
            expanded = 0
            was_updates_enabled = self.buffer_tree.updatesEnabled()
            self.buffer_tree.setUpdatesEnabled(False)
            while stack:
                it = stack.pop()
                if it.childCount() > 0:
                    if not it.isExpanded():
                        self.buffer_tree.expandItem(it)
                        expanded += 1
                        changed = True
                    for j in range(it.childCount()):
                        stack.append(it.child(j))
            tag = f" reason={reason}" if reason else ""
            self._dbg_hot(f"[DBG][BUF][EXPAND_GROUPS]{tag} expanded={expanded}")
        except Exception as e:
            try:
                tag = f" reason={reason}" if reason else ""
                self._dbg_hot(f"[DBG][BUF][EXPAND_GROUPS][FAIL]{tag} {e}")
            except Exception:
                pass
        finally:
            try:
                self.buffer_tree.setUpdatesEnabled(was_updates_enabled)
                if changed and was_updates_enabled:
                    self.buffer_tree.viewport().update()
            except Exception:
                pass

    def _rebuild_buffer_item_index(self) -> None:
        self._buffer_item_index = {}
        self._first_buffer_item = None
        self._buffer_search_index = []
        self._buffer_search_last_match_records = []
        try:
            root = self.buffer_tree.invisibleRootItem()
            stack = [root.child(i) for i in range(root.childCount() - 1, -1, -1)]
            while stack:
                item = stack.pop()
                payload = item.data(0, ROLE_DATA) or {}
                item_id = payload.get("id")
                if item_id:
                    self._buffer_item_index[item_id] = item
                    if (
                        item.data(0, ROLE_TYPE) == "buffer"
                        and self._first_buffer_item is None
                    ):
                        self._first_buffer_item = item
                if item.data(0, ROLE_TYPE) == "buffer":
                    search_key = self._project_buffer_search_key_for_item(item)
                    if search_key:
                        parents = []
                        parent = item.parent()
                        while parent is not None:
                            parents.append(parent)
                            parent = parent.parent()
                        self._buffer_search_index.append(
                            {"item": item, "key": search_key, "parents": tuple(parents)}
                        )
                for i in range(item.childCount() - 1, -1, -1):
                    stack.append(item.child(i))
        except Exception as e:
            self._dbg_hot(f"[DBG][BUF][INDEX][FAIL] {e}")

    def _rebuild_module_search_index(self) -> None:
        self._module_search_index = []
        self._module_search_last_match_records = []
        try:
            root = self.fav_tree.invisibleRootItem()
            stack = [root.child(i) for i in range(root.childCount() - 1, -1, -1)]
            while stack:
                item = stack.pop()
                search_key = _normalize_project_search_key(item.text(0))
                if search_key:
                    parents = []
                    parent = item.parent()
                    while parent is not None:
                        parents.append(parent)
                        parent = parent.parent()
                    self._module_search_index.append(
                        {"item": item, "key": search_key, "parents": tuple(parents)}
                    )
                for i in range(item.childCount() - 1, -1, -1):
                    stack.append(item.child(i))
        except Exception as e:
            self._dbg_hot(f"[DBG][MOD][SEARCH_INDEX][FAIL] {e}")

    def _set_buffer_search_highlight(self, item: QTreeWidgetItem, enabled: bool) -> None:
        if item is None:
            return
        try:
            if enabled:
                item.setBackground(0, self._buffer_search_highlight_bg)
                item.setForeground(0, self._buffer_search_highlight_fg)
            else:
                item.setBackground(0, self._buffer_search_clear_bg)
                item.setForeground(0, self._buffer_search_clear_fg)
        except Exception:
            pass

    def _set_module_search_highlight(self, item: QTreeWidgetItem, enabled: bool) -> None:
        self._set_buffer_search_highlight(item, enabled)

    def _clear_project_buffer_search_highlights(self) -> None:
        was_updates_enabled = True
        fav_updates_enabled = True
        changed = False
        module_changed = False
        try:
            highlighted = list(self._buffer_search_highlighted_by_id.values())
            module_highlighted = list(self._module_search_highlighted_by_id.values())
            was_updates_enabled = self.buffer_tree.updatesEnabled()
            if highlighted:
                self.buffer_tree.setUpdatesEnabled(False)
                for item in highlighted:
                    self._set_buffer_search_highlight(item, False)
                changed = True
            fav_updates_enabled = self.fav_tree.updatesEnabled()
            if module_highlighted:
                self.fav_tree.setUpdatesEnabled(False)
                for item in module_highlighted:
                    self._set_module_search_highlight(item, False)
                module_changed = True
            self._buffer_search_highlighted_by_id = {}
            self._buffer_search_last_match_records = []
            self._module_search_highlighted_by_id = {}
            self._module_search_last_match_records = []
            self._buffer_search_match_count = 0
            self._module_search_match_count = 0
            self._buffer_search_last_applied_key = ""
            self._buffer_search_last_first_match_id = 0
            self._module_search_last_first_match_id = 0
            self._buffer_search_pending_key = ""
        except Exception:
            pass
        finally:
            try:
                self.buffer_tree.setUpdatesEnabled(was_updates_enabled)
                if changed and was_updates_enabled:
                    self.buffer_tree.viewport().update()
                self.fav_tree.setUpdatesEnabled(fav_updates_enabled)
                if module_changed and fav_updates_enabled:
                    self.fav_tree.viewport().update()
            except Exception:
                pass

    def _schedule_project_buffer_search_highlight(self, text: str = "") -> None:
        self._buffer_search_pending_text = text or ""
        self._buffer_search_pending_key = _normalize_project_search_key(text)
        if not self._buffer_search_pending_key:
            self._buffer_search_timer.stop()
            self._clear_project_buffer_search_highlights()
            return
        self._buffer_search_timer.start(35)

    def _apply_project_search_to_tree(
        self,
        *,
        tree: QTreeWidget,
        records: List[Dict[str, Any]],
        previous_records: List[Dict[str, Any]],
        previous_highlights: Dict[int, QTreeWidgetItem],
        query: str,
        previous_query: str,
        set_highlight: Callable[[QTreeWidgetItem, bool], None],
        previous_first_match_id: int,
        scroll: bool,
    ) -> Dict[str, Any]:
        records_to_scan = records
        if previous_query and query.startswith(previous_query):
            records_to_scan = previous_records

        match_count = 0
        matched_records = []
        match_by_id: Dict[int, QTreeWidgetItem] = {}
        first_match = None
        changed_highlights = False
        was_updates_enabled = True
        try:
            was_updates_enabled = tree.updatesEnabled()
            for record in records_to_scan:
                item = record.get("item")
                item_key = record.get("key") or ""
                if item is not None and item_key and query in item_key:
                    match_count += 1
                    matched_records.append(record)
                    if first_match is None:
                        first_match = item
                    match_by_id[id(item)] = item

            prev_ids = set(previous_highlights)
            next_ids = set(match_by_id)
            removed_ids = prev_ids - next_ids
            added_ids = next_ids - prev_ids
            changed_highlights = bool(removed_ids or added_ids)

            if changed_highlights:
                tree.setUpdatesEnabled(False)
                for item_id in removed_ids:
                    set_highlight(previous_highlights.get(item_id), False)
                expanded_parent_ids = set()
                for record in matched_records:
                    item = record.get("item")
                    if item is None or id(item) not in added_ids:
                        continue
                    set_highlight(item, True)
                    for parent in record.get("parents") or ():
                        parent_id = id(parent)
                        if parent_id in expanded_parent_ids:
                            continue
                        expanded_parent_ids.add(parent_id)
                        if not parent.isExpanded():
                            tree.expandItem(parent)
        finally:
            try:
                tree.setUpdatesEnabled(was_updates_enabled)
                if was_updates_enabled and changed_highlights:
                    tree.viewport().update()
            except Exception:
                pass

        first_match_id = id(first_match) if first_match is not None else 0
        if (
            scroll
            and first_match is not None
            and first_match_id != previous_first_match_id
        ):
            try:
                tree.scrollToItem(
                    first_match,
                    QAbstractItemView.ScrollHint.EnsureVisible,
                )
            except Exception:
                pass

        return {
            "matched_records": matched_records,
            "match_by_id": match_by_id,
            "match_count": match_count,
            "first_match_id": first_match_id,
            "changed": changed_highlights,
        }

    def _highlight_project_buffers_from_module_search(
        self,
        text: str = "",
        *,
        update_status: bool = True,
        scroll: bool = True,
        precomputed_query: Optional[str] = None,
    ) -> None:
        query = precomputed_query if precomputed_query is not None else _normalize_project_search_key(text)
        if not query:
            self._clear_project_buffer_search_highlights()
            return

        if not self._buffer_search_index:
            self._rebuild_buffer_item_index()
        if not self._module_search_index:
            self._rebuild_module_search_index()

        previous_query = self._buffer_search_last_applied_key
        if query == previous_query:
            return
        self._buffer_search_last_applied_key = query

        buffer_result = self._apply_project_search_to_tree(
            tree=self.buffer_tree,
            records=self._buffer_search_index,
            previous_records=getattr(self, "_buffer_search_last_match_records", []),
            previous_highlights=self._buffer_search_highlighted_by_id,
            query=query,
            previous_query=previous_query,
            set_highlight=self._set_buffer_search_highlight,
            previous_first_match_id=self._buffer_search_last_first_match_id,
            scroll=scroll,
        )
        module_result = self._apply_project_search_to_tree(
            tree=self.fav_tree,
            records=self._module_search_index,
            previous_records=getattr(self, "_module_search_last_match_records", []),
            previous_highlights=self._module_search_highlighted_by_id,
            query=query,
            previous_query=previous_query,
            set_highlight=self._set_module_search_highlight,
            previous_first_match_id=self._module_search_last_first_match_id,
            scroll=scroll,
        )

        self._buffer_search_last_match_records = buffer_result["matched_records"]
        self._buffer_search_highlighted_by_id = buffer_result["match_by_id"]
        self._buffer_search_match_count = buffer_result["match_count"]
        self._buffer_search_last_first_match_id = buffer_result["first_match_id"]
        self._module_search_last_match_records = module_result["matched_records"]
        self._module_search_highlighted_by_id = module_result["match_by_id"]
        self._module_search_match_count = module_result["match_count"]
        self._module_search_last_first_match_id = module_result["first_match_id"]

        if update_status:
            try:
                self.connection_status_label.setText(
                    _project_search_status_text(
                        text,
                        self._buffer_search_match_count,
                        self._module_search_match_count,
                    )
                )
            except Exception:
                pass

_publish_context(globals())

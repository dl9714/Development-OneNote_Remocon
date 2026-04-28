# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin32:

    def _open_all_notebooks_from_connected_onenote(self):
        if (
            self._open_all_notebooks_worker is not None
            and self._open_all_notebooks_worker.isRunning()
        ):
            try:
                self._open_all_notebooks_worker.requestInterruption()
            except Exception:
                pass
            self.update_status_and_ui(
                f"{_open_unchecked_notebooks_button_label()} 중지 요청됨...",
                True,
            )
            return

        win = getattr(self, "onenote_window", None)
        hwnd = getattr(win, "handle", None) if win is not None else None
        if callable(hwnd):
            try:
                hwnd = hwnd()
            except Exception:
                hwnd = None
        if not hwnd:
            self.update_status_and_ui(
                "OneNote 창이 연결되지 않았습니다. 먼저 OneNote 창을 연결하세요.",
                False,
            )
            return

        ensure_pywinauto()
        if not IS_MACOS and not _pwa_ready:
            self.update_status_and_ui("오류: 자동화 모듈이 로드되지 않았습니다.", True)
            return

        if not IS_MACOS:
            try:
                win.set_focus()
            except Exception:
                pass

        if IS_MACOS:
            saved_sig = (
                self.settings.get("connection_signature")
                if isinstance(self.settings.get("connection_signature"), dict)
                else None
            )
            sig = build_window_signature_quick(win, saved_sig)
        else:
            sig = build_window_signature(win)
        if not sig:
            self.update_status_and_ui("오류: OneNote 창 정보를 읽지 못했습니다.", True)
            return

        notebook_candidates = self._collect_open_all_notebook_candidates()
        candidate_scope = str(
            getattr(self, "_last_open_all_candidate_scope", "SETTINGS_FALLBACK") or ""
        )
        print(
            "[DBG][OPEN_ALL][CANDIDATES]",
            f"scope={candidate_scope}",
            f"count={len(notebook_candidates)}",
            f"sample={[str((record or {}).get('name') or '').strip() for record in notebook_candidates[:8]]!r}",
        )
        self._open_all_candidate_scope = candidate_scope
        self._open_all_candidate_count = (
            len(notebook_candidates) if candidate_scope == "AGG_UNCHECKED" else None
        )
        self._open_all_candidate_stats = (
            self._summarize_open_all_notebook_candidates(notebook_candidates)
            if candidate_scope == "AGG_UNCHECKED"
            else {}
        )
        self._open_all_candidate_count_dirty = False
        if candidate_scope == "AGG_UNCHECKED" and not notebook_candidates:
            self.update_status_and_ui(
                "체크 없는 전자필기장 열 대상이 없습니다.",
                False,
            )
            return
        worker = OpenAllNotebooksWorker(
            sig,
            notebook_candidates=notebook_candidates,
            candidate_scope=candidate_scope,
            parent=self,
        )
        self._open_all_notebooks_worker = worker
        if candidate_scope == "AGG_UNCHECKED":
            stats_text = self._format_open_all_candidate_stats_for_tip(
                getattr(self, "_open_all_candidate_stats", {})
            )
            suffix = f" ({stats_text})" if stats_text else ""
            self.update_status_and_ui(
                f"{_open_unchecked_notebooks_button_label(len(notebook_candidates))} 준비 중...{suffix}",
                True,
            )
        else:
            self.update_status_and_ui("실제 OneNote 전체 열기 준비 중...", True)

        def _on_progress(message: str):
            if self._open_all_notebooks_worker is worker:
                self.connection_status_label.setText(message)

        def _on_done(result: Dict[str, Any]):
            if self._open_all_notebooks_worker is not worker:
                return

            self._open_all_notebooks_worker = None
            self._open_all_candidate_count_dirty = True
            try:
                worker.deleteLater()
            except Exception:
                pass

            connected = self._apply_connected_window_info(result.get("window_info"))
            is_connected = connected or bool(getattr(self, "onenote_window", None))
            opened_count = int(result.get("opened_count") or 0)
            verified_open_count = int(result.get("verified_open_count") or 0)
            candidate_total = int(result.get("candidate_total") or 0)
            pending_count = int(result.get("pending_count") or 0)
            already_open_count = int(result.get("already_open_count") or 0)
            result_scope = str(result.get("candidate_scope") or candidate_scope or "")
            remaining = result.get("remaining_names") or []
            error = (result.get("error") or "").strip()

            should_refresh_open_notebooks = opened_count > 0 or bool(
                result.get("refresh_open_notebooks")
            )
            refresh_delay_ms = 1800 if IS_MACOS else 400

            if result.get("ok"):
                if opened_count > 0:
                    status = (
                        f"{_open_unchecked_notebooks_button_label()} 완료: {opened_count}개"
                        if result_scope == "AGG_UNCHECKED"
                        else f"실제 OneNote 전체 열기 완료: {opened_count}개"
                    )
                    if should_refresh_open_notebooks:
                        status += (
                            ", 종합 새로고침 예약"
                            if IS_MACOS
                            else ", 종합 새로고침"
                    )
                    self.update_status_and_ui(status, is_connected)
                elif (
                    result_scope == "AGG_UNCHECKED"
                    and candidate_total > 0
                    and pending_count == 0
                ):
                    already = already_open_count or verified_open_count
                    status = (
                        "체크 없는 전자필기장 확인 완료: "
                        f"후보 {candidate_total}개 중 열 대상 0개"
                    )
                    if already > 0:
                        status += f", 이미 열림 {already}개"
                    if should_refresh_open_notebooks:
                        status += ", 종합 새로고침 예약"
                    self.update_status_and_ui(status, is_connected)
                elif verified_open_count > 0:
                    self.update_status_and_ui(
                        (
                            (
                                f"{_open_unchecked_notebooks_button_label()} 확인 완료: "
                                if result_scope == "AGG_UNCHECKED"
                                else "실제 OneNote 전체 열기 확인 완료: "
                            )
                            + f"이미 열린 전자필기장 {verified_open_count}개"
                        ),
                        is_connected,
                    )
                else:
                    self.update_status_and_ui(
                        "열어야 할 전자필기장이 더 이상 없습니다.",
                        is_connected,
                    )
                if should_refresh_open_notebooks:
                    QTimer.singleShot(
                        refresh_delay_ms,
                        lambda: self._register_all_notebooks_from_current_onenote(
                            force=True
                        ),
                    )
                return

            if remaining:
                remain_preview = ", ".join(remaining[:3])
                if len(remaining) > 3:
                    remain_preview += " ..."
                suffix = f" 남은 후보: {remain_preview}"
            else:
                suffix = ""
            detail = error or (
                f"{_open_unchecked_notebooks_button_label()}에 실패했습니다."
                if result_scope == "AGG_UNCHECKED"
                else "실제 OneNote 전체 열기에 실패했습니다."
            )
            self.update_status_and_ui(
                f"{detail} (시도 {opened_count}개).{suffix}",
                is_connected,
            )
            if should_refresh_open_notebooks:
                QTimer.singleShot(
                    refresh_delay_ms,
                    lambda: self._register_all_notebooks_from_current_onenote(force=True),
                )

        worker.progress.connect(_on_progress)
        worker.done.connect(_on_done)
        worker.finished.connect(lambda: None)
        worker.start()

    def _search_and_select_section(self):
        """입력창의 텍스트로 섹션을 검색하고 선택 및 중앙 정렬합니다."""
        if not self._pre_action_check():
            return

        search_text = self.search_input.text().strip()
        if not search_text:
            self.update_status_and_ui("검색할 내용을 입력하세요.", True)
            return

        if not self.tree_control:
            self.tree_control = _find_tree_or_list(self.onenote_window)

        self.update_status_and_ui(f"'{search_text}' 섹션을 검색 중...", True)

        success = select_section_by_text(
            self.onenote_window, search_text, self.tree_control
        )

        if success:
            QTimer.singleShot(100, self.center_selected_item_action)
            self.update_status_and_ui(f"검색 성공: '{search_text}' 선택 완료.", True)
        else:
            self.update_status_and_ui(
                f"검색 실패: '{search_text}' 섹션을 찾을 수 없습니다.", True
            )

    def _calc_nodes_signature(self, obj):
        """리스트/딕트의 안정적인 시그니처를 계산합니다."""
        try:
            raw = json.dumps(obj, sort_keys=True, ensure_ascii=False)
            return hashlib.md5(raw.encode("utf-8")).hexdigest()
        except Exception:
            return None

    def _invalidate_aggregate_cache(self, *, invalidate_classified_keys: bool = True):
        """종합 버퍼 계산/표시 캐시를 무효화합니다."""
        self._aggregate_cache_valid = False
        self._aggregate_cache = []
        self._aggregate_display_cache_sig = None
        self._aggregate_display_cache = []
        self._aggregate_display_cache_kind = None
        self._aggregate_display_cache_source_id = 0
        if invalidate_classified_keys:
            self._aggregate_classified_keys_cache_valid = False
            self._aggregate_classified_keys_cache = set()
        self._open_all_candidate_count_dirty = True

    def _get_sorted_aggregate_display_nodes(self, nodes, *, kind: str):
        """
        종합 버퍼 표시용 정렬 결과를 캐시합니다.
        - kind='saved': 종합 버퍼에 직접 저장된 notebook/group 표시본
        - kind='built': 전체 버퍼에서 수집해 만든 fallback 표시본
        """
        source_id = id(nodes)
        if (
            source_id
            and source_id == getattr(self, "_aggregate_display_cache_source_id", 0)
            and kind == getattr(self, "_aggregate_display_cache_kind", None)
            and isinstance(getattr(self, "_aggregate_display_cache", None), list)
        ):
            return self._aggregate_display_cache
        sig = self._calc_nodes_signature(nodes)
        if (
            sig is not None
            and sig == getattr(self, "_aggregate_display_cache_sig", None)
            and kind == getattr(self, "_aggregate_display_cache_kind", None)
            and isinstance(getattr(self, "_aggregate_display_cache", None), list)
        ):
            return self._aggregate_display_cache
        data = self._sorted_copy_nodes_by_name(nodes)
        self._aggregate_display_cache_sig = sig
        self._aggregate_display_cache_kind = kind
        self._aggregate_display_cache = data
        self._aggregate_display_cache_source_id = source_id
        return data

    def _aggregate_notebook_key_from_node(self, node: Any) -> str:
        keys = self._aggregate_notebook_keys_from_node(node)
        if not keys:
            return ""
        id_keys = sorted(k for k in keys if k.startswith("id:"))
        if id_keys:
            return id_keys[0]
        return sorted(keys)[0]

    def _aggregate_notebook_keys_from_node(self, node: Any) -> Set[str]:
        if not isinstance(node, dict):
            return frozenset()
        target = node.get("target") or {}
        notebook_id = str(target.get("notebook_id") or "").strip()
        notebook_id_key = notebook_id.casefold()
        name = (
            str(target.get("notebook_text") or "").strip()
            or str(node.get("name") or "").strip()
        )
        cache_key = (notebook_id_key, name)
        cache = getattr(self, "_aggregate_notebook_keys_cache", None)
        if cache is None:
            cache = {}
            self._aggregate_notebook_keys_cache = cache
        cached = cache.get(cache_key)
        if cached is not None:
            return cached

        keys: Set[str] = set()
        if notebook_id_key:
            keys.add("id:" + notebook_id_key)
        name_key = _normalize_notebook_name_key(name)
        if name_key:
            keys.add("name:" + name_key)
        result = frozenset(keys)
        if len(cache) >= 4096:
            cache.clear()
        cache[cache_key] = result
        return result

    def _collect_notebook_nodes_from_nodes(self, nodes: Any, *, include_keys: bool = False) -> List[Any]:
        if not isinstance(nodes, list):
            return []
        found: List[Any] = []
        seen: Dict[str, int] = {}
        stack = list(reversed(nodes))
        while stack:
            node = stack.pop()
            if not isinstance(node, dict):
                continue
            if node.get("type") == "notebook":
                node_keys = self._aggregate_notebook_keys_from_node(node)
                if node_keys:
                    duplicate_index = None
                    for key in node_keys:
                        if key in seen:
                            duplicate_index = seen[key]
                            break
                    is_open = bool(
                        node.get("is_open")
                        or node.get("open")
                        or (node.get("target") or {}).get("is_open")
                    )
                    if duplicate_index is not None:
                        if IS_WINDOWS and is_open:
                            existing = (
                                found[duplicate_index][0]
                                if include_keys
                                else found[duplicate_index]
                            )
                            if isinstance(existing, dict):
                                existing["is_open"] = True
                                target = dict(existing.get("target") or {})
                                target["is_open"] = True
                                existing["target"] = target
                        continue
                    target = dict(node.get("target") or {})
                    sig = target.get("sig")
                    if isinstance(sig, dict):
                        target["sig"] = dict(sig)
                    record = {
                        "type": "notebook",
                        "id": node.get("id") or str(uuid.uuid4()),
                        "name": node.get("name") or "전자필기장",
                        "target": target,
                        "is_open": is_open,
                    }
                    found.append((record, node_keys) if include_keys else record)
                    record_index = len(found) - 1
                    for key in node_keys:
                        seen[key] = record_index
                continue
            children = node.get("children")
            if isinstance(children, list):
                stack.extend(reversed(children))
            data = node.get("data")
            if isinstance(data, list):
                stack.extend(reversed(data))
        try:
            found.sort(key=lambda n: _name_sort_key(((n[0] if include_keys else n) or {}).get("name", "")))
        except Exception:
            pass
        return found

_publish_context(globals())

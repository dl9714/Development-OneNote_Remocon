# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin41:

    def _register_all_notebooks_from_current_onenote(
        self,
        *,
        force: bool = False,
        _prefetched_records: Optional[List[Dict[str, str]]] = None,
        _prefetched_source: str = "",
        _prefetched_sig: Optional[Dict[str, Any]] = None,
        _prefetched_error: str = "",
    ):
        """종합 버퍼를 OneNote의 열린 전자필기장 목록으로 새로고침하고 분류까지 갱신합니다."""
        started_at = time.perf_counter()
        refresh_button = getattr(self, "btn_register_all_notebooks", None)
        if refresh_button is not None:
            refresh_button.setEnabled(False)

        try:
            cur_item = self.buffer_tree.currentItem()
            cur_payload = cur_item.data(0, ROLE_DATA) if cur_item else {}
            is_agg = (
                getattr(self, "active_buffer_id", None) == AGG_BUFFER_ID
                or bool(isinstance(cur_payload, dict) and cur_payload.get("id") == AGG_BUFFER_ID)
            )
            if not is_agg and not force:
                QMessageBox.information(self, "안내", "이 기능은 '종합' 버퍼에서만 사용할 수 있습니다.")
                return

            onenote_window = getattr(self, "onenote_window", None)
            hwnd = None
            try:
                if onenote_window is not None:
                    hwnd = getattr(onenote_window, "handle", None)
                    if callable(hwnd):
                        hwnd = onenote_window.handle()
            except Exception:
                hwnd = None
            if not hwnd:
                self.update_status_and_ui("OneNote 창이 연결되지 않았습니다. 먼저 OneNote 창 연결/선택을 해주세요.", False)
                return

            if IS_MACOS and _prefetched_records is None:
                worker = getattr(self, "_open_notebooks_refresh_worker", None)
                if worker is not None and worker.isRunning():
                    self.update_status_and_ui("종합 새로고침이 이미 실행 중입니다.", True)
                    return

                saved_sig = (
                    self.settings.get("connection_signature")
                    if isinstance(self.settings.get("connection_signature"), dict)
                    else None
                )
                sig = build_window_signature_quick(onenote_window, saved_sig)
                worker = OpenNotebookRecordsWorker(sig, self)
                self._open_notebooks_refresh_worker = worker
                self.update_status_and_ui("종합 새로고침 중... 열린 전자필기장 목록 수집", True)

                def _on_done(result: Dict[str, Any]):
                    if self._open_notebooks_refresh_worker is not worker:
                        return
                    self._open_notebooks_refresh_worker = None
                    try:
                        worker.deleteLater()
                    except Exception:
                        pass
                    self._register_all_notebooks_from_current_onenote(
                        force=force,
                        _prefetched_records=[
                            dict(record) for record in (result.get("records") or [])
                        ],
                        _prefetched_source=str(result.get("source") or "MAC_SIDEBAR"),
                        _prefetched_sig=result.get("sig") if isinstance(result.get("sig"), dict) else sig,
                        _prefetched_error=str(result.get("error") or ""),
                    )

                worker.done.connect(_on_done)
                worker.start()
                return

            try:
                self.connection_status_label.setText("종합 새로고침 중...")
                QApplication.processEvents()
            except Exception:
                pass

            if _prefetched_sig is not None:
                sig = dict(_prefetched_sig or {})
            else:
                try:
                    sig = (
                        build_window_signature_quick(
                            onenote_window,
                            self.settings.get("connection_signature")
                            if isinstance(self.settings.get("connection_signature"), dict)
                            else None,
                        )
                        if IS_MACOS
                        else _build_connection_signature_for_save(
                            onenote_window,
                            self.settings.get("connection_signature")
                            if isinstance(self.settings.get("connection_signature"), dict)
                            else None,
                        )
                    )
                except Exception as e:
                    print(f"[WARN][AGG_REFRESH] signature build failed: {e}")
                    sig = {}

            notebook_nodes: List[Dict[str, Any]] = []
            seen_keys: Set[str] = set()
            com_error = ""
            current_open_keys: Set[str] = set()

            def _append_notebook_node(
                nb_name: str,
                *,
                notebook_id: str = "",
                notebook_path: str = "",
                notebook_url: str = "",
                last_accessed_at: int = 0,
                notebook_source: str = "",
                is_open: bool = False,
            ) -> None:
                nb_name_clean = _strip_stale_favorite_prefix(str(nb_name or "").strip())
                name_key = _normalize_notebook_name_key(nb_name_clean)
                id_key = str(notebook_id or "").strip().casefold()
                dedupe_key = f"id:{id_key}" if id_key else f"name:{name_key}"
                if not nb_name_clean or not name_key or dedupe_key in seen_keys:
                    return
                seen_keys.add(dedupe_key)
                target = {"sig": sig, "notebook_text": nb_name_clean}
                if is_open:
                    target["is_open"] = True
                if notebook_id:
                    target["notebook_id"] = str(notebook_id).strip()
                if notebook_path:
                    target["path"] = str(notebook_path).strip()
                if notebook_url:
                    target["url"] = str(notebook_url).strip()
                try:
                    notebook_last_accessed_at = max(0, int(last_accessed_at or 0))
                except Exception:
                    notebook_last_accessed_at = 0
                if notebook_last_accessed_at:
                    target["last_accessed_at"] = notebook_last_accessed_at
                if notebook_source:
                    target["source"] = str(notebook_source).strip()
                notebook_nodes.append(
                    {
                        "type": "notebook",
                        "id": str(uuid.uuid4()),
                        "name": nb_name_clean,
                        "target": target,
                        "is_open": bool(is_open),
                    }
                )

            prefetched_records = _prefetched_records if _prefetched_records is not None else None
            try:
                source_records = (
                    [dict(record) for record in prefetched_records]
                    if prefetched_records is not None
                    else _get_open_notebook_records_via_com(refresh=True)
                )
                current_open_keys = {
                    _normalize_notebook_name_key((record or {}).get("name"))
                    for record in source_records
                    if _normalize_notebook_name_key((record or {}).get("name"))
                }
                for record in source_records:
                    _append_notebook_node(
                        record.get("name", ""),
                        notebook_id=record.get("id", ""),
                        notebook_path=record.get("path", ""),
                        notebook_url=record.get("url", ""),
                        last_accessed_at=record.get("last_accessed_at", 0),
                        notebook_source=record.get("source", ""),
                        is_open=True,
                    )
            except Exception as e:
                com_error = str(e)
                print(f"[WARN][AGG_REFRESH][COM] {e}")
            if _prefetched_error and not com_error:
                com_error = _prefetched_error

            source = _prefetched_source or "COM"
            allow_ui_fallback = not (
                IS_MACOS and prefetched_records is not None and bool(prefetched_records)
            )
            if IS_MACOS and _prefetched_error and "AX 직접 목록" in _prefetched_error:
                allow_ui_fallback = False
            if IS_MACOS and not notebook_nodes and onenote_window is not None and allow_ui_fallback:
                source = "CONNECTED_WINDOW"
                try:
                    fallback_window = (
                        MacWindow(dict(sig))
                        if isinstance(sig, dict) and sig
                        else onenote_window
                    )
                    for nb_name in mac_current_open_notebook_names(fallback_window):
                        key = _normalize_notebook_name_key(nb_name)
                        if key:
                            current_open_keys.add(key)
                        _append_notebook_node(nb_name, notebook_path=nb_name, is_open=True)
                except Exception as e:
                    if not com_error:
                        com_error = str(e)
                    print(f"[WARN][AGG_REFRESH][MAC_WINDOW] {e}")

            if not notebook_nodes and allow_ui_fallback:
                source = "UI"
                try:
                    ensure_pywinauto()
                    if hasattr(self, "_bring_onenote_to_front"):
                        self._bring_onenote_to_front()
                    if not getattr(self, "tree_control", None):
                        self.tree_control = _find_tree_or_list(onenote_window)
                    for nb_name in _collect_root_notebook_names_from_tree(
                        self.tree_control,
                        limit=512,
                    ):
                        key = _normalize_notebook_name_key(nb_name)
                        if key:
                            current_open_keys.add(key)
                        _append_notebook_node(nb_name, is_open=True)
                except Exception as e:
                    print(f"[WARN][AGG_REFRESH][UI_FALLBACK] {e}")

            known_records_added = 0
            try:
                known_records = _collect_known_notebook_name_records(self.settings)
            except Exception as e:
                known_records = []
                print(f"[WARN][AGG_REFRESH][KNOWN] {e}")
            before_known_merge = len(notebook_nodes)
            for record in known_records:
                _append_notebook_node(
                    record.get("name", ""),
                    notebook_id=record.get("id", ""),
                    notebook_path=record.get("path", ""),
                    notebook_url=record.get("url", ""),
                    last_accessed_at=record.get("last_accessed_at", 0),
                    notebook_source=record.get("source", ""),
                    is_open=(
                        _normalize_notebook_name_key(record.get("name"))
                        in current_open_keys
                    ),
                )
            known_records_added = max(0, len(notebook_nodes) - before_known_merge)
            if known_records_added:
                source = f"{source}+KNOWN"

            if not notebook_nodes:
                message = "등록할 전자필기장을 찾지 못했습니다."
                if com_error:
                    message += f"\n\nCOM 조회 오류: {com_error}"
                status_message = "열린 전자필기장을 찾지 못했습니다."
                if com_error:
                    status_message = f"{status_message} {str(com_error).splitlines()[0][:140]}"
                self.update_status_and_ui(status_message, True)
                if not force:
                    QMessageBox.information(self, "안내", message)
                return

            self._invalidate_aggregate_cache(invalidate_classified_keys=True)
            categorized = self._build_aggregate_categorized_display_nodes(notebook_nodes)

            self._aggregate_reclassify_in_progress = True
            try:
                if is_agg:
                    self._load_favorites_into_center_tree(categorized)
                    self._fav_reset_undo_context_from_data(
                        categorized,
                        reason="aggregate_onenote_refresh",
                    )
                self._persist_active_aggregate_data(categorized)
            finally:
                self._aggregate_reclassify_in_progress = False

            unclassified_count = len(categorized[0].get("children") or []) if categorized else 0
            classified_count = len(categorized[1].get("children") or []) if len(categorized) > 1 else 0
            if IS_WINDOWS:
                counted_nodes = self._collect_notebook_nodes_from_nodes(notebook_nodes)
                total_count = len(counted_nodes)
                open_checked_count = sum(
                    1
                    for node in counted_nodes
                    if bool(node.get("is_open") or (node.get("target") or {}).get("is_open"))
                )
            else:
                total_count = len(notebook_nodes)
                open_checked_count = sum(
                    1 for node in notebook_nodes if bool(node.get("is_open"))
                )
            elapsed_ms = (time.perf_counter() - started_at) * 1000.0
            prefix = "종합 새로고침 완료" if is_agg else "종합 데이터 자동 갱신 완료"
            self.update_status_and_ui(
                (
                    f"{prefix}: "
                    f"전체 {total_count}개, "
                    f"열림 체크 {open_checked_count}개, "
                    f"분류 안 됨 {unclassified_count}개, "
                    f"분류됨 {classified_count}개 "
                    f"(known+{known_records_added}) "
                    f"({source}, {elapsed_ms:.0f}ms)"
                ),
                True,
            )
        except Exception as e:
            print(f"[ERROR][AGG_REFRESH] {e}")
            traceback.print_exc()
            QMessageBox.warning(self, "오류", f"종합 새로고침 실패: {e}")
        finally:
            if refresh_button is not None:
                refresh_button.setEnabled(
                    getattr(self, "active_buffer_id", None) == AGG_BUFFER_ID
                )

_publish_context(globals())

# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin29:

    def _try_activate_favorite_fastpath_v2(
        self,
        item: QTreeWidgetItem,
        sig: Dict[str, Any],
        target: Dict[str, Any],
        display_name: str,
        *,
        started_at: Optional[float] = None,
    ) -> bool:
        direct_source = "same_window"
        if self._is_sig_same_as_connected_window(sig):
            win = self.onenote_window
        else:
            direct_source = "direct_connect"
            win = resolve_window_target(sig)
            if win is None:
                return False
            self.onenote_window = win
            try:
                save_connection_info(self.onenote_window)
            except Exception:
                pass
            self._cache_tree_control()

        tree = self.tree_control or _find_tree_or_list(self.onenote_window)
        self.tree_control = tree
        if not tree and not IS_MACOS:
            return False

        notebook_text = _strip_stale_favorite_prefix(
            str((target or {}).get("notebook_text") or "").strip()
        )
        section_text = str((target or {}).get("section_text") or "").strip()
        display_text = _strip_stale_favorite_prefix(display_name)
        target_kind = "section" if section_text else "notebook"
        expected_text = section_text
        expected_notebook_text = notebook_text
        resolved_name = ""
        resolved_notebook_id = ""
        resolution_mode = "quick"
        selected_notebook_item = None
        section_notebook_checked = False

        def _attempt_select(kind: str, text: str) -> bool:
            nonlocal selected_notebook_item
            if not text:
                return False
            if kind == "section":
                return select_section_by_text(self.onenote_window, text, tree)
            selected_notebook_item = select_notebook_item_by_text(
                self.onenote_window,
                text,
                tree,
                center_after_select=False,
            )
            return selected_notebook_item is not None

        def _activate_macos_notebook_context() -> bool:
            if not (IS_MACOS and target_kind == "notebook"):
                return False
            requested_name = expected_notebook_text or display_text
            if not requested_name:
                return False
            if _mac_outline_notebook_matches(self.onenote_window, requested_name):
                notebook_result = {
                    "ok": True,
                    "name": requested_name,
                    "source": "current",
                    "error": "",
                }
            else:
                refreshed_win = self._refresh_macos_onenote_main_window()
                if refreshed_win is not None:
                    self.onenote_window = refreshed_win
                notebook_result = _mac_ensure_notebook_context_for_section(
                    self.onenote_window,
                    requested_name,
                )
            if not notebook_result.get("ok", True):
                fail_msg = notebook_result.get("error") or (
                    f"전자필기장 '{requested_name}'을(를) 열지 못했습니다."
                )
                self.update_status_and_ui(
                    fail_msg,
                    self.center_button.isEnabled(),
                )
                print(
                    "[DBG][FAV][FASTPATH]",
                    "notebook_context_abort",
                    f"error={fail_msg!r}",
                    f"elapsed_ms={(time.perf_counter() - (started_at or time.perf_counter())) * 1000.0:.1f}",
                )
                return True
            final_name = str(notebook_result.get("name") or requested_name).strip()
            self._sync_favorite_notebook_target(
                item,
                final_name,
                resolved_notebook_id,
            )
            self._cache_tree_control()
            self._cancel_pending_center_after_activate()
            self.update_status_and_ui(f"활성화: '{final_name}'", True)
            print(
                "[DBG][FAV][FASTPATH]",
                "mac_notebook_context",
                f"text={final_name!r}",
                f"source={notebook_result.get('source')!r}",
                f"elapsed_ms={(time.perf_counter() - (started_at or time.perf_counter())) * 1000.0:.1f}",
            )
            return True

        def _ensure_section_notebook_context() -> bool:
            nonlocal tree, section_notebook_checked
            if not (IS_MACOS and target_kind == "section" and expected_notebook_text):
                return True
            if section_notebook_checked:
                return True
            section_notebook_checked = True
            refreshed_win = self._refresh_macos_onenote_main_window()
            if refreshed_win is not None:
                self.onenote_window = refreshed_win
            notebook_result = _mac_ensure_notebook_context_for_section(
                self.onenote_window,
                expected_notebook_text,
            )
            if not notebook_result.get("ok", True):
                fail_msg = notebook_result.get("error") or (
                    f"전자필기장 '{expected_notebook_text}'을(를) 열지 못해 "
                    f"섹션 '{expected_text}' 복구를 중단했습니다."
                )
                self.update_status_and_ui(
                    fail_msg,
                    self.center_button.isEnabled(),
                )
                print(
                    "[DBG][FAV][FASTPATH]",
                    "section_notebook_abort",
                    f"error={fail_msg!r}",
                    f"elapsed_ms={(time.perf_counter() - (started_at or time.perf_counter())) * 1000.0:.1f}",
                )
                return False
            self._cache_tree_control()
            tree = self.tree_control or tree
            return True

        if IS_MACOS and target_kind == "notebook":
            if _activate_macos_notebook_context():
                return True

        ok = False
        if target_kind == "section":
            if not _ensure_section_notebook_context():
                return True
            ok = _attempt_select("section", expected_text)
        else:
            quick_candidates = []
            for cand in (notebook_text, display_text):
                if cand and cand not in quick_candidates:
                    quick_candidates.append(cand)
            for cand in quick_candidates:
                if _attempt_select("notebook", cand):
                    expected_text = cand
                    ok = True
                    break

        if not ok:
            requested_notebook_id = str((target or {}).get("notebook_id") or "").strip()
            if target_kind == "notebook" and not requested_notebook_id:
                visible_names = _collect_root_notebook_names_from_tree(tree)
                if visible_names:
                    quick_error = _build_notebook_not_found_error(
                        notebook_text or display_text,
                        visible_names,
                    )
                    stale_name = self._mark_favorite_item_stale(item, display_name)
                    fail_msg = quick_error or f"항목 찾기 실패: '{stale_name or display_name}'"
                    self.update_status_and_ui(
                        fail_msg,
                        self.center_button.isEnabled(),
                    )
                    print(
                        "[DBG][FAV][FASTPATH]",
                        "quick_abort",
                        f"error={fail_msg!r}",
                        f"elapsed_ms={(time.perf_counter() - (started_at or time.perf_counter())) * 1000.0:.1f}",
                    )
                    return True
            target_info = _resolve_favorite_activation_target(target, display_name)
            if not target_info.get("ok", True):
                stale_name = self._mark_favorite_item_stale(item, display_name)
                fail_msg = (
                    target_info.get("error")
                    or f"항목 찾기 실패: '{stale_name or display_name}'"
                )
                self.update_status_and_ui(
                    fail_msg,
                    self.center_button.isEnabled(),
                )
                print(
                    "[DBG][FAV][FASTPATH]",
                    "resolve_abort",
                    f"error={fail_msg!r}",
                    f"elapsed_ms={(time.perf_counter() - (started_at or time.perf_counter())) * 1000.0:.1f}",
                )
                return True
            target_kind = target_info.get("target_kind") or target_kind
            expected_text = target_info.get("expected_center_text") or expected_text
            expected_notebook_text = (
                target_info.get("expected_notebook_text")
                or expected_notebook_text
            )
            resolved_name = target_info.get("resolved_name") or ""
            resolved_notebook_id = target_info.get("resolved_notebook_id") or ""
            resolution_mode = "resolved"

        print(
            "[DBG][FAV][FASTPATH]",
            direct_source,
            f"kind={target_kind}",
            f"text={expected_text!r}",
            f"mode={resolution_mode}",
            f"elapsed_ms={(time.perf_counter() - (started_at or time.perf_counter())) * 1000.0:.1f}",
        )

        try:
            self.onenote_window.set_focus()
        except Exception:
            pass

        if _activate_macos_notebook_context():
            return True

        if not ok:
            if not _ensure_section_notebook_context():
                return True
            ok = _attempt_select(target_kind, expected_text)

        if not ok:
            self._cache_tree_control()
            tree = self.tree_control or tree
            if tree or IS_MACOS:
                ok = _attempt_select(target_kind, expected_text)

        if not ok:
            print(
                "[DBG][FAV][FASTPATH]",
                f"{direct_source}_select_failed",
                f"kind={target_kind}",
                f"text={expected_text!r}",
            )
            return False

        if target_kind == "notebook":
            self._sync_favorite_notebook_target(
                item,
                resolved_name,
                resolved_notebook_id,
            )
        elif IS_MACOS:
            page_text = str((target or {}).get("page_text") or "").strip()
            if page_text:
                self._restore_macos_page_context(page_text)
            self._restore_favorite_item_from_stale(item, display_name)

        self.center_selected_item_action(
            debug_source="fav_fastpath",
            started_at=started_at,
            skip_precheck=True,
            allow_retry=(target_kind != "notebook"),
            preselected_item=selected_notebook_item if target_kind == "notebook" else None,
            preselected_tree_control=tree if target_kind == "notebook" else None,
            expected_text=expected_text,
        )
        return True

    def eventFilter(self, obj, event):
        try:
            list_widget = getattr(self, "onenote_list_widget", None)
            if list_widget is not None:
                event_type = event.type()
                if obj is list_widget.viewport():
                    if event_type == QEvent.Type.MouseButtonPress:
                        app = QApplication.instance()
                        delay_ms = 120
                        if app is not None:
                            try:
                                delay_ms = int(app.doubleClickInterval()) + 30
                            except Exception:
                                delay_ms = 120
                        self._schedule_onenote_list_auto_refresh(delay_ms=delay_ms)
                    elif event_type == QEvent.Type.MouseButtonDblClick:
                        self._cancel_pending_onenote_list_auto_refresh()
                elif obj is list_widget and event_type == QEvent.Type.KeyPress:
                    if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                        current_item = list_widget.currentItem()
                        if current_item is not None:
                            self.connect_and_center_from_list_item(current_item)
                            return True
        except Exception:
            pass
        return super().eventFilter(obj, event)

    def _start_auto_reconnect(self):
        self.refresh_button.setEnabled(False)
        self._reconnect_worker = ReconnectWorker()
        self._reconnect_worker.finished.connect(self._on_reconnect_done)
        self._reconnect_worker.start()

    def _run_macos_auto_reconnect(self):
        if self._reconnect_worker is not None:
            return
        try:
            ensure_pywinauto()
            current_sig = self.settings.get("connection_signature")
            win = reacquire_window_by_signature(current_sig or {})
            if win:
                next_sig = build_window_signature_quick(win, current_sig)
                title = _preferred_connected_window_title_quick(win, next_sig)
                status = f"(자동 재연결) '{title}'"
                payload = {"ok": True, "status": status, "sig": next_sig}
            else:
                payload = {
                    "ok": False,
                    "status": "(재연결 실패) 이전 앱을 찾을 수 없습니다.",
                }
        except Exception as e:
            payload = {"ok": False, "status": f"연결되지 않음 (오류: {e})"}
        self._on_reconnect_done(payload)

    def _sync_onenote_list_action_buttons(self) -> None:
        list_widget = getattr(self, "onenote_list_widget", None)
        connect_button = getattr(self, "connect_selected_list_button", None)
        if list_widget is None or connect_button is None:
            return
        try:
            current_item = list_widget.currentItem()
        except Exception:
            current_item = None
        connect_button.setEnabled(bool(current_item) and list_widget.isEnabled())

    def _connect_selected_onenote_list_item(self) -> None:
        item = None
        try:
            item = self.onenote_list_widget.currentItem()
        except Exception:
            item = None
        if item is None:
            self.update_status_and_ui("먼저 연결할 OneNote 창을 선택해 주세요.", False)
            return
        self.connect_and_center_from_list_item(item)

_publish_context(globals())

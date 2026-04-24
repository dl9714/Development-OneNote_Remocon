# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin30:

    def _on_reconnect_done(self, payload):
        self._reconnect_worker = None
        status = payload.get("status", "연결되지 않음")
        if payload.get("ok"):
            ensure_pywinauto()
            sig = payload.get("sig", {})
            if IS_MACOS:
                target_info = payload.get("target_info") if isinstance(payload, dict) else None
                target = MacWindow(dict(target_info or sig))
            else:
                target = resolve_window_target(sig)

            if target:
                self.onenote_window = target
                if IS_MACOS:
                    self.settings["connection_signature"] = _merge_connection_signature(
                        sig,
                        self.settings.get("connection_signature")
                        if isinstance(self.settings.get("connection_signature"), dict)
                        else None,
                    )
                    try:
                        save_settings(self.settings)
                    except Exception:
                        pass
                else:
                    try:
                        save_connection_info(self.onenote_window)
                    except Exception:
                        pass
                    self._remember_connection_signature(self.onenote_window)
                self.update_status_and_ui(f"연결됨: {status}", True)
                QTimer.singleShot(0, self._cache_tree_control)
                self.refresh_onenote_list()
                return

        self.onenote_window = None
        self.tree_control = None
        self.update_status_and_ui(f"상태: {status}", False)
        self.refresh_onenote_list()

    def refresh_onenote_list(self, reset_retry_budget: bool = True):
        if self._scanner_worker and self._scanner_worker.isRunning():
            return
        self._last_onenote_list_refresh_at = time.monotonic()
        if self._onenote_list_refresh_timer.isActive():
            self._onenote_list_refresh_timer.stop()
        if reset_retry_budget:
            self._mac_empty_scan_retry_attempts = 0
            if self._mac_empty_scan_retry_timer.isActive():
                self._mac_empty_scan_retry_timer.stop()

        self.onenote_list_widget.clear()
        self.onenote_list_widget.addItem("OneNote 창을 검색 중입니다...")
        self.onenote_list_widget.setEnabled(False)
        self.refresh_button.setEnabled(False)
        self.connect_selected_list_button.setEnabled(False)

        self._scanner_worker = OneNoteWindowScanner(self.my_pid)
        self._scanner_worker.done.connect(self._on_onenote_list_ready)
        self._scanner_worker.start()

    def _on_onenote_list_ready(self, results: List[Dict]):
        self.onenote_windows_info = results
        self.onenote_list_widget.clear()
        print(f"[DBG][LIST] onenote_windows={len(results)}")
        selection_key = self._pending_onenote_list_selection_key
        self._pending_onenote_list_selection_key = None

        if not results:
            if IS_MACOS and self._mac_empty_scan_retry_attempts < 2:
                self._mac_empty_scan_retry_attempts += 1
                retry_delay_ms = 1200 * self._mac_empty_scan_retry_attempts
                self.onenote_list_widget.addItem(
                    "실행 중인 OneNote 창을 찾지 못했습니다. 잠시 후 다시 확인합니다..."
                )
                self._mac_empty_scan_retry_timer.start(retry_delay_ms)
                print(
                    "[DBG][LIST][MAC]",
                    f"empty_retry={self._mac_empty_scan_retry_attempts}",
                    f"delay_ms={retry_delay_ms}",
                )
            else:
                self.onenote_list_widget.addItem("실행 중인 OneNote 창을 찾지 못했습니다.")
        else:
            self._mac_empty_scan_retry_attempts = 0
            if self._mac_empty_scan_retry_timer.isActive():
                self._mac_empty_scan_retry_timer.stop()
            duplicate_title_counts: Dict[str, int] = {}
            for info in results:
                display_title = self._preferred_onenote_list_display_title(info)
                title_key = display_title.strip().casefold()
                duplicate_title_counts[title_key] = (
                    duplicate_title_counts.get(title_key, 0) + 1
                )
            for info in results:
                item = QListWidgetItem(
                    self._format_onenote_list_item_label(info, duplicate_title_counts)
                )
                item.setData(Qt.ItemDataRole.UserRole, copy.deepcopy(info))
                self.onenote_list_widget.addItem(item)
                item_key = (info.get("handle"), info.get("pid"), info.get("title"))
                if selection_key and item_key == selection_key:
                    self.onenote_list_widget.setCurrentItem(item)
            if self.onenote_list_widget.currentItem() is None and self.onenote_list_widget.count() > 0:
                self.onenote_list_widget.setCurrentRow(0)

        self.onenote_list_widget.setEnabled(True)
        self.refresh_button.setEnabled(True)
        self._sync_onenote_list_action_buttons()

    def _retry_onenote_list_after_empty_macos_scan(self):
        if not IS_MACOS:
            return
        if self._scanner_worker and self._scanner_worker.isRunning():
            return
        self.refresh_onenote_list(reset_retry_budget=False)

    def _remember_connection_signature(self, window_element) -> None:
        try:
            current_sig = self.settings.get("connection_signature")
            next_sig = build_window_signature(window_element)
            self.settings["connection_signature"] = _merge_connection_signature(
                next_sig,
                current_sig if isinstance(current_sig, dict) else None,
            )
        except Exception:
            pass

    def _preferred_onenote_list_display_title(self, info: Dict[str, Any]) -> str:
        raw_title = str(info.get("title") or "").strip()
        app_name = str(info.get("app_name") or "").strip()
        class_name = str(info.get("class_name") or "").strip()
        bundle_id = str(info.get("bundle_id") or class_name or "").strip()

        if raw_title and raw_title.casefold() not in MACOS_GENERIC_ONENOTE_TITLES:
            return raw_title

        if IS_MACOS and bundle_id == ONENOTE_MAC_BUNDLE_ID:
            # 창 목록은 앱 시작 직후에도 즉시 그려져야 한다. OneNote 내부
            # 접근(사이드바/최근 목록/페이지 트리)은 여기서 하지 않고,
            # 연결/특수 기능 실행 시점의 백그라운드 작업에 맡긴다.
            hydrated_title = _preferred_connected_window_title_quick(
                MacWindow(dict(info)),
                info,
            )
            if (
                hydrated_title
                and hydrated_title.casefold() not in MACOS_GENERIC_ONENOTE_TITLES
            ):
                return hydrated_title

        candidate_sigs: List[Dict[str, Any]] = []
        current_window = getattr(self, "onenote_window", None)
        if current_window is not None:
            try:
                candidate_sigs.append(
                    build_window_signature_quick(
                        current_window,
                        self.settings.get("connection_signature")
                        if isinstance(self.settings.get("connection_signature"), dict)
                        else None,
                    )
                )
            except Exception:
                pass
        saved_sig = self.settings.get("connection_signature")
        if isinstance(saved_sig, dict):
            candidate_sigs.append(saved_sig)

        info_handle = int(info.get("handle") or 0)
        info_pid = int(info.get("pid") or 0)
        for sig in candidate_sigs:
            sig_title = str(sig.get("title") or "").strip()
            if not sig_title:
                continue
            sig_bundle = str(sig.get("bundle_id") or sig.get("class_name") or "").strip()
            sig_handle = int(sig.get("handle") or 0)
            sig_pid = int(sig.get("pid") or 0)
            if info_handle and sig_handle and info_handle == sig_handle:
                return sig_title
            if bundle_id and sig_bundle and bundle_id == sig_bundle:
                if info_pid and sig_pid and info_pid == sig_pid:
                    return sig_title
                if bundle_id == ONENOTE_MAC_BUNDLE_ID:
                    return sig_title

        return raw_title or app_name or class_name or "이름 없는 창"

    def _format_onenote_list_item_label(
        self,
        info: Dict[str, Any],
        duplicate_title_counts: Optional[Dict[str, int]] = None,
    ) -> str:
        display_title = self._preferred_onenote_list_display_title(info)
        title_key = display_title.strip().casefold()
        duplicates = 0
        if duplicate_title_counts:
            duplicates = int(duplicate_title_counts.get(title_key, 0))

        if duplicates <= 1:
            return display_title

        if IS_MACOS:
            app_name = str(info.get("app_name") or "").strip()
            if app_name and app_name.casefold() not in {"microsoft onenote", "onenote"}:
                return f"{display_title} ({app_name})"
            pid = info.get("pid")
            if pid:
                return f"{display_title} (pid {pid})"
            return display_title

        class_name = str(info.get("class_name") or "").strip()
        if class_name:
            return f"{display_title}  [{class_name}]"
        return display_title

    def _cache_tree_control(self):
        if IS_MACOS:
            # macOS resolves the active OneNote row on demand. Doing a full tree
            # lookup during startup can block the Qt main thread on Accessibility.
            self.tree_control = None
            return
        self.tree_control = _find_tree_or_list(self.onenote_window)
        if self.tree_control:
            try:
                _ = self.tree_control.children()
            except Exception:
                pass

    def _perform_connection(self, info: Dict) -> bool:
        t0 = time.perf_counter()
        ensure_pywinauto()
        if not _pwa_ready:
            self.update_status_and_ui("pywinauto가 준비되지 않았습니다.", False)
            return False
        try:
            print(
                "[DBG][CONNECT] try",
                f"handle={info.get('handle')}",
                f"pid={info.get('pid')}",
                f"class={info.get('class_name')}",
                f"title={info.get('title')!r}",
            )
            target = None
            target = resolve_window_target(info)
            if target is None:
                raise ElementNotFoundError

            self.onenote_window = target
            window_title = _preferred_connected_window_title(self.onenote_window, info)
            save_connection_info(self.onenote_window)
            self._remember_connection_signature(self.onenote_window)
            elapsed_ms = (time.perf_counter() - t0) * 1000.0
            print(
                f"[DBG][CONNECT] success title={window_title!r} "
                f"elapsed_ms={elapsed_ms:.1f} at_s={(time.perf_counter() - self._t_boot):.3f}"
            )

            status_text = f"연결됨: '{window_title}'"
            self.update_status_and_ui(status_text, True)
            self._cache_tree_control()
            return True

        except ElementNotFoundError:
            elapsed_ms = (time.perf_counter() - t0) * 1000.0
            print(
                "[DBG][CONNECT] fail: target not found/visible "
                f"elapsed_ms={elapsed_ms:.1f} at_s={(time.perf_counter() - self._t_boot):.3f}"
            )
            self.update_status_and_ui("연결 실패: 선택한 창이 보이지 않습니다.", False)
            self.refresh_onenote_list()
            return False
        except Exception as e:
            elapsed_ms = (time.perf_counter() - t0) * 1000.0
            print(
                f"[DBG][CONNECT] exception elapsed_ms={elapsed_ms:.1f} "
                f"at_s={(time.perf_counter() - self._t_boot):.3f} err={e}"
            )
            self.update_status_and_ui(f"연결 실패: {e}", False)
            return False

_publish_context(globals())

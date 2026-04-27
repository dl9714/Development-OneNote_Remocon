# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin43:

    def _mac_favorite_notebook_repeat_key(
        self,
        item: Optional[QTreeWidgetItem],
        notebook_name: str = "",
    ):
        if not (IS_MACOS and item is not None):
            return None
        name = str(notebook_name or "").strip()
        if not name:
            try:
                payload = item.data(0, ROLE_DATA) or {}
                target = payload.get("target") if isinstance(payload, dict) else {}
                if isinstance(target, dict):
                    name = str(target.get("notebook_text") or "").strip()
            except Exception:
                name = ""
        if not name:
            try:
                name = str(item.text(0) or "").strip()
            except Exception:
                name = ""
        notebook_key = _normalize_notebook_name_key(_strip_stale_favorite_prefix(name))
        if not notebook_key:
            return None
        window_key = self._mac_center_window_key() if hasattr(self, "_mac_center_window_key") else ()
        return id(item), window_key, notebook_key

    def _remember_mac_favorite_notebook_hit(
        self,
        item: Optional[QTreeWidgetItem],
        notebook_name: str,
    ) -> None:
        key = self._mac_favorite_notebook_repeat_key(item, notebook_name)
        if key is not None:
            self._last_macos_fav_notebook_hit = (
                key,
                time.monotonic(),
                str(notebook_name or "").strip(),
            )

    def _mac_repeat_favorite_notebook_hit(
        self,
        item: Optional[QTreeWidgetItem],
        node_type: str,
    ) -> bool:
        if not (IS_MACOS and node_type == "notebook"):
            return False
        last = getattr(self, "_last_macos_fav_notebook_hit", None)
        now = time.monotonic()
        if not (isinstance(last, tuple) and len(last) == 3):
            return False
        last_key, last_at, last_name = last
        if (now - float(last_at or 0.0)) > 0.85:
            return False

        window_key = self._mac_center_window_key() if hasattr(self, "_mac_center_window_key") else ()
        if (
            isinstance(last_key, tuple)
            and len(last_key) >= 2
            and id(item) == last_key[0]
            and window_key == last_key[1]
        ):
            key = last_key
        else:
            key = self._mac_favorite_notebook_repeat_key(item)
            if key != last_key:
                return False

        self._last_macos_fav_notebook_hit = (key, now, last_name)
        self._last_favorite_activation_at = now
        final_name = str(last_name or "").strip()
        status = f"활성화: '{final_name}'" if final_name else ""
        if status and self.connection_status_label.text() != status:
            self.update_status_and_ui(status, True)
        return True

    def _sync_favorite_notebook_target(
        self,
        item: Optional[QTreeWidgetItem],
        resolved_name: str,
        resolved_notebook_id: str,
    ) -> Dict[str, Any]:
        current_name = ""
        if item is not None:
            try:
                current_name = item.text(0) or ""
            except Exception:
                current_name = ""

        result = {
            "display_name": current_name,
            "changed": False,
            "name_changed": False,
            "was_stale": False,
        }
        if item is None:
            return result

        try:
            payload = item.data(0, ROLE_DATA) or {}
            if not isinstance(payload, dict):
                payload = {}

            target = payload.get("target") or {}
            if not isinstance(target, dict):
                target = {}

            display_name = current_name
            stale_prefixes = ("(구) ", "(old) ")
            for prefix in stale_prefixes:
                if display_name.startswith(prefix):
                    display_name = display_name[len(prefix):]
                    result["was_stale"] = True
                    break

            clean_name = display_name
            resolved_name = (resolved_name or "").strip()
            resolved_notebook_id = (resolved_notebook_id or "").strip()

            if resolved_name and clean_name != resolved_name:
                clean_name = resolved_name
                result["name_changed"] = True

            if current_name != clean_name:
                item.setText(0, clean_name)
                result["changed"] = True

            if resolved_name and target.get("notebook_text") != resolved_name:
                target["notebook_text"] = resolved_name
                result["changed"] = True

            if resolved_notebook_id and target.get("notebook_id") != resolved_notebook_id:
                target["notebook_id"] = resolved_notebook_id
                result["changed"] = True

            result["display_name"] = clean_name or current_name

            if result["changed"]:
                payload["target"] = target
                item.setData(0, ROLE_DATA, payload)
                self._save_favorites()
        except Exception:
            print("[ERR][FAV][SYNC] exception")
            traceback.print_exc()

        return result

    def _handle_favorite_activation_result(
        self,
        item: Optional[QTreeWidgetItem],
        sig: Dict[str, Any],
        display_name: str,
        result: Dict[str, Any],
    ) -> None:
        try:
            connected = self._apply_connected_window_info(result.get("window_info"))
            ok = bool(result.get("ok"))
            target_kind = result.get("target_kind")
            expected_center_text = result.get("expected_center_text")
            resolved_name = str(result.get("resolved_name") or "").strip()
            resolved_notebook_id = str(result.get("resolved_notebook_id") or "").strip()

            if connected and ok:
                notebook_sync = {
                    "display_name": display_name,
                    "changed": False,
                    "name_changed": False,
                    "was_stale": False,
                }
                payload = item.data(0, ROLE_DATA) if item is not None else {}
                target = payload.get("target") if isinstance(payload, dict) else {}
                if target_kind == "notebook":
                    notebook_sync = self._sync_favorite_notebook_target(
                        item, resolved_name, resolved_notebook_id
                    )
                elif IS_MACOS and isinstance(target, dict):
                    page_text = str(target.get("page_text") or "").strip()
                    if page_text:
                        self._restore_macos_page_context(page_text)
                    notebook_sync = self._restore_favorite_item_from_stale(
                        item, display_name
                    )

                aligned_now = False
                if (
                    target_kind == "notebook" and not IS_MACOS
                    and expected_center_text
                    and getattr(self, "onenote_window", None) is not None
                ):
                    try:
                        aligned_now, _ = scroll_selected_item_to_center(
                            self.onenote_window,
                            self.tree_control,
                        )
                    except Exception:
                        aligned_now = False

                if target_kind in ("section", "notebook") and expected_center_text:
                    if not aligned_now:
                        self._schedule_center_after_activate(
                            sig,
                            expected_center_text,
                            target_kind=target_kind,
                        )
                else:
                    self._cancel_pending_center_after_activate()

                final_name = notebook_sync.get("display_name") or display_name
                if notebook_sync.get("was_stale"):
                    self.update_status_and_ui(
                        f"활성화: '{final_name}' (이름 복원)", True
                    )
                elif notebook_sync.get("name_changed"):
                    self.update_status_and_ui(
                        f"활성화: '{final_name}' (이름 갱신)", True
                    )
                else:
                    self.update_status_and_ui(f"활성화: '{final_name}'", True)
                return

            current_name = ""
            if item is not None:
                try:
                    current_name = item.text(0) or ""
                except Exception:
                    current_name = ""

            stale_prefixes = ("(구) ", "(old) ")
            if item is not None and not any(
                current_name.startswith(prefix) for prefix in stale_prefixes
            ):
                fail_msg = result.get("error") or f"항목 찾기 실패: '{current_name}'"
                self.update_status_and_ui(fail_msg, True)
            else:
                fail_msg = result.get("error") or f"항목 찾기 실패: '{display_name}'"
                self.update_status_and_ui(fail_msg, True)
        except Exception as e:
            print("[ERR][FAV][ACTIVATE][RESULT] exception")
            traceback.print_exc()
            self.update_status_and_ui(f"즐겨찾기 처리 중 오류: {e}", True)

    def _apply_connected_window_info(self, info: Optional[Dict[str, Any]]) -> bool:
        if not info or not info.get("handle"):
            return False
        try:
            if IS_MACOS:
                self.onenote_window = MacWindow(dict(info))
                next_sig = build_window_signature_quick(
                    self.onenote_window,
                    self.settings.get("connection_signature")
                    if isinstance(self.settings.get("connection_signature"), dict)
                    else None,
                )
                self.settings["connection_signature"] = _merge_connection_signature(
                    next_sig,
                    self.settings.get("connection_signature")
                    if isinstance(self.settings.get("connection_signature"), dict)
                    else None,
                )
                try:
                    save_settings(self.settings)
                except Exception:
                    pass
                return True

            self.onenote_window = resolve_window_target(info)
            if self.onenote_window is None:
                raise ElementNotFoundError
            save_connection_info(self.onenote_window)
            self._cache_tree_control()
            return True
        except Exception:
            return False

    def _schedule_center_after_activate(
        self,
        sig: Dict[str, Any],
        expected_text: str,
        *,
        target_kind: str = "",
    ):
        self._cancel_pending_center_after_activate()
        if not sig or not expected_text:
            return

        self._center_request_seq += 1
        request_seq = self._center_request_seq
        target_kind = (target_kind or "").strip().lower()
        timer = QTimer(self)
        timer.setSingleShot(True)
        self._pending_center_timer = timer

        def _start_worker():
            if self._pending_center_timer is not timer:
                return

            self._pending_center_timer = None
            worker = CenterAfterActivateWorker(
                sig,
                expected_text,
                target_kind=target_kind,
                parent=self,
            )
            self._center_worker = worker
            self._retain_qthread_until_finished(worker, "_center_worker")

            def _on_done(ok: bool, selected_name: str):
                if self._center_worker is not worker:
                    return

                if ok and selected_name:
                    self._dbg_hot(
                        f"[DBG][CENTER][DONE] selected='{selected_name}' req={request_seq}"
                    )
                else:
                    self._dbg_hot(f"[DBG][CENTER][SKIP] req={request_seq}")

            worker.done.connect(_on_done)
            worker.start()

        timer.timeout.connect(_start_worker)
        timer.start(40 if target_kind == "notebook" else 120)

    def _activate_favorite_section(
        self,
        item: QTreeWidgetItem,
        *,
        started_at: Optional[float] = None,
    ):
        self._last_favorite_activation_at = time.monotonic()
        ensure_pywinauto()
        if not IS_MACOS and not _pwa_ready:
            self.update_status_and_ui(
                "오류: 자동화 모듈이 로드되지 않았습니다.",
                self.center_button.isEnabled(),
            )
            return

        node_type = item.data(0, ROLE_TYPE)
        payload = item.data(0, ROLE_DATA) or {}
        target = payload.get("target") or {}
        display_name = item.text(0)

        sig = target.get("sig") or {}
        if not sig:
            self.update_status_and_ui(
                "오류: 즐겨찾기에 대상 창 정보가 없습니다.",
                self.center_button.isEnabled(),
            )
            self._cancel_pending_favorite_activation()
            return

        self._cancel_pending_favorite_activation()
        self._cancel_pending_center_after_activate()
        if not (IS_MACOS and node_type == "notebook") and self._try_activate_favorite_fastpath(
            item,
            sig,
            target,
            display_name,
            started_at=started_at,
        ):
            return
        worker = FavoriteActivationWorker(
            sig=sig,
            target=target,
            display_name=display_name,
            auto_center_after_activate=self._auto_center_after_activate,
            parent=self,
        )
        self._favorite_activation_worker = worker
        self._retain_qthread_until_finished(worker, "_favorite_activation_worker")

        def _on_done(result: Dict[str, Any]):
            if self._favorite_activation_worker is not worker:
                return
            return self._handle_favorite_activation_result(
                item=item,
                sig=sig,
                display_name=display_name,
                result=result,
            )

        worker.done.connect(_on_done)
        worker.start()

_publish_context(globals())

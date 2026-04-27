# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin27:

    # ----------------- 14.1 창 닫기 이벤트 핸들러 (Geometry/Favorites 저장) -----------------
    def _save_window_state(self):
        """창 지오메트리와 스플리터 상태를 self.settings에 업데이트하고 파일에 즉시 저장합니다."""
        # self.settings (메모리)를 직접 수정합니다. load_settings()를 호출하지 않습니다.
        # 이렇게 함으로써 다른 세션 변경사항이 유지됩니다.
        if not self.isMinimized() and not self.isMaximized():
            geom = self.geometry()
            self.settings["window_geometry"] = {
                "x": geom.x(),
                "y": geom.y(),
                "width": geom.width(),
                "height": geom.height(),
            }

        try:
            splitter_states = {
                "main": self.main_splitter.saveState()
                .toBase64()
                .data()
                .decode("ascii"),
                "left": self.left_splitter.saveState()
                .toBase64()
                .data()
                .decode("ascii"),
            }
            if getattr(self, "codex_splitter", None) is not None:
                splitter_states["codex"] = (
                    self.codex_splitter.saveState().toBase64().data().decode("ascii")
                )
            self.settings["splitter_states"] = splitter_states
        except Exception as e:
            print(f"[WARN] 스플리터 상태 저장 실패: {e}")

        # 수정된 self.settings 객체 전체를 파일에 저장합니다.
        # 즐겨찾기 등 다른 모든 변경사항도 함께 저장됩니다.
        self._save_settings_to_file(immediate=True)

    def closeEvent(self, event):
        # 실행 중 QThread 정리 (종료 시 'Destroyed while thread is still running' 방지)
        busy_threads = []
        for attr in [
            "_reconnect_worker",
            "_scanner_worker",
            "_scan_worker",
            "_window_list_worker",
            "_center_worker",
            "_favorite_activation_worker",
            "_open_all_notebooks_worker",
            "_open_notebooks_refresh_worker",
            "_codex_location_lookup_worker",
        ]:
            t = getattr(self, attr, None)
            try:
                if t is not None and hasattr(t, "isRunning") and t.isRunning():
                    print(f"[DBG][THREAD][STOP] {attr} stopping...")
                    try:
                        t.requestInterruption()
                    except Exception:
                        pass
                    try:
                        t.quit()
                    except Exception:
                        pass
                    try:
                        t.wait(1500)
                    except Exception:
                        pass
                    try:
                        if t.isRunning():
                            busy_threads.append(attr)
                    except Exception:
                        pass
            except Exception:
                pass

        if busy_threads:
            print(f"[WARN][THREAD][CLOSE] still_running={busy_threads}")
            try:
                self.update_status_and_ui(
                    "백그라운드 작업 종료 중입니다. 잠시 후 다시 닫아주세요.",
                    self.center_button.isEnabled(),
                )
            except Exception:
                pass
            event.ignore()
            return

        try:
            self._save_window_state()
            try:
                flushed_favorites = self._flush_pending_favorites_save()
            except Exception:
                flushed_favorites = False
            self._flush_pending_buffer_structure_save()
            flushed_settings = self._flush_pending_settings_save()
            self._dbg_hot(
                f"[DBG][FLUSH] close favorites={flushed_favorites} settings={flushed_settings}"
            )
        except Exception as e:
            print(f"[ERR][FLUSH] Failed to save favorites on exit: {e}")
        super().closeEvent(event)

    def update_status_and_ui(self, status_text: str, is_connected: bool):
        status_text = str(status_text or "")
        is_connected = bool(is_connected)
        status_changed = self.connection_status_label.text() != status_text
        if status_changed:
            self.connection_status_label.setText(status_text)
        last_connected = getattr(self, "_last_status_is_connected", None)
        connected_changed = last_connected is None or bool(last_connected) != is_connected
        self._last_status_is_connected = is_connected

        def set_enabled_if_needed(widget) -> bool:
            if widget is None or widget.isEnabled() == is_connected:
                return False
            widget.setEnabled(is_connected)
            return True

        controls_changed = set_enabled_if_needed(self.center_button)
        search_input = getattr(self, "search_input", None)
        controls_changed = set_enabled_if_needed(search_input) or controls_changed
        search_button = getattr(self, "search_button", None)
        controls_changed = set_enabled_if_needed(search_button) or controls_changed
        if connected_changed or controls_changed:
            self._sync_connected_onenote_special_actions(is_connected)

    def _sync_connected_onenote_special_actions(self, is_connected: bool) -> None:
        open_all_busy = bool(
            self._open_all_notebooks_worker
            and not self._open_all_notebooks_worker.isFinished()
        )
        refresh_open_busy = bool(
            self._open_notebooks_refresh_worker
            and not self._open_notebooks_refresh_worker.isFinished()
        )
        open_all_enabled = is_connected
        refresh_enabled = is_connected and not refresh_open_busy
        self._sync_open_all_notebooks_action_label(
            recalculate=False,
            busy=open_all_busy,
        )
        if hasattr(self, "open_all_notebooks_action"):
            self.open_all_notebooks_action.setEnabled(open_all_enabled)
        if hasattr(self, "open_all_notebooks_button"):
            self.open_all_notebooks_button.setEnabled(open_all_enabled)
        if hasattr(self, "refresh_open_notebooks_action"):
            self.refresh_open_notebooks_action.setEnabled(refresh_enabled)
        if hasattr(self, "refresh_open_notebooks_button"):
            self.refresh_open_notebooks_button.setEnabled(refresh_enabled)

    def _sync_open_all_notebooks_action_label(
        self,
        *,
        recalculate: bool = False,
        busy: bool = False,
    ) -> None:
        if busy:
            label = "체크 없는 전자필기장 열기 중지"
            tip = "실행 중인 체크 없는 전자필기장 일괄 열기 작업을 중지합니다."
        else:
            if recalculate and getattr(self, "_open_all_candidate_count_dirty", True):
                try:
                    candidates = self._collect_open_all_notebook_candidates()
                    scope = str(
                        getattr(self, "_last_open_all_candidate_scope", "") or ""
                    )
                    self._open_all_candidate_scope = scope
                    self._open_all_candidate_count = (
                        len(candidates) if scope == "AGG_UNCHECKED" else None
                    )
                    self._open_all_candidate_stats = (
                        self._summarize_open_all_notebook_candidates(candidates)
                        if scope == "AGG_UNCHECKED"
                        else {}
                    )
                except Exception as exc:
                    print(f"[WARN][OPEN_ALL][COUNT] {exc}")
                    self._open_all_candidate_scope = ""
                    self._open_all_candidate_count = None
                    self._open_all_candidate_stats = {}
                self._open_all_candidate_count_dirty = False
            count = getattr(self, "_open_all_candidate_count", None)
            label = _open_unchecked_notebooks_button_label(count)
            tip = _open_unchecked_notebooks_tip()
            if isinstance(count, int):
                tip += f"\n현재 체크 없는 후보: {count}개"
                stats_text = self._format_open_all_candidate_stats_for_tip(
                    getattr(self, "_open_all_candidate_stats", {})
                )
                if stats_text:
                    tip += f"\n자동화 방식: {stats_text}"
        if hasattr(self, "open_all_notebooks_action"):
            self.open_all_notebooks_action.setText(label)
            self.open_all_notebooks_action.setStatusTip(tip)
        if hasattr(self, "open_all_notebooks_button"):
            self.open_all_notebooks_button.setText(label)
            self.open_all_notebooks_button.setToolTip(tip)

    def _summarize_open_all_notebook_candidates(
        self, candidates: List[Dict[str, Any]]
    ) -> Dict[str, int]:
        stats = {
            "total": 0,
            "url": 0,
            "path": 0,
            "direct": 0,
            "ui": 0,
            "missing": 0,
        }
        for record in candidates or []:
            if not str((record or {}).get("name") or "").strip():
                continue
            stats["total"] += 1
            has_url = bool(str((record or {}).get("url") or "").strip())
            has_path = bool(str((record or {}).get("path") or "").strip())
            if has_url:
                stats["url"] += 1
                stats["direct"] += 1
            elif has_path and not IS_MACOS:
                stats["path"] += 1
                stats["direct"] += 1
            elif IS_MACOS and _mac_record_is_app_only_without_launch_info(record):
                stats["ui"] += 1
            else:
                stats["ui"] += 1
        return stats

    def _format_open_all_candidate_stats_for_tip(self, stats: Dict[str, int]) -> str:
        if not isinstance(stats, dict) or int(stats.get("total") or 0) <= 0:
            return ""
        if IS_MACOS:
            return (
                f"URL 일괄 {int(stats.get('url') or 0)}개, "
                f"UI/이름 검색 {int(stats.get('ui') or 0)}개"
            )
        return (
            f"바로가기/URL 자동 {int(stats.get('direct') or 0)}개, "
            f"UI 보조 {int(stats.get('ui') or 0)}개"
        )

    def _capture_onenote_list_selection_key(self):
        item = None
        try:
            item = self.onenote_list_widget.currentItem()
        except Exception:
            item = None
        if item is not None:
            try:
                raw = item.data(Qt.ItemDataRole.UserRole)
                if isinstance(raw, dict):
                    return (
                        raw.get("handle"),
                        raw.get("pid"),
                        raw.get("title"),
                    )
            except Exception:
                pass
        return None

    def _schedule_onenote_list_auto_refresh(self, delay_ms: int = 120):
        if not hasattr(self, "onenote_list_widget"):
            return
        if self._scanner_worker and self._scanner_worker.isRunning():
            return
        now = time.monotonic()
        if (now - self._last_onenote_list_refresh_at) < 0.4:
            return
        self._pending_onenote_list_selection_key = (
            self._capture_onenote_list_selection_key()
        )
        self._onenote_list_refresh_timer.start(max(0, int(delay_ms)))

    def _cancel_pending_onenote_list_auto_refresh(self):
        if self._onenote_list_refresh_timer.isActive():
            self._onenote_list_refresh_timer.stop()

    def _refresh_onenote_list_from_click(self):
        if self._scanner_worker and self._scanner_worker.isRunning():
            return
        self._last_onenote_list_refresh_at = time.monotonic()
        self.refresh_onenote_list()

    def _current_onenote_handle(self) -> Optional[int]:
        win = getattr(self, "onenote_window", None)
        if win is None:
            return None
        try:
            handle = getattr(win, "handle", None)
            if callable(handle):
                handle = handle()
            if handle:
                return int(handle)
        except Exception:
            return None
        return None

    def _coerce_macos_window(self, window: Optional[object] = None) -> Optional[MacWindow]:
        if not IS_MACOS:
            return None
        candidate = window or getattr(self, "onenote_window", None)
        if candidate is None:
            return None
        if isinstance(candidate, MacWindow):
            info = dict(getattr(candidate, "info", {}) or {})
            if info:
                try:
                    resolved = resolve_window_target(info)
                    if isinstance(resolved, MacWindow):
                        return resolved
                except Exception:
                    pass
            return candidate
        try:
            info = getattr(candidate, "info", None)
        except Exception:
            info = None
        if isinstance(info, dict) and info:
            try:
                return MacWindow(dict(info))
            except Exception:
                pass

        rebuilt: Dict[str, Any] = {}
        for key, attr_name in (
            ("handle", "handle"),
            ("pid", "process_id"),
            ("title", "window_text"),
            ("bundle_id", "bundle_id"),
            ("class_name", "class_name"),
            ("app_name", "app_name"),
        ):
            try:
                value = getattr(candidate, attr_name)
                value = value() if callable(value) else value
            except Exception:
                value = None
            if value not in (None, "", 0):
                rebuilt[key] = value
        if rebuilt:
            try:
                return MacWindow(rebuilt)
            except Exception:
                pass

        try:
            sig = build_window_signature_quick(candidate)
        except Exception:
            sig = {}
        if isinstance(sig, dict) and sig:
            try:
                resolved = resolve_window_target(sig)
                if isinstance(resolved, MacWindow):
                    return resolved
            except Exception:
                pass
            try:
                return MacWindow(dict(sig))
            except Exception:
                pass
        return None

_publish_context(globals())

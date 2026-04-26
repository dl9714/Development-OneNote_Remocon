# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin28:

    def _refresh_macos_onenote_main_window(self) -> Optional[MacWindow]:
        if not IS_MACOS:
            return None

        current_info = dict(getattr(getattr(self, "onenote_window", None), "info", {}) or {})
        try:
            preferred_pid = int(current_info.get("pid") or 0)
        except Exception:
            preferred_pid = 0
        try:
            preferred_window_number = int(current_info.get("window_number") or 0)
        except Exception:
            preferred_window_number = 0

        try:
            candidates = [
                dict(info)
                for info in enumerate_macos_windows()
                if is_macos_onenote_window_info(info, os.getpid())
            ]
        except Exception:
            candidates = []
        if not candidates:
            return self._coerce_macos_window(getattr(self, "onenote_window", None))

        dialog_tokens = (
            "삽입할 파일 선택",
            "최근 전자 필기장",
            "최근 전자필기장",
            "새 전자 필기장",
            "새 전자필기장",
            "choose a file",
            "insert file",
            "recent notebook",
            "new notebook",
        )

        def _is_dialog_candidate(info: Dict[str, Any]) -> bool:
            title = str(info.get("title") or "").lower()
            return any(token.lower() in title for token in dialog_tokens)

        main_candidates = [info for info in candidates if not _is_dialog_candidate(info)]
        ranked_candidates = main_candidates or candidates

        def _score(info: Dict[str, Any]) -> int:
            score = 0
            try:
                if preferred_window_number and int(info.get("window_number") or 0) == preferred_window_number:
                    score += 300
            except Exception:
                pass
            try:
                if preferred_pid and int(info.get("pid") or 0) == preferred_pid:
                    score += 90
            except Exception:
                pass
            if str(info.get("bundle_id") or "") == ONENOTE_MAC_BUNDLE_ID:
                score += 80
            if not _is_dialog_candidate(info):
                score += 60
            if info.get("frontmost"):
                score += 30
            title = str(info.get("title") or "").strip()
            if title:
                score += 20
            if title and title.lower() != "onenote":
                score += 10
            return score

        try:
            best = max(ranked_candidates, key=_score)
        except ValueError:
            return self._coerce_macos_window(getattr(self, "onenote_window", None))

        try:
            win = MacWindow(dict(best))
        except Exception:
            return self._coerce_macos_window(getattr(self, "onenote_window", None))
        self.onenote_window = win
        try:
            save_connection_info(win)
        except Exception:
            pass
        return win

    def _mac_selected_outline_context(
        self, window: Optional[object] = None
    ) -> Dict[str, str]:
        if not IS_MACOS:
            return {}
        win = self._coerce_macos_window(window)
        if win is None:
            return {}
        try:
            context = mac_current_outline_context(win)
            if any(context.get(key) for key in ("notebook", "section", "page")):
                return context
            try:
                win.set_focus()
            except Exception:
                pass
            return mac_current_outline_context(win)
        except Exception as e:
            print(f"[WARN] 맥 현재 위치 조회 실패: {e}")
            return {}

    def _restore_macos_page_context(self, page_text: str) -> bool:
        if not IS_MACOS:
            return False
        text = str(page_text or "").strip()
        win = self._coerce_macos_window(getattr(self, "onenote_window", None))
        if not text or win is None:
            return False
        try:
            return mac_select_page_row_by_text(win, text)
        except Exception as e:
            print(f"[WARN] 맥 페이지 복구 실패: {e}")
            return False

    def _is_sig_same_as_connected_window(self, sig: Dict[str, Any]) -> bool:
        if not sig or not getattr(self, "onenote_window", None):
            return False

        current_handle = self._current_onenote_handle()
        try:
            target_handle = int(sig.get("handle") or 0)
        except Exception:
            target_handle = 0
        if current_handle and target_handle and current_handle == target_handle:
            return True

        try:
            current_sig = build_window_signature(self.onenote_window)
        except Exception:
            current_sig = {}

        current_pid = current_sig.get("pid")
        target_pid = sig.get("pid")
        current_class = current_sig.get("class_name") or ""
        target_class = sig.get("class_name") or ""
        current_exe = current_sig.get("exe_name") or ""
        target_exe = sig.get("exe_name") or ""
        if (
            current_pid
            and target_pid
            and current_pid == target_pid
            and (not current_class or not target_class or current_class == target_class)
            and (not current_exe or not target_exe or current_exe == target_exe)
        ):
            return True

        try:
            if current_handle and len(self.onenote_windows_info or []) == 1:
                return True
        except Exception:
            pass

        return False

    def _try_activate_favorite_fastpath(
        self,
        item: QTreeWidgetItem,
        sig: Dict[str, Any],
        target: Dict[str, Any],
        display_name: str,
        *,
        started_at: Optional[float] = None,
    ) -> bool:
        return self._try_activate_favorite_fastpath_v2(
            item,
            sig,
            target,
            display_name,
            started_at=started_at,
        )
        target_info = _resolve_favorite_activation_target(target, display_name)
        if not target_info.get("ok", True):
            self.update_status_and_ui(
                target_info.get("error") or "즐겨찾기 대상을 찾지 못했습니다.",
                self.center_button.isEnabled(),
            )
            print(
                "[DBG][FAV][FASTPATH]",
                "resolve_abort",
                f"error={target_info.get('error')!r}",
            )
            return True

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
        if not tree:
            return False

        target_kind = target_info.get("target_kind")
        expected_text = target_info.get("expected_center_text") or ""
        print(
            "[DBG][FAV][FASTPATH]",
            direct_source,
            f"kind={target_kind}",
            f"text={expected_text!r}",
            f"elapsed_ms={(time.perf_counter() - (started_at or time.perf_counter())) * 1000.0:.1f}",
        )

        try:
            self.onenote_window.set_focus()
        except Exception:
            pass

        ok = False
        if target_kind == "notebook":
            ok = select_notebook_by_text(
                self.onenote_window,
                expected_text,
                tree,
                center_after_select=False,
            )
        elif target_kind == "section":
            ok = select_section_by_text(self.onenote_window, expected_text, tree)

        if not ok:
            self._cache_tree_control()
            tree = self.tree_control
            if tree:
                if target_kind == "notebook":
                    ok = select_notebook_by_text(
                        self.onenote_window,
                        expected_text,
                        tree,
                        center_after_select=False,
                    )
                elif target_kind == "section":
                    ok = select_section_by_text(
                        self.onenote_window, expected_text, tree
                    )

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
                target_info.get("resolved_name") or "",
                target_info.get("resolved_notebook_id") or "",
            )

        self.center_selected_item_action(
            debug_source="fav_fastpath",
            started_at=started_at,
        )
        return True

_publish_context(globals())

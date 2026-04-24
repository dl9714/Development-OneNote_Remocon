# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


class OpenAllNotebooksWorkerMacCollectMixin:

    def _collect_macos_open_all_sources(self, result, win):
        candidate_limited_mode = self.candidate_scope in (
            "AGG_UNCHECKED",
            "AGG_UNCLASSIFIED",
        )
        mac_open_action_label = (
            _open_unchecked_notebooks_button_label()
            if candidate_limited_mode
            else "실제 OneNote 전체 열기"
        )

        def _mac_open_all_debug(*parts: Any) -> None:
            line = " ".join(str(part) for part in parts)
            print(line)
            _append_open_all_debug_log(line)

        _mac_open_all_debug("[DBG][OPEN_ALL][MAC] branch-start")
        _mac_open_all_debug(
            "[DBG][OPEN_ALL][MAC]",
            f"resolved-window-title={_preferred_connected_window_title_quick(win, self.sig)!r}",
            f"pid={getattr(win, 'process_id', lambda: 0)() if hasattr(win, 'process_id') else 0}",
        )
        self.progress.emit(f"{mac_open_action_label} 준비 중... 현재 열린 목록 확인")
        initial_notebook_name = _preferred_connected_window_title_quick(
            win,
            self.sig,
        )
        open_names: Set[str] = set()
        initial_notebook_key = _normalize_notebook_name_key(initial_notebook_name)
        if initial_notebook_key:
            open_names.add(initial_notebook_key)
        open_detected_names: List[str] = []
        open_detect_debug: Dict[str, Any] = {}
        sidebar_error = ""
        accessibility_trusted = macos_accessibility_is_trusted()

        if accessibility_trusted:
            quick_open_snapshot = mac_current_open_notebook_names_quick(
                win,
                ax_timeout_sec=2.2,
                plist_timeout_sec=1.2,
                sidebar_timeout_sec=0.6,
                min_names_before_sidebar=8,
            )
            open_detected_names = [
                str(name or "").strip()
                for name in (quick_open_snapshot.get("names") or [])
                if str(name or "").strip()
            ]
            open_detect_debug = dict(quick_open_snapshot.get("debug") or {})
            sidebar_error = str(open_detect_debug.get("sidebar_error") or "")
            open_names.update(
                _normalize_notebook_name_key(name)
                for name in open_detected_names
                if _normalize_notebook_name_key(name)
            )
            _mac_open_all_debug(
                "[DBG][OPEN_ALL][MAC]",
                f"open-detected={len(open_detected_names)}",
                f"title={int(open_detect_debug.get('title_count') or 0)}",
                f"ax={int(open_detect_debug.get('ax_count') or 0)}",
                f"plist={int(open_detect_debug.get('plist_count') or 0)}",
                f"sidebar={int(open_detect_debug.get('sidebar_count') or 0)}",
                "timeouts="
                f"ax:{int(bool(open_detect_debug.get('ax_timed_out')))}"
                f"/plist:{int(bool(open_detect_debug.get('plist_timed_out')))}"
                f"/sidebar:{int(bool(open_detect_debug.get('sidebar_timed_out')))}",
                f"sidebar-error={sidebar_error!r}",
                f"open-sample={open_detected_names[:8]!r}",
            )
            if int(open_detect_debug.get("ax_count") or 0) <= 0:
                _mac_open_all_debug(
                    "[DBG][OPEN_ALL][MAC]",
                    f"ax-debug={macos_last_ax_notebook_debug()!r}",
                )
        else:
            _mac_open_all_debug("[DBG][OPEN_ALL][MAC] accessibility trusted=false")

        self.progress.emit(f"{mac_open_action_label} 준비 중... 최근 전자필기장 캐시 확인")
        recent_box: Dict[str, Any] = {}
        recent_done = threading.Event()

        def _read_recent_records() -> None:
            try:
                recent_box["records"] = [
                    dict(record)
                    for record in mac_recent_notebook_records(None)
                    if str((record or {}).get("name") or "").strip()
                ]
            except Exception as exc:
                recent_box["error"] = exc
            finally:
                recent_done.set()

        threading.Thread(
            target=_read_recent_records,
            name="onenote-recent-cache-scan",
            daemon=True,
        ).start()
        if recent_done.wait(1.5):
            if "error" in recent_box:
                print(
                    "[WARN][OPEN_ALL_NOTEBOOKS][MAC][RECENT_CACHE]",
                    str(recent_box["error"]),
                )
                recent_records = []
            else:
                recent_records = [
                    dict(record)
                    for record in (recent_box.get("records") or [])
                    if str((record or {}).get("name") or "").strip()
                ]
        else:
            print(
                "[WARN][OPEN_ALL_NOTEBOOKS][MAC][RECENT_CACHE] timed out after 1.5s"
            )
            recent_records = []
        _mac_open_all_debug(
            f"[DBG][OPEN_ALL][MAC] recent-records={len(recent_records)}"
        )
        for record in recent_records:
            if not str(record.get("source") or "").strip():
                record["source"] = "MAC_RECENT_CACHE"
        self.progress.emit(f"{mac_open_action_label} 준비 중... OneDrive 바로가기 확인")
        shortcut_box: Dict[str, Any] = {}
        shortcut_done = threading.Event()

        def _read_shortcut_records() -> None:
            try:
                shortcut_box["records"] = [
                    dict(record)
                    for record in _collect_onenote_notebook_shortcuts()
                    if str((record or {}).get("name") or "").strip()
                ]
            except Exception as exc:
                shortcut_box["error"] = exc
            finally:
                shortcut_done.set()

        threading.Thread(
            target=_read_shortcut_records,
            name="onenote-shortcut-scan",
            daemon=True,
        ).start()
        if shortcut_done.wait(2.0):
            if "error" in shortcut_box:
                print(
                    "[WARN][OPEN_ALL_NOTEBOOKS][MAC][SHORTCUTS]",
                    str(shortcut_box["error"]),
                )
                shortcut_records = []
            else:
                shortcut_records = [
                    dict(record)
                    for record in (shortcut_box.get("records") or [])
                    if str((record or {}).get("name") or "").strip()
                ]
        else:
            print(
                "[WARN][OPEN_ALL_NOTEBOOKS][MAC][SHORTCUTS] timed out after 2s"
            )
            shortcut_records = []
        _mac_open_all_debug(
            "[DBG][OPEN_ALL][MAC]",
            f"shortcut-records={len(shortcut_records)}",
        )
        for record in shortcut_records:
            if not str(record.get("source") or "").strip():
                record["source"] = "MAC_SHORTCUT"
        self.progress.emit(f"{mac_open_action_label} 준비 중... 저장된 후보 병합")
        settings_records = [
            dict(record)
            for record in self.notebook_candidates
            if str((record or {}).get("name") or "").strip()
        ]
        _mac_open_all_debug(
            "[DBG][OPEN_ALL][MAC]",
            f"settings-records={len(settings_records)}",
            f"settings-sample={[str((record or {}).get('name') or '').strip() for record in settings_records[:8]]!r}",
        )
        open_tab_records: List[Dict[str, Any]] = []
        if (
            candidate_limited_mode
            and settings_records
            and accessibility_trusted
        ):
            self.progress.emit(
                f"{mac_open_action_label} 준비 중... OneNote 열기 탭 전체 목록"
            )
            open_tab_box: Dict[str, Any] = {}
            open_tab_done = threading.Event()

            def _read_open_tab_records() -> None:
                try:
                    open_tab_box["records"] = [
                        dict(record)
                        for record in mac_open_tab_notebook_records(win, fast=True)
                        if str((record or {}).get("name") or "").strip()
                    ]
                except Exception as exc:
                    open_tab_box["error"] = exc
                finally:
                    open_tab_done.set()

            threading.Thread(
                target=_read_open_tab_records,
                name="onenote-open-tab-fast-scan",
                daemon=True,
            ).start()
            if open_tab_done.wait(18.0):
                if "error" in open_tab_box:
                    _mac_open_all_debug(
                        "[WARN][OPEN_ALL_NOTEBOOKS][MAC][OPEN_TAB]",
                        str(open_tab_box["error"]),
                    )
                else:
                    open_tab_records = [
                        dict(record)
                        for record in (open_tab_box.get("records") or [])
                        if str((record or {}).get("name") or "").strip()
                    ]
                    for record in open_tab_records:
                        if not str(record.get("source") or "").strip():
                            record["source"] = "MAC_OPEN_TAB"
            else:
                _mac_open_all_debug(
                    "[WARN][OPEN_ALL_NOTEBOOKS][MAC][OPEN_TAB] timed out after 18s"
                )
            _mac_open_all_debug(
                "[DBG][OPEN_ALL][MAC]",
                f"open-tab-records={len(open_tab_records)}",
                f"open-tab-sample={[str((record or {}).get('name') or '').strip() for record in open_tab_records[:8]]!r}",
            )
        if (
            candidate_limited_mode
            and settings_records
            and not recent_records
            and accessibility_trusted
        ):
            self.progress.emit(
                f"{mac_open_action_label} 준비 중... 최근 목록 UI 빠른 스냅샷"
            )
            dialog_recent_box: Dict[str, Any] = {}
            dialog_recent_done = threading.Event()

            def _read_dialog_recent_records() -> None:
                try:
                    dialog_recent_box["records"] = [
                        dict(record)
                        for record in mac_recent_notebook_records(win, fast=True)
                        if str((record or {}).get("name") or "").strip()
                    ]
                except Exception as exc:
                    dialog_recent_box["error"] = exc
                finally:
                    dialog_recent_done.set()

            threading.Thread(
                target=_read_dialog_recent_records,
                name="onenote-recent-dialog-fast-scan",
                daemon=True,
            ).start()
            if dialog_recent_done.wait(10.0):
                if "error" in dialog_recent_box:
                    _mac_open_all_debug(
                        "[WARN][OPEN_ALL_NOTEBOOKS][MAC][RECENT_DIALOG]",
                        str(dialog_recent_box["error"]),
                    )
                else:
                    recent_records = [
                        dict(record)
                        for record in (dialog_recent_box.get("records") or [])
                        if str((record or {}).get("name") or "").strip()
                    ]
                    for record in recent_records:
                        if not str(record.get("source") or "").strip():
                            record["source"] = "MAC_RECENT_DIALOG"
            else:
                _mac_open_all_debug(
                    "[WARN][OPEN_ALL_NOTEBOOKS][MAC][RECENT_DIALOG] timed out after 10s"
                )
            _mac_open_all_debug(
                "[DBG][OPEN_ALL][MAC]",
                f"recent-dialog-records={len(recent_records)}",
                f"recent-dialog-sample={[str((record or {}).get('name') or '').strip() for record in recent_records[:8]]!r}",
            )

        return {
            "candidate_limited_mode": candidate_limited_mode,
            "mac_open_action_label": mac_open_action_label,
            "_mac_open_all_debug": _mac_open_all_debug,
            "initial_notebook_name": initial_notebook_name,
            "open_names": open_names,
            "open_detected_names": open_detected_names,
            "open_detect_debug": open_detect_debug,
            "sidebar_error": sidebar_error,
            "accessibility_trusted": accessibility_trusted,
            "recent_records": recent_records,
            "shortcut_records": shortcut_records,
            "settings_records": settings_records,
            "open_tab_records": open_tab_records,
        }

_publish_context(globals())

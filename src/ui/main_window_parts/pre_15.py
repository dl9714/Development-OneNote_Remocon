# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


class OpenAllNotebooksWorker(
    OpenAllNotebooksWorkerMacCollectMixin,
    OpenAllNotebooksWorkerMacPrepareMixin,
    OpenAllNotebooksWorkerMacLaunchMixin,
    OpenAllNotebooksWorkerMacMixin,
    OpenAllNotebooksWorkerWindowsMixin,
    QThread,
):
    progress = pyqtSignal(str)
    done = pyqtSignal(dict)

    def __init__(
        self,
        sig: Dict[str, Any],
        notebook_candidates: Optional[List[Dict[str, Any]]] = None,
        candidate_scope: str = "",
        parent=None,
    ):
        super().__init__(parent)
        self.sig = copy.deepcopy(sig or {})
        self.candidate_scope = str(candidate_scope or "").strip()
        self.notebook_candidates = [
            dict(record)
            for record in (notebook_candidates or [])
            if str((record or {}).get("name") or "").strip()
        ]

    def run(self):
        result = {
            "ok": False,
            "window_info": None,
            "opened_count": 0,
            "opened_names": [],
            "remaining_names": [],
            "error": "",
            "verified_open_count": 0,
            "candidate_scope": self.candidate_scope,
            "candidate_total": 0,
            "pending_count": 0,
            "already_open_count": 0,
        }
        try:
            ensure_pywinauto()
            if not _pwa_ready:
                result["error"] = "자동화 모듈이 로드되지 않았습니다."
                self.done.emit(result)
                return

            win = reacquire_window_by_signature(self.sig)
            if not win:
                result["error"] = "연결된 OneNote 창을 다시 찾지 못했습니다."
                self.done.emit(result)
                return
            if IS_MACOS:
                resolved_mac_win = _resolve_macos_primary_notebook_window(win, self.sig)
                if resolved_mac_win is not None:
                    win = resolved_mac_win

            try:
                result["window_info"] = {
                    "handle": win.handle,
                    "title": _preferred_connected_window_title_quick(win, self.sig),
                    "class_name": win.class_name(),
                    "pid": win.process_id(),
                }
            except Exception:
                result["window_info"] = None

            if IS_MACOS:
                self._run_macos_open_all(result, win)
                return

            self._run_windows_open_all(result, win)
        except Exception as e:
            print(f"[WARN][OPEN_ALL_NOTEBOOKS][WORKER] {e}")
            result["error"] = str(e)
            self.done.emit(result)

_publish_context(globals())

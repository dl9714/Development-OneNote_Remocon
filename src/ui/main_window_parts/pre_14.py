# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class FavoriteActivationWorker(QThread):
    done = pyqtSignal(dict)
    def __init__(
        self,
        sig: Dict[str, Any],
        target: Dict[str, Any],
        display_name: str,
        auto_center_after_activate: bool,
        parent=None,
    ):
        super().__init__(parent)
        self.sig = copy.deepcopy(sig or {})
        self.target = copy.deepcopy(target or {})
        self.display_name = display_name or ""
        self.auto_center_after_activate = bool(auto_center_after_activate)
    def run(self):
        result = {
            "ok": False,
            "display_name": self.display_name,
            "window_info": None,
            "target_kind": None,
            "expected_center_text": "",
            "resolved_name": "",
            "resolved_notebook_id": "",
            "error": "",
        }
        try:
            ensure_pywinauto()
            if not IS_MACOS and not _pwa_ready:
                result["error"] = "자동화 모듈이 로드되지 않았습니다."
                self.done.emit(result)
                return
            win = reacquire_window_by_signature(self.sig)
            if not win:
                result["error"] = f"대상 창 '{self.display_name}'을(를) 찾을 수 없습니다."
                self.done.emit(result)
                return
            try:
                result["window_info"] = {
                    "handle": win.handle,
                    "title": win.window_text(),
                    "class_name": win.class_name(),
                    "pid": win.process_id(),
                }
            except Exception:
                result["window_info"] = None
            if not self.auto_center_after_activate:
                result["ok"] = True
                self.done.emit(result)
                return
            exe_name = (self.sig.get("exe_name") or "").lower()
            sig_title = (self.sig.get("title") or "").lower()
            if "onenote" not in exe_name and "onenote" not in sig_title:
                result["ok"] = True
                self.done.emit(result)
                return
            tree = _find_tree_or_list(win)
            if not tree and not IS_MACOS:
                result["error"] = "OneNote 트리를 찾지 못했습니다."
                self.done.emit(result)
                return
            target_info = _resolve_favorite_activation_target(
                self.target, self.display_name
            )
            result["target_kind"] = target_info.get("target_kind")
            result["expected_center_text"] = (
                target_info.get("expected_center_text") or ""
            )
            result["resolved_name"] = target_info.get("resolved_name") or ""
            result["resolved_notebook_id"] = (
                target_info.get("resolved_notebook_id") or ""
            )
            if not target_info.get("ok", True):
                result["error"] = target_info.get("error") or ""
                self.done.emit(result)
                return
            if IS_MACOS and result["target_kind"] == "notebook":
                requested_name = (
                    result["expected_center_text"]
                    or result["resolved_name"]
                    or str((self.target or {}).get("notebook_text") or "")
                    or self.display_name
                )
                notebook_result = _mac_ensure_notebook_context_for_section(
                    win,
                    requested_name,
                    wait_for_visible=False,
                )
                if not notebook_result.get("ok", True):
                    result["error"] = notebook_result.get("error") or ""
                    self.done.emit(result)
                    return
                final_name = str(notebook_result.get("name") or requested_name).strip()
                result["ok"] = True
                result["resolved_name"] = final_name
                result["expected_center_text"] = final_name
                self.done.emit(result)
                return
            if IS_MACOS and result["target_kind"] == "section":
                notebook_text = str(
                    target_info.get("expected_notebook_text")
                    or (self.target or {}).get("notebook_text")
                    or ""
                ).strip()
                if notebook_text:
                    notebook_result = _mac_ensure_notebook_context_for_section(
                        win,
                        notebook_text,
                    )
                    if not notebook_result.get("ok", True):
                        result["error"] = notebook_result.get("error") or ""
                        self.done.emit(result)
                        return
            ok = False
            if result["target_kind"] == "notebook":
                ok = select_notebook_by_text(
                    win,
                    result["expected_center_text"],
                    tree,
                    center_after_select=False,
                )
            elif result["target_kind"] == "section":
                ok = select_section_by_text(
                    win, result["expected_center_text"], tree
                )
            else:
                ok = select_notebook_by_text(
                    win, result["expected_center_text"], tree, center_after_select=False
                )
            if not ok:
                try:
                    win.set_focus()
                except Exception:
                    pass
                if result["target_kind"] == "notebook":
                    ok = select_notebook_by_text(
                        win, result["expected_center_text"], tree, center_after_select=False
                    )
                elif result["target_kind"] == "section":
                    ok = select_section_by_text(win, result["expected_center_text"], tree)
            result["ok"] = ok
            self.done.emit(result)
        except Exception as e:
            print(f"[WARN][ACTIVATE][WORKER] {e}")
            result["error"] = str(e)
            self.done.emit(result)

_publish_context(globals())

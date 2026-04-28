# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


FavoriteActivationWorker = lazy_class(
    "src.ui.main_window_parts.pre_14", "FavoriteActivationWorker"
)
OpenNotebookRecordsWorker = lazy_class(
    "src.ui.main_window_parts.pre_16", "OpenNotebookRecordsWorker"
)
CodexLocationLookupWorker = lazy_class(
    "src.ui.main_window_parts.pre_16", "CodexLocationLookupWorker"
)
OtherWindowSelectionDialog = lazy_class(
    "src.ui.main_window_parts.pre_16", "OtherWindowSelectionDialog"
)


def _score_candidate_dict(c, sig) -> int:
    try:
        title = (c.get("title") or "").lower()
        cls = c.get("class_name") or ""
        pid = c.get("pid")
        if IS_MACOS:
            exe_path = ""
            exe_name = os.path.basename(str(c.get("bundle_id") or cls or "")).lower()
        else:
            exe_path = get_process_image_path(pid) or ""
            exe_name = os.path.basename(exe_path).lower() if exe_path else ""

        score = 0
        if sig.get("handle") and c.get("handle") == sig["handle"]:
            score += 100
        if sig.get("exe_name") and exe_name == sig["exe_name"]:
            score += 50
        if IS_MACOS and str(c.get("bundle_id") or "") == ONENOTE_MAC_BUNDLE_ID:
            score += 50
        elif "onenote.exe" in exe_name:
            score += 50
        if "onenote" in title or "원노트" in title:
            score += 25
        if sig.get("class_name") and cls == sig["class_name"]:
            score += 10
        if sig.get("pid") and pid == sig["pid"]:
            score += 8
        prev_title = (sig.get("title") or "").lower()
        if prev_title:
            if prev_title in title:
                score += 6
            else:
                if "onenote" in prev_title and "onenote" in title:
                    score += 4
                if "원노트" in prev_title and "원노트" in title:
                    score += 4
        if cls == "Framework::CFrame":
            score += 8
        elif cls == ONENOTE_CLASS_NAME or (IS_MACOS and str(c.get("bundle_id") or "") == ONENOTE_MAC_BUNDLE_ID):
            score += 5
        return score
    except Exception:
        return -1


def _windows_onenote_class_sort_key(info: Dict[str, Any]) -> int:
    class_name = str((info or {}).get("class_name") or "").casefold()
    if class_name == "framework::cframe" or "omain" in class_name:
        return 0
    if class_name == str(ONENOTE_CLASS_NAME).casefold():
        return 1
    return 2


def _signature_looks_like_windows_onenote(sig: Optional[Dict[str, Any]]) -> bool:
    if not IS_WINDOWS or not isinstance(sig, dict):
        return False
    title = str(sig.get("title") or "").casefold()
    class_name = str(sig.get("class_name") or "").casefold()
    exe_name = str(sig.get("exe_name") or "").casefold()
    return (
        "onenote" in title
        or "원노트" in title
        or "onenote" in exe_name
        or "onenoteim" in exe_name
        or "omain" in class_name
        or class_name == "framework::cframe"
    )


def _window_info_from_wrapper(win) -> Dict[str, Any]:
    try:
        handle = int(getattr(win, "handle", 0) or 0)
    except Exception:
        handle = 0
    try:
        title = win.window_text() or ""
    except Exception:
        title = ""
    try:
        class_name = win.class_name() or ""
    except Exception:
        class_name = ""
    try:
        pid = win.process_id()
    except Exception:
        pid = 0
    return {
        "handle": handle,
        "title": title,
        "class_name": class_name,
        "pid": pid,
    }


def reacquire_window_by_signature(sig) -> Optional[object]:
    ensure_pywinauto()
    if not IS_MACOS and not _pwa_ready:
        return None
    h = sig.get("handle")
    if IS_WINDOWS and h:
        try:
            w = Desktop(backend="uia").window(handle=h)
            if _signature_looks_like_windows_onenote(sig):
                info = _window_info_from_wrapper(w)
                if is_strict_onenote_window(info, os.getpid()):
                    return w
            elif w.is_visible():
                return w
        except Exception:
            pass

    candidates = (
        enumerate_macos_windows_quick(filter_title_substr=None)
        if IS_MACOS
        else enum_windows_fast(filter_title_substr=None)
    )
    if IS_WINDOWS and _signature_looks_like_windows_onenote(sig):
        candidates = [
            c
            for c in candidates
            if is_strict_onenote_window(c, os.getpid())
        ]
    if IS_MACOS:
        exact = None
        if h:
            for candidate in candidates:
                if int(candidate.get("handle") or 0) == int(h):
                    exact = candidate
                    break
        if exact:
            return MacWindow(dict(exact))
    best, best_score = None, -1
    for c in candidates:
        s = _score_candidate_dict(c, sig)
        if s > best_score:
            best, best_score = c, s

    if best and best_score >= 30:
        try:
            if IS_MACOS:
                return MacWindow(dict(best))
            w = Desktop(backend="uia").window(handle=best["handle"])
            if _signature_looks_like_windows_onenote(sig):
                info = _window_info_from_wrapper(w)
                if is_strict_onenote_window(info, os.getpid()):
                    return w
            elif w.is_visible():
                return w
        except Exception:
            return None
    return None


def resolve_window_target(sig: Dict[str, Any]) -> Optional[object]:
    ensure_pywinauto()
    if not sig:
        return None
    if IS_MACOS:
        resolved = reacquire_window_by_signature(sig)
        if resolved is not None:
            return resolved
        return MacWindow(dict(sig))

    handle = sig.get("handle")
    if handle:
        try:
            target = Desktop(backend="uia").window(handle=handle)
            if _signature_looks_like_windows_onenote(sig):
                info = _window_info_from_wrapper(target)
                if is_strict_onenote_window(info, os.getpid()):
                    return target
            elif target.is_visible():
                return target
        except Exception:
            pass
    return reacquire_window_by_signature(sig)


def ensure_windows_onenote_ready_for_tree_lookup(win) -> bool:
    if not IS_WINDOWS or win is None:
        return False

    rect = _safe_rectangle(win)
    should_restore = True
    if rect is not None:
        left = int(getattr(rect, "left", 0) or 0)
        top = int(getattr(rect, "top", 0) or 0)
        width = int(getattr(rect, "right", 0) or 0) - left
        height = int(getattr(rect, "bottom", 0) or 0) - top
        should_restore = width <= 0 or height <= 0 or left <= -30000 or top <= -30000

    if should_restore:
        try:
            win.restore()
        except Exception:
            pass
        try:
            win.set_focus()
        except Exception:
            pass
    return True


# ----------------- 12. 저장된 정보로 재연결 -----------------
def load_connection_info_and_reconnect():
    ensure_pywinauto()
    settings = load_settings()
    sig = settings.get("connection_signature")
    if not sig:
        return None, "연결되지 않음"
    try:
        win = reacquire_window_by_signature(sig)
        win_is_ready = False
        if win:
            try:
                win_is_ready = bool(win.is_visible())
            except Exception:
                win_is_ready = False
            if (
                not win_is_ready
                and IS_WINDOWS
                and _signature_looks_like_windows_onenote(sig)
            ):
                win_is_ready = is_strict_onenote_window(
                    _window_info_from_wrapper(win),
                    os.getpid(),
                )

        if win and win_is_ready:
            window_title = _preferred_connected_window_title(win, sig)
            try:
                save_connection_info(win)
            except Exception:
                pass
            return win, f"(자동 재연결) '{window_title}'"

        return None, "(재연결 실패) 이전 앱을 찾을 수 없습니다."
    except Exception:
        return None, "연결되지 않음"


# ----------------- 13. 백그라운드 자동 재연결 워커 -----------------
class ReconnectWorker(QThread):
    finished = pyqtSignal(object)

    def run(self):
        try:
            ensure_pywinauto()
            if IS_MACOS:
                settings = load_settings()
                sig = settings.get("connection_signature") or {}
                win = reacquire_window_by_signature(sig) if sig else None
                if win:
                    next_sig = build_window_signature_quick(win, sig)
                    title = _preferred_connected_window_title_quick(win, next_sig)
                    payload = {
                        "ok": True,
                        "status": f"(자동 재연결) '{title}'",
                        "sig": next_sig,
                        "target_info": dict(_window_info_dict(win) or next_sig),
                    }
                else:
                    payload = {
                        "ok": False,
                        "status": "(재연결 실패) 이전 앱을 찾을 수 없습니다.",
                    }
                self.finished.emit(payload)
                return

            win, status = load_connection_info_and_reconnect()
            if win:
                payload = {
                    "ok": True,
                    "status": status,
                    "sig": _build_connection_signature_for_save(
                        win,
                        load_settings().get("connection_signature"),
                    ),
                }
            else:
                payload = {"ok": False, "status": status}
        except Exception as e:
            payload = {"ok": False, "status": f"연결되지 않음 (오류: {e})"}
        self.finished.emit(payload)


class WindowsTreeWarmWorker(QThread):
    done = pyqtSignal(object)

    def __init__(self, sig: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.sig = copy.deepcopy(sig or {})
        self.cached_tree = None
        self.selected_name = ""

    def run(self):
        payload = {"ok": False, "tree": None, "handle": 0, "selected_name": ""}
        try:
            ensure_pywinauto()
            if not IS_WINDOWS or not _pwa_ready or not self.sig:
                self.done.emit(payload)
                return

            win = reacquire_window_by_signature(self.sig)
            if not win or self.isInterruptionRequested():
                self.done.emit(payload)
                return

            ensure_windows_onenote_ready_for_tree_lookup(win)
            if self.isInterruptionRequested():
                self.done.emit(payload)
                return

            tree = _find_tree_or_list(win)
            if not tree or self.isInterruptionRequested():
                self.done.emit(payload)
                return

            selected_item = get_selected_tree_item_fast(tree)
            selected_name = ""
            if selected_item is not None:
                try:
                    selected_name = selected_item.window_text()
                except Exception:
                    selected_name = ""

            self.cached_tree = tree
            self.selected_name = selected_name
            payload = {
                "ok": True,
                "tree": tree,
                "handle": _window_handle_cache_key(win),
                "selected_name": selected_name,
            }
        except Exception as e:
            payload = {"ok": False, "tree": None, "handle": 0, "selected_name": "", "error": str(e)}
        self.done.emit(payload)


# ----------------- 3-A. OneNote 창 목록 스캔 워커 -----------------
class OneNoteWindowScanner(QThread):
    done = pyqtSignal(list)

    def __init__(self, my_pid: int, parent=None):
        super().__init__(parent)
        self.my_pid = my_pid

    def run(self):
        results = []
        try:
            wins = (
                enumerate_macos_windows_quick(filter_title_substr=None)
                if IS_MACOS
                else enum_windows_fast(filter_title_substr=None)
            )
            for w in wins:
                try:
                    if is_strict_onenote_window(w, self.my_pid):
                        results.append(w)
                except Exception:
                    continue

            results.sort(
                key=lambda r: (
                    _windows_onenote_class_sort_key(r),
                    r.get("title", ""),
                )
            )
        except Exception as e:
            print(f"[ERROR] OneNote 창 스캔 중 오류: {e}")
        finally:
            self.done.emit(results)


# ----------------- 3-B/C. 기타 창 스캔 및 선택 다이얼로그 -----------------
class WindowListWorker(QThread):
    done = pyqtSignal(list)

    def run(self):
        try:
            results = enum_windows_fast(filter_title_substr=None)
            self.done.emit(results)
        except Exception:
            self.done.emit([])


class CenterAfterActivateWorker(QThread):
    done = pyqtSignal(bool, str)

    def __init__(
        self,
        sig: Dict[str, Any],
        expected_text: str,
        target_kind: str = "",
        parent=None,
    ):
        super().__init__(parent)
        self.sig = copy.deepcopy(sig or {})
        self.expected_text = expected_text or ""
        self.target_kind = (target_kind or "").strip().lower()
        self.cached_tree = None

    def run(self):
        try:
            ensure_pywinauto()
            if not IS_MACOS and not _pwa_ready:
                self.done.emit(False, "")
                return

            win = reacquire_window_by_signature(self.sig)
            if not win:
                self.done.emit(False, "")
                return

            if IS_WINDOWS:
                ensure_windows_onenote_ready_for_tree_lookup(win)

            if IS_MACOS:
                ok, item_name = mac_center_selected_row(
                    win,
                    prefer_leftmost=True,
                    target_text=self.expected_text,
                )
                self.done.emit(bool(ok), item_name or self.expected_text or "")
                return

            tree = _find_tree_or_list(win)
            if not tree:
                self.done.emit(False, "")
                return
            self.cached_tree = tree

            expected_norm = _normalize_text(self.expected_text)
            if not expected_norm:
                selected_item = get_selected_tree_item_fast(tree)
                if selected_item is not None:
                    anchor_element, _anchor_source, placement = _resolve_alignment_target_for_selected_item(
                        selected_item,
                        tree,
                    )
                    _center_element_in_view(
                        selected_item,
                        tree,
                        anchor_element=anchor_element,
                        placement=placement,
                    )
                    try:
                        item_name = selected_item.window_text()
                    except Exception:
                        item_name = ""
                    self.done.emit(True, item_name)
                    return
                self.done.emit(False, "")
                return

            deadline = time.monotonic() + (
                0.28 if self.target_kind == "notebook" else 0.6
            )
            last_selected = None

            while not self.isInterruptionRequested() and time.monotonic() < deadline:
                selected_item = get_selected_tree_item_fast(tree)
                if selected_item is not None:
                    last_selected = selected_item
                    try:
                        selected_norm = _normalize_text(selected_item.window_text())
                    except Exception:
                        selected_norm = ""

                    if expected_norm and selected_norm == expected_norm:
                        anchor_element, _anchor_source, placement = _resolve_alignment_target_for_selected_item(
                            selected_item, tree
                        )
                        _center_element_in_view(
                            selected_item,
                            tree,
                            anchor_element=anchor_element,
                            placement=placement,
                        )
                        self.done.emit(True, selected_item.window_text())
                        return

                self.msleep(10)

            if self.isInterruptionRequested():
                return

            if last_selected is not None:
                anchor_element, _anchor_source, placement = _resolve_alignment_target_for_selected_item(
                    last_selected, tree
                )
                _center_element_in_view(
                    last_selected,
                    tree,
                    anchor_element=anchor_element,
                    placement=placement,
                )
                try:
                    last_text = last_selected.window_text()
                except Exception:
                    last_text = self.expected_text
                self.done.emit(True, last_text)
                return

            self.done.emit(False, "")
        except Exception as e:
            print(f"[WARN][CENTER][WORKER] {e}")
            self.done.emit(False, "")

_publish_context(globals())

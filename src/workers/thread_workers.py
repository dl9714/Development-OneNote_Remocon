# -*- coding: utf-8 -*-
"""
워커 스레드 모듈

백그라운드에서 실행되는 QThread 워커들을 정의합니다.
"""

from PyQt6.QtCore import QThread, pyqtSignal
from src.core.window_manager import WindowManager
from src.constants import ONENOTE_CLASS_NAME


class ReconnectWorker(QThread):
    """백그라운드에서 이전 OneNote 연결을 복구하는 워커"""

    finished = pyqtSignal(object)

    def __init__(self, settings_manager, parent=None):
        super().__init__(parent)
        self.settings_manager = settings_manager
        self.window_manager = WindowManager()

    def run(self):
        try:
            sig = self.settings_manager.get_connection_signature()
            if not sig:
                self.finished.emit({"ok": False, "status": "연결되지 않음"})
                return

            win = self.window_manager.find_window_by_signature(sig)
            if win and win.is_visible():
                window_title = win.window_text()
                # 새 시그니처 저장
                new_sig = self.window_manager.create_window_signature(win)
                self.settings_manager.set_connection_signature(new_sig)
                self.settings_manager.save()

                payload = {
                    "ok": True,
                    "status": f"(자동 재연결) '{window_title}'",
                    "sig": new_sig,
                }
            else:
                payload = {"ok": False, "status": "(재연결 실패) 이전 앱을 찾을 수 없습니다."}
        except Exception as e:
            payload = {"ok": False, "status": f"연결되지 않음 (오류: {e})"}

        self.finished.emit(payload)


class OneNoteWindowScanner(QThread):
    """OneNote 창 목록을 스캔하는 워커"""

    done = pyqtSignal(list)

    def __init__(self, my_pid: int, parent=None):
        super().__init__(parent)
        self.my_pid = my_pid
        self.window_manager = WindowManager()

    def run(self):
        results = []
        try:
            wins = self.window_manager.enumerate_windows(filter_title_substr=None)
            for w in wins:
                try:
                    if self.window_manager.is_onenote_window(w):
                        results.append(w)
                except Exception:
                    continue

            # OneNote 클래스 이름 우선 정렬
            results.sort(
                key=lambda r: (
                    r.get("class_name", "") != ONENOTE_CLASS_NAME,
                    r.get("title", ""),
                )
            )
        except Exception as e:
            print(f"[ERROR] OneNote 창 스캔 중 오류: {e}")
        finally:
            self.done.emit(results)


class WindowListWorker(QThread):
    """모든 윈도우 목록을 스캔하는 워커"""

    done = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.window_manager = WindowManager()

    def run(self):
        try:
            results = self.window_manager.enumerate_windows(filter_title_substr=None)
            self.done.emit(results)
        except Exception:
            self.done.emit([])

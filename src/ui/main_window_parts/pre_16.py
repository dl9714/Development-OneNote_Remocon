# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



class OpenNotebookRecordsWorker(QThread):
    done = pyqtSignal(dict)

    def __init__(self, sig: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.sig = copy.deepcopy(sig or {})

    def run(self):
        result = {
            "records": [],
            "error": "",
            "sig": copy.deepcopy(self.sig),
            "source": "MAC_SIDEBAR" if IS_MACOS else "COM",
        }
        try:
            ensure_pywinauto()
            if not _pwa_ready:
                result["error"] = "자동화 모듈이 로드되지 않았습니다."
                self.done.emit(result)
                return
            if IS_MACOS and self.sig:
                win = MacWindow(dict(self.sig))
                quick = mac_current_open_notebook_names_quick(
                    win,
                    ax_timeout_sec=1.2,
                    plist_timeout_sec=0.6,
                    sidebar_timeout_sec=14.0,
                    min_names_before_sidebar=48,
                )
                names = [
                    str(name or "").strip()
                    for name in (quick.get("names") or [])
                    if str(name or "").strip()
                ]
                debug = dict(quick.get("debug") or {})
                records = [
                    {
                        "id": "",
                        "name": name,
                        "path": name,
                        "url": "",
                        "last_accessed_at": 0,
                        "source": "MAC_QUICK_OPEN_NOTEBOOKS",
                    }
                    for name in names
                ]
                result["source"] = "MAC_QUICK_OPEN_NOTEBOOKS"
                if not records:
                    if not macos_accessibility_is_trusted():
                        result["error"] = (
                            "macOS 손쉬운 사용 권한이 없어 OneNote 화면 목록을 읽지 못했습니다. "
                            "개인정보 보호 및 보안 > 손쉬운 사용에서 OneNote_Remocon을 허용해야 합니다."
                        )
                    else:
                        result["error"] = (
                            "macOS 빠른 열린 전자필기장 조회가 빈 결과입니다 "
                            f"(title={debug.get('title_count')}, "
                            f"ax={debug.get('ax_count')}, "
                            f"plist={debug.get('plist_count')}, "
                            f"sidebar={debug.get('sidebar_count')})."
                        )
            else:
                records = _get_open_notebook_records_via_com(refresh=True)
            result["records"] = [dict(record) for record in records]
        except Exception as e:
            result["error"] = str(e)
        self.done.emit(result)


class CodexLocationLookupWorker(QThread):
    done = pyqtSignal(dict)

    def __init__(self, script: str, timeout: int = 60, parent=None):
        super().__init__(parent)
        self.script = script
        self.timeout = timeout

    def run(self):
        started_at = time.perf_counter()
        result = {"ok": False, "raw": "", "error": "", "elapsed_ms": 0}
        try:
            if IS_MACOS:
                wins = [
                    info
                    for info in enumerate_macos_windows(filter_title_substr=None)
                    if is_macos_onenote_window_info(info, os.getpid())
                ]
                wins.sort(key=lambda item: (not bool(item.get("frontmost")), item.get("title", "")))
                if not wins:
                    raise RuntimeError("열린 OneNote 창을 찾지 못했습니다.")
                win = MacWindow(dict(wins[0]))
                result["raw"] = macos_lookup_targets_json(win)
            else:
                result["raw"] = _run_powershell(self.script, timeout=self.timeout)
            result["ok"] = True
        except Exception as e:
            result["error"] = str(e)
        finally:
            result["elapsed_ms"] = int((time.perf_counter() - started_at) * 1000)
            self.done.emit(result)


class OtherWindowSelectionDialog(QDialog):
    def __init__(self, my_pid: int, parent=None):
        super().__init__(parent)
        self.my_pid = my_pid
        self.setWindowTitle("연결할 창을 더블클릭하세요.")
        self.setGeometry(400, 400, 500, 420)

        self.layout = QVBoxLayout(self)
        self.tip_label = QLabel("창 목록을 검색 중입니다...")
        self.layout.addWidget(self.tip_label)

        self.other_list_widget = QListWidget()
        self.layout.addWidget(self.other_list_widget)
        self.other_list_widget.hide()

        self.windows_info = []
        self.selected_info = None

        self.other_list_widget.itemDoubleClicked.connect(self.on_window_selected)

        self.worker = WindowListWorker()
        self.worker.done.connect(self._on_windows_list_ready)
        self.worker.start()

    def _on_windows_list_ready(self, results):
        self.tip_label.hide()
        if not results:
            self.tip_label.setText("표시할 창이 없습니다. 다시 시도해 주세요.")
            self.tip_label.show()
            return

        for r in results:
            pid = r.get("pid")
            if pid == self.my_pid:
                continue
            if not is_strict_onenote_window(r, self.my_pid):
                self.windows_info.append(r)

        self.windows_info.sort(key=lambda r: r.get("title", ""))

        if self.windows_info:
            items = [
                f'{r["title"]}  [{r["class_name"]}] (0x{r["handle"]:X})'
                for r in self.windows_info
            ]
            self.other_list_widget.addItems(items)
            self.other_list_widget.show()
        else:
            self.tip_label.setText("OneNote를 제외한 다른 창이 없습니다.")
            self.tip_label.show()

    def on_window_selected(self, item):
        row = self.other_list_widget.currentRow()
        if 0 <= row < len(self.windows_info):
            self.selected_info = self.windows_info[row]
            self.accept()

_publish_context(globals())

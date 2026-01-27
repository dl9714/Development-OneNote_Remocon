# -*- coding: utf-8 -*-
import sys
import json
import os
import time
import uuid
import ctypes
from ctypes import wintypes
from typing import Optional, List, Dict, Any

from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QDialog,
    QListWidget,
    QGroupBox,
    QTreeWidget,
    QTreeWidgetItem,
    QToolButton,
    QSplitter,
    QMenu,
    QInputDialog,
    QMessageBox,
    QAbstractItemView,
)
from PyQt6.QtCore import QThread, pyqtSignal, QTimer, Qt
from PyQt6.QtGui import QIcon, QAction

# ----------------- 0. 전역 상수 -----------------
PERSISTENCE_FILE = "onenote_connection.json"
FAVORITES_FILE = "favorites.json"

ONENOTE_CLASS_NAME = "ApplicationFrameWindow"  # UWP/Modern OneNote Class Name
SCROLL_STEP_SENSITIVITY = 40

# QTreeWidget 커스텀 데이터 롤
ROLE_TYPE = Qt.ItemDataRole.UserRole + 1  # 'group' | 'section'
ROLE_DATA = Qt.ItemDataRole.UserRole + 2  # dict payload

# ----------------- 0.1 pywinauto 지연 로딩 -----------------
Desktop = None
WindowNotFoundError = None
ElementNotFoundError = None
TimeoutError = None
UIAWrapper = None
UIAElementInfo = None
mouse = None
keyboard = None

_pwa_ready = False


def ensure_pywinauto():
    global _pwa_ready, Desktop, WindowNotFoundError, ElementNotFoundError, TimeoutError, UIAWrapper, UIAElementInfo, mouse, keyboard
    if _pwa_ready:
        return
    from pywinauto import Desktop as _Desktop, mouse as _mouse, keyboard as _keyboard
    from pywinauto.findwindows import (
        WindowNotFoundError as _WNF,
        ElementNotFoundError as _ENF,
    )
    from pywinauto.timings import TimeoutError as _TO
    from pywinauto.controls.uiawrapper import UIAWrapper as _UIAWrapper
    from pywinauto.uia_element_info import UIAElementInfo as _UIAElementInfo

    Desktop = _Desktop
    WindowNotFoundError = _WNF
    ElementNotFoundError = _ENF
    TimeoutError = _TO
    UIAWrapper = _UIAWrapper
    UIAElementInfo = _UIAElementInfo
    mouse = _mouse
    keyboard = _keyboard
    _pwa_ready = True


# ----------------- 0.2 Win32 빠른 창 열거 -----------------
_user32 = ctypes.windll.user32


def _win_get_window_text(hwnd):
    length = _user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(length + 1 if length > 0 else 1)
    _user32.GetWindowTextW(hwnd, buf, len(buf))
    return buf.value


def _win_get_class_name(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    _user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def enum_windows_fast(filter_title_substr=None):
    # filter_title_substr: None | str | Iterable[str]
    if isinstance(filter_title_substr, str):
        filters = [filter_title_substr.lower()]
    elif filter_title_substr:
        filters = [str(s).lower() for s in filter_title_substr]
    else:
        filters = None

    results = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def _enum_proc(hwnd, lparam):
        try:
            if not _user32.IsWindowVisible(hwnd):
                return True
            title = _win_get_window_text(hwnd)
            if not title:
                return True
            if filters and not any(f in title.lower() for f in filters):
                # 제목 필터가 있지만 일치하는 것이 없으면 건너뜀
                return True

            cls = _win_get_class_name(hwnd)
            pid = wintypes.DWORD()
            _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            results.append(
                {
                    "handle": int(hwnd),
                    "title": title,
                    "class_name": cls,
                    "pid": pid.value,
                }
            )
        except Exception:
            pass
        return True

    _user32.EnumWindows(_enum_proc, 0)
    return results


# ----------------- 1. 프로세스 실행 파일 경로 얻기 -----------------
def get_process_image_path(pid: int) -> Optional[str]:
    if not pid:
        return None
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    kernel32 = ctypes.windll.kernel32
    OpenProcess = kernel32.OpenProcess
    CloseHandle = kernel32.CloseHandle
    QueryFullProcessImageNameW = kernel32.QueryFullProcessImageNameW

    hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not hProcess:
        return None
    try:
        buf_len = wintypes.DWORD(1024)
        buf = ctypes.create_unicode_buffer(buf_len.value)
        if QueryFullProcessImageNameW(hProcess, 0, buf, ctypes.byref(buf_len)):
            return buf.value
        return None
    finally:
        CloseHandle(hProcess)


# ----------------- 1.1 엄격한 OneNote 창 검증 헬퍼 -----------------
def is_strict_onenote_window(w: Dict[str, Any], my_pid: int) -> bool:
    """주어진 창 정보가 실제로 OneNote 앱 창인지 엄격하게 확인합니다."""

    if w.get("pid") == my_pid:
        return False

    title_lower = w.get("title", "").lower()
    cls = w.get("class_name", "")
    pid = w.get("pid")

    # 1. Classic Desktop (OMain*)
    if "omain" in cls.lower():
        return True

    # 2. Modern App (ApplicationFrameWindow)
    if cls == ONENOTE_CLASS_NAME and (
        "onenote" in title_lower or "원노트" in title_lower
    ):
        # ApplicationFrameWindow는 다른 UWP 앱과 공유되지만,
        # 제목 필터와 함께 사용하면 높은 확률로 OneNote Modern 앱입니다.
        return True

    # 3. Fallback: 제목에 키워드가 있고, EXE가 OneNote인지 확인 (PID 검증)
    if "onenote" in title_lower or "원노트" in title_lower:
        # 이 단계는 조금 느릴 수 있지만 정확합니다.
        exe_path = get_process_image_path(pid)
        if exe_path:
            exe_name = os.path.basename(exe_path).lower()
            # OneNote 앱은 일반적으로 OneNote.exe 또는 OneNoteIm.exe를 사용
            if "onenote.exe" in exe_name or "onenoteim.exe" in exe_name:
                return True

    return False


# ----------------- 2. 디버깅용 UI 트리 출력 -----------------
def print_element_tree(element, depth=0):
    indent = "  " * depth
    try:
        title = element.window_text()
        control_type = element.element_info.control_type
        print(f'{indent}Title: "{title}", ControlType: "{control_type}"')
        for child in element.children():
            print_element_tree(child, depth + 1)
    except Exception as e:
        print(f"{indent} -> Error processing element: {e}")


# ----------------- 3-A. OneNote 창 목록 스캔 워커 (엄격한 스캔) -----------------
class OneNoteWindowScanner(QThread):
    done = pyqtSignal(list)

    def __init__(self, my_pid: int, parent=None):
        super().__init__(parent)
        self.my_pid = my_pid

    def run(self):
        results = []
        try:
            # 제목 필터 없이 모든 창을 빠르게 열거 (PID 검증을 위해)
            wins = enum_windows_fast(filter_title_substr=None)

            # 엄격한 필터링 적용
            for w in wins:
                try:
                    if is_strict_onenote_window(w, self.my_pid):
                        results.append(w)
                except Exception:
                    continue

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


# ----------------- 3-B. 전체 창 목록 스캔 워커(빠른 스캔) -----------------
class WindowListWorker(QThread):
    done = pyqtSignal(list)

    def run(self):
        try:
            results = enum_windows_fast(filter_title_substr=None)
            self.done.emit(results)
        except Exception:
            self.done.emit([])


# ----------------- 3-C. '기타 창' 선택 다이얼로그 -----------------
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

            # OneNote 앱이 아닌 창만 목록에 추가
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


# ----------------- 4. 짧은 폴링으로 Rect 안정화 대기 -----------------
def _wait_rect_settle(get_rect, timeout=0.3, interval=0.03):
    start = time.perf_counter()
    prev = get_rect()
    while time.perf_counter() - start < timeout:
        time.sleep(interval)
        cur = get_rect()
        if abs(cur.top - prev.top) < 2 and abs(cur.bottom - prev.bottom) < 2:
            break
        prev = cur


# ----------------- 5. 패턴 기반 수직 스크롤 시도 -----------------
def _scroll_vertical_via_pattern(
    container, direction: str, small=True, repeats=1
) -> bool:
    try:
        iface = getattr(container, "iface_scroll", None)
        if iface is None:
            return False

        from comtypes.gen.UIAutomationClient import (
            ScrollAmount_LargeIncrement,
            ScrollAmount_LargeDecrement,
            ScrollAmount_SmallIncrement,
            ScrollAmount_SmallDecrement,
            ScrollAmount_NoAmount,
        )

        v_inc = ScrollAmount_SmallIncrement if small else ScrollAmount_LargeIncrement
        v_dec = ScrollAmount_SmallDecrement if small else ScrollAmount_LargeDecrement
        v_amount = v_inc if direction == "down" else v_dec

        for _ in range(max(1, repeats)):
            iface.Scroll(ScrollAmount_NoAmount, v_amount)
        return True
    except Exception:
        return False


# ----------------- 6. 마우스 휠 기반 스크롤(폴백) -----------------
def _safe_wheel(scroll_container, steps: int):
    if steps == 0:
        return

    ensure_pywinauto()

    try:
        if hasattr(scroll_container, "wheel_scroll"):
            scroll_container.wheel_scroll(steps)
            return
    except Exception:
        pass

    try:
        if hasattr(scroll_container, "wheel_mouse_input"):
            scroll_container.wheel_mouse_input(wheel_dist=steps)
            return
    except Exception:
        pass

    try:
        rect = scroll_container.rectangle()
        center = rect.mid_point()
        try:
            mouse.scroll(coords=(center.x, center.y), wheel_dist=steps)
            return
        except Exception:
            pass
        try:
            mouse.wheel(coords=(center.x, center.y), wheel_dist=steps)
            return
        except Exception:
            pass
    except Exception:
        pass

    try:
        scroll_container.set_focus()
        if steps > 0:
            keyboard.send_keys("{UP %d}" % steps)
        else:
            keyboard.send_keys("{DOWN %d}" % abs(steps))
    except Exception:
        pass


# ----------------- 7. 선택 항목을 가장 빠르게 얻기 -----------------
def get_selected_tree_item_fast(tree_control):
    ensure_pywinauto()

    try:
        if hasattr(tree_control, "get_selection"):
            sel = tree_control.get_selection()
            if sel:
                return sel[0]
    except Exception:
        pass

    try:
        iface_sel = getattr(tree_control, "iface_selection", None)
        if iface_sel:
            arr = iface_sel.GetSelection()
            length = getattr(arr, "Length", 0)
            if length and length > 0:
                el = arr.GetElement(0)
                return UIAWrapper(UIAElementInfo(el))
    except Exception:
        pass

    try:
        for item in tree_control.children():
            try:
                if item.is_selected():
                    return item
            except Exception:
                pass
    except Exception:
        pass

    try:
        for item in tree_control.descendants(control_type="TreeItem"):
            try:
                if item.is_selected():
                    return item
            except Exception:
                pass
    except Exception:
        pass

    return None


# ----------------- 8. 페이지/섹션 컨테이너(Tree/List) 찾기 -----------------
def _find_tree_or_list(onenote_window):
    ensure_pywinauto()
    for ctype in ("Tree", "List"):
        try:
            return onenote_window.child_window(
                control_type=ctype, found_index=0
            ).wrapper_object()
        except Exception:
            continue
    return None


# ----------------- 8.1 지정 텍스트 섹션 찾기/선택 -----------------
def _normalize_text(s: Optional[str]) -> str:
    return " ".join(((s or "").strip().split())).lower()


def select_section_by_text(
    onenote_window, text: str, tree_control: Optional[object] = None
) -> bool:
    ensure_pywinauto()
    try:
        tree_control = tree_control or _find_tree_or_list(onenote_window)
        if not tree_control:
            return False

        target_norm = _normalize_text(text)

        # TreeItem 먼저, 안되면 ListItem 탐색
        def _scan(types: List[str]):
            for t in types:
                try:
                    for itm in tree_control.descendants(control_type=t):
                        try:
                            if _normalize_text(itm.window_text()) == target_norm:
                                try:
                                    itm.select()
                                except Exception:
                                    try:
                                        itm.click_input()
                                    except Exception:
                                        return False
                                return True
                        except Exception:
                            continue
                except Exception:
                    continue
            return False

        if _scan(["TreeItem"]) or _scan(["ListItem"]):
            _center_element_in_view(
                get_selected_tree_item_fast(tree_control), tree_control
            )
            return True
        return False
    except Exception:
        return False


# ----------------- 9. 요소를 중앙으로 위치시키는 함수(최적화) -----------------
def _center_element_in_view(element_to_center, scroll_container):
    ensure_pywinauto()
    try:
        try:
            element_to_center.iface_scroll_item.ScrollIntoView()
        except AttributeError:
            print(f"[WARN] ScrollItem 미지원: '{element_to_center.window_text()}'")
            return

        _wait_rect_settle(
            lambda: element_to_center.rectangle(), timeout=0.3, interval=0.03
        )

        rect_container = scroll_container.rectangle()
        rect_item = element_to_center.rectangle()
        item_center_y = (rect_item.top + rect_item.bottom) / 2
        container_center_y = (rect_container.top + rect_container.bottom) / 2
        offset = item_center_y - container_center_y
        print(
            f"[DEBUG] Item Center: {item_center_y}, Container Center: {container_center_y}, Offset: {offset}"
        )

        if abs(offset) <= 10:
            print("[INFO] 항목이 이미 중앙 근처입니다.")
            return

        def step_for(dy):
            return max(1, min(5, int(abs(dy) / 150)))

        for _ in range(3):
            if abs(offset) <= 10:
                break

            direction = "down" if offset > 0 else "up"
            repeats = step_for(offset)

            used_pattern = _scroll_vertical_via_pattern(
                scroll_container, direction=direction, small=True, repeats=repeats
            )
            if not used_pattern:
                wheel_steps = -repeats if offset > 0 else repeats
                _safe_wheel(scroll_container, wheel_steps)

            time.sleep(0.03)

            rect_container = scroll_container.rectangle()
            rect_item = element_to_center.rectangle()
            item_center_y = (rect_item.top + rect_item.bottom) / 2
            container_center_y = (rect_container.top + rect_container.bottom) / 2
            offset = item_center_y - container_center_y

        if abs(offset) > 10:
            print(f"[INFO] 중앙 근처까지 보정했지만 잔여 오프셋: {offset:.1f}")
        else:
            print("[INFO] 중앙 정렬 완료(보정).")
    except Exception as e:
        print(f"[WARN] 중앙 정렬 중 오류: {e}")


# ----------------- 10. 선택된 항목을 중앙으로 스크롤 -----------------
def scroll_selected_item_to_center(
    onenote_window, tree_control: Optional[object] = None
):
    ensure_pywinauto()
    try:
        tree_control = tree_control or _find_tree_or_list(onenote_window)
        if not tree_control:
            print("[ERROR] 페이지/섹션 목록(Tree/List)을 찾지 못했습니다.")
            return False, None

        selected_item = get_selected_tree_item_fast(tree_control)
        if not selected_item:
            print("[ERROR] 현재 선택된 항목을 찾을 수 없습니다.")
            return False, None

        item_name = selected_item.window_text()
        print(f"[INFO] 선택 항목 '{item_name}' 중앙 정렬 시작")
        _center_element_in_view(selected_item, tree_control)
        print(f"[INFO] 중앙 정렬 완료: '{item_name}'")
        return True, item_name
    except (ElementNotFoundError, TimeoutError):
        print(f"[ERROR] 목록 컨트롤을 찾지 못했거나 시간 초과되었습니다.")
        return False, None
    except Exception as e:
        print(f"[ERROR] 선택 항목 정렬 중 오류: {e}")
        return False, None


# ----------------- 11. 연결 시그니처 저장/스코어 기반 재획득 -----------------
def build_window_signature(win) -> dict:
    try:
        pid = win.process_id()
    except Exception:
        pid = None
    exe_path = get_process_image_path(pid) if pid else None
    exe_name = os.path.basename(exe_path).lower() if exe_path else None
    try:
        handle = win.handle
    except Exception:
        handle = None
    try:
        title = win.window_text()
    except Exception:
        title = None
    try:
        cls_name = win.class_name()
    except Exception:
        cls_name = None

    return {
        "handle": handle,
        "pid": pid,
        "class_name": cls_name,
        "title": title,
        "exe_path": exe_path,
        "exe_name": exe_name,
    }


def save_connection_info(window_element):
    try:
        info = build_window_signature(window_element)
        with open(PERSISTENCE_FILE, "w", encoding="utf-8") as f:
            json.dump(info, f, ensure_ascii=False)
    except Exception as e:
        print(f"[ERROR] 연결 정보 저장 실패: {e}")


def _score_candidate(win, sig) -> int:
    try:
        title = (win.window_text() or "").lower()
        cls = win.class_name() or ""
        pid = win.process_id()
        exe_path = get_process_image_path(pid) or ""
        exe_name = os.path.basename(exe_path).lower() if exe_path else ""

        score = 0
        if sig.get("handle") and getattr(win, "handle", None) == sig["handle"]:
            score += 100
        if sig.get("exe_name") and exe_name == sig["exe_name"]:
            score += 50
        if "onenote.exe" in exe_name:
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
        if cls == ONENOTE_CLASS_NAME:
            score += 5
        return score
    except Exception:
        return -1


def _score_candidate_dict(c, sig) -> int:
    try:
        title = (c.get("title") or "").lower()
        cls = c.get("class_name") or ""
        pid = c.get("pid")
        exe_path = get_process_image_path(pid) or ""
        exe_name = os.path.basename(exe_path).lower() if exe_path else ""

        score = 0
        if sig.get("handle") and c.get("handle") == sig["handle"]:
            score += 100
        if sig.get("exe_name") and exe_name == sig["exe_name"]:
            score += 50
        if "onenote.exe" in exe_name:
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
        if cls == ONENOTE_CLASS_NAME:
            score += 5
        return score
    except Exception:
        return -1


def reacquire_window_by_signature(sig) -> Optional[object]:
    ensure_pywinauto()

    # 1) 핸들 직행
    h = sig.get("handle")
    if h:
        try:
            w = Desktop(backend="uia").window(handle=h)
            if w.is_visible():
                return w
        except Exception:
            pass

    # 2) 빠른 열거로 후보 점수화
    candidates = enum_windows_fast(filter_title_substr=None)
    best, best_score = None, -1
    for c in candidates:
        s = _score_candidate_dict(c, sig)
        if s > best_score:
            best, best_score = c, s

    if best and best_score >= 30:
        try:
            w = Desktop(backend="uia").window(handle=best["handle"])
            if w.is_visible():
                return w
        except Exception:
            return None
    return None


# ----------------- 12. 저장된 정보로 재연결 -----------------
def load_connection_info_and_reconnect():
    ensure_pywinauto()
    if not os.path.exists(PERSISTENCE_FILE):
        return None, "연결되지 않음"
    try:
        with open(PERSISTENCE_FILE, "r", encoding="utf-8") as f:
            sig = json.load(f)

        win = reacquire_window_by_signature(sig)
        if win and win.is_visible():
            window_title = win.window_text()
            print(f"[INFO] 이전 OneNote 앱에 자동 재연결 성공: '{window_title}'")
            try:
                save_connection_info(win)
            except Exception:
                pass
            return win, f"(자동 재연결) '{window_title}'"

        print("[INFO] 이전 OneNote 앱을 찾지 못했습니다.")
        return None, "(재연결 실패) 이전 앱을 찾을 수 없습니다."
    except Exception as e:
        print(f"[ERROR] 재연결 중 오류: {e}")
        return None, "연결되지 않음"


# ----------------- 13. 백그라운드 자동 재연결 워커 -----------------
class ReconnectWorker(QThread):
    finished = pyqtSignal(object)

    def run(self):
        try:
            ensure_pywinauto()
            win, status = load_connection_info_and_reconnect()
            if win:
                try:
                    sig = build_window_signature(win)
                    save_connection_info(win)
                    payload = {"ok": True, "status": status, "sig": sig}
                except Exception:
                    payload = {"ok": False, "status": status}
            else:
                payload = {"ok": False, "status": status}
        except Exception as e:
            payload = {"ok": False, "status": f"연결되지 않음 (오류: {e})"}
        self.finished.emit(payload)


# ----------------- 14-A. 즐겨찾기 트리 위젯 -----------------
class FavoritesTree(QTreeWidget):
    structureChanged = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderHidden(True)
        self.setColumnCount(1)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setIndentation(16)
        self.setAnimated(True)
        self.setExpandsOnDoubleClick(True)

    def dropEvent(self, event):
        super().dropEvent(event)
        self.structureChanged.emit()


# ----------------- 14. PyQt GUI -----------------
class OneNoteScrollRemoconApp(QWidget):
    def __init__(self):
        super().__init__()
        self.onenote_window = None
        self.tree_control = None
        self._reconnect_worker = None
        self._scanner_worker = None
        self.onenote_windows_info: List[Dict] = []
        self.my_pid = os.getpid()

        # 즐겨찾기 관련
        self.fav_tree: FavoritesTree = None
        self._auto_center_after_activate = (
            True  # 섹션 활성화 시 OneNote면 선택+중앙정렬
        )

        self.init_ui("준비됨 (자동 재연결 중...)")
        self._load_favorites()
        QTimer.singleShot(0, self.refresh_onenote_list)  # 시작 즉시 가벼운 스캔 시작
        QTimer.singleShot(0, self._start_auto_reconnect)  # 병렬로 자동 재연결

    def init_ui(self, initial_status):
        self.setWindowTitle("OneNote 전자필기장 스크롤 리모컨")
        self.setGeometry(200, 180, 960, 540)

        # --- 스타일시트 정의 ---
        COLOR_BACKGROUND = "#2E2E2E"
        COLOR_PRIMARY_TEXT = "#E0E0E0"
        COLOR_SECONDARY_TEXT = "#B0B0B0"
        COLOR_GROUPBOX_BG = "#3C3C3C"
        COLOR_ACCENT = "#A6D854"
        COLOR_ACCENT_HOVER = "#B8E966"
        COLOR_ACCENT_PRESSED = "#95C743"
        COLOR_SECONDARY_BUTTON = "#555555"
        COLOR_SECONDARY_BUTTON_HOVER = "#666666"
        COLOR_SECONDARY_BUTTON_PRESSED = "#444444"
        COLOR_LIST_BG = "#252525"
        COLOR_LIST_SELECTED = "#0078D7"

        self.setStyleSheet(
            f"""
            QWidget {{
                background-color: {COLOR_BACKGROUND};
                color: {COLOR_PRIMARY_TEXT};
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }}
            QGroupBox {{
                background-color: {COLOR_GROUPBOX_BG};
                border: 1px solid {COLOR_BACKGROUND};
                border-radius: 8px;
                margin-top: 10px;
                font-weight: bold;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 10px;
                left: 10px;
            }}
            QLabel {{
                color: {COLOR_SECONDARY_TEXT};
                font-weight: normal;
            }}
            QListWidget {{
                background-color: {COLOR_LIST_BG};
                border: 1px solid {COLOR_GROUPBOX_BG};
                border-radius: 4px;
            }}
            QListWidget::item:selected {{
                background-color: {COLOR_LIST_SELECTED};
                color: white;
            }}
            QTreeWidget {{
                background-color: {COLOR_LIST_BG};
                border: 1px solid {COLOR_GROUPBOX_BG};
                border-radius: 6px;
            }}
            QToolButton {{
                background-color: {COLOR_SECONDARY_BUTTON};
                color: {COLOR_PRIMARY_TEXT};
                border: none;
                border-radius: 4px;
                padding: 6px 8px;
            }}
            QToolButton:hover {{
                background-color: {COLOR_SECONDARY_BUTTON_HOVER};
            }}
            QToolButton:pressed {{
                background-color: {COLOR_SECONDARY_BUTTON_PRESSED};
            }}
            QPushButton {{
                background-color: {COLOR_SECONDARY_BUTTON};
                color: {COLOR_PRIMARY_TEXT};
                border: none;
                border-radius: 4px;
                padding: 6px 12px;
            }}
            QPushButton:hover {{
                background-color: {COLOR_SECONDARY_BUTTON_HOVER};
            }}
            QPushButton:pressed {{
                background-color: {COLOR_SECONDARY_BUTTON_PRESSED};
            }}
            QPushButton:disabled {{
                background-color: #404040;
                color: #808080;
            }}
        """
        )

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(10)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        main_layout.addWidget(splitter, stretch=1)

        # --- 왼쪽: 즐겨찾기 패널 ---
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        fav_group = QGroupBox("즐겨찾기")
        fav_layout = QVBoxLayout(fav_group)

        # 툴바
        tb_layout = QHBoxLayout()
        self.btn_add_group = QToolButton()
        self.btn_add_group.setText("그룹 추가")
        self.btn_add_group.setToolTip("그룹(폴더) 추가")
        self.btn_add_group.clicked.connect(self._add_group)

        self.btn_add_section_current = QToolButton()
        self.btn_add_section_current.setText("현재 전자필기장 추가")
        self.btn_add_section_current.setToolTip(
            "현재 연결된 창(및 선택된 OneNote 섹션)을 즐겨찾기에 추가"
        )
        self.btn_add_section_current.clicked.connect(self._add_section_from_current)

        self.btn_add_section_other = QToolButton()
        self.btn_add_section_other.setText("다른 창 추가")
        self.btn_add_section_other.setToolTip(
            "다른 실행 중인 창을 즐겨찾기 섹션으로 추가"
        )
        self.btn_add_section_other.clicked.connect(self._add_section_from_other_window)

        self.btn_rename = QToolButton()
        self.btn_rename.setText("이름 바꾸기")
        self.btn_rename.clicked.connect(self._rename_favorite_item)

        self.btn_delete = QToolButton()
        self.btn_delete.setText("삭제")
        self.btn_delete.clicked.connect(self._delete_favorite_item)

        tb_layout.addWidget(self.btn_add_group)
        tb_layout.addWidget(self.btn_add_section_current)
        tb_layout.addWidget(self.btn_add_section_other)
        tb_layout.addStretch(1)
        tb_layout.addWidget(self.btn_rename)
        tb_layout.addWidget(self.btn_delete)

        fav_layout.addLayout(tb_layout)

        self.fav_tree = FavoritesTree()
        self.fav_tree.itemDoubleClicked.connect(self._on_fav_item_double_clicked)
        self.fav_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.fav_tree.customContextMenuRequested.connect(self._on_fav_context_menu)
        self.fav_tree.structureChanged.connect(self._save_favorites)
        self.fav_tree.itemChanged.connect(lambda *_: self._save_favorites())

        fav_layout.addWidget(self.fav_tree)
        left_layout.addWidget(fav_group, stretch=1)

        splitter.addWidget(left_panel)
        splitter.setStretchFactor(0, 0)

        # --- 오른쪽: 기존 UI들 묶음 ---
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        # --- OneNote 창 목록 그룹 ---
        connection_group = QGroupBox("OneNote 창 목록")
        connection_layout = QVBoxLayout(connection_group)

        list_header_layout = QHBoxLayout()
        list_header_layout.addWidget(
            QLabel("더블클릭하여 연결 및 중앙 정렬"),
            alignment=Qt.AlignmentFlag.AlignLeft,
        )
        list_header_layout.addStretch()

        self.refresh_button = QPushButton(" 새로고침")
        refresh_icon = self.style().standardIcon(
            QApplication.style().StandardPixmap.SP_BrowserReload
        )
        self.refresh_button.setIcon(QIcon(refresh_icon))
        self.refresh_button.setToolTip("목록 새로고침")
        self.refresh_button.clicked.connect(self.refresh_onenote_list)
        list_header_layout.addWidget(self.refresh_button)

        connection_layout.addLayout(list_header_layout)

        self.onenote_list_widget = QListWidget()
        self.onenote_list_widget.addItem("자동 재연결 시도 중...")
        self.onenote_list_widget.itemDoubleClicked.connect(
            self.connect_and_center_from_list_item
        )
        connection_layout.addWidget(self.onenote_list_widget)

        right_layout.addWidget(connection_group)

        # --- 자동화 기능 그룹 ---
        actions_group = QGroupBox("자동화 기능")
        actions_layout = QVBoxLayout(actions_group)

        self.center_button = QPushButton("현재 선택된 전자필기장 중앙으로 정렬")
        center_icon = self.style().standardIcon(
            QApplication.style().StandardPixmap.SP_ArrowRight
        )
        self.center_button.setIcon(QIcon(center_icon))
        self.center_button.setStyleSheet(
            f"""
            QPushButton {{
                background-color: {COLOR_ACCENT};
                color: #111;
                font-weight: bold;
                padding: 8px 16px;
            }}
            QPushButton:hover {{ background-color: {COLOR_ACCENT_HOVER}; }}
            QPushButton:pressed {{ background-color: {COLOR_ACCENT_PRESSED}; }}
            QPushButton:disabled {{
                background-color: #555555;
                color: #999999;
                border: 1px solid #444444;
            }}
        """
        )
        self.center_button.clicked.connect(self.center_selected_item_action)
        self.center_button.setEnabled(False)
        actions_layout.addWidget(self.center_button)

        other_buttons_layout = QHBoxLayout()
        connect_other_button = QPushButton("다른 앱에 연결...")
        connect_other_button.clicked.connect(self.select_other_window)
        other_buttons_layout.addWidget(connect_other_button)

        disconnect_button = QPushButton("연결 해제")
        disconnect_button.clicked.connect(self.disconnect_and_clear_info)
        other_buttons_layout.addWidget(disconnect_button)
        actions_layout.addLayout(other_buttons_layout)

        right_layout.addWidget(actions_group)

        right_layout.addStretch(1)

        # --- 상태 표시줄 ---
        self.connection_status_label = QLabel(initial_status)
        self.connection_status_label.setStyleSheet("color: #B0B0B0; padding: 5px;")
        right_layout.addWidget(
            self.connection_status_label, alignment=Qt.AlignmentFlag.AlignRight
        )

        splitter.addWidget(right_panel)
        splitter.setStretchFactor(1, 1)

        self.setLayout(main_layout)

        # 초기 splitter 사이즈
        splitter.setSizes([280, 680])

    # <<< 상태/버튼 업데이트 통합
    def update_status_and_ui(self, status_text: str, is_connected: bool):
        self.connection_status_label.setText(status_text)
        self.center_button.setEnabled(is_connected)

    def _start_auto_reconnect(self):
        self.refresh_button.setEnabled(False)
        self._reconnect_worker = ReconnectWorker()
        self._reconnect_worker.finished.connect(self._on_reconnect_done)
        self._reconnect_worker.start()

    def _on_reconnect_done(self, payload):
        self._reconnect_worker = None
        status = payload.get("status", "연결되지 않음")
        if payload.get("ok"):
            ensure_pywinauto()
            sig = payload.get("sig", {})
            target = None
            try:
                h = sig.get("handle")
                if h:
                    target = Desktop(backend="uia").window(handle=h)
                if not target or not target.is_visible():
                    target = reacquire_window_by_signature(sig)
            except Exception:
                target = None

            if target:
                self.onenote_window = target
                try:
                    save_connection_info(self.onenote_window)
                except Exception:
                    pass
                self.update_status_and_ui(f"연결됨: {status}", True)
                QTimer.singleShot(0, self._cache_tree_control)
                self.refresh_onenote_list()
                return

        self.onenote_window = None
        self.tree_control = None
        self.update_status_and_ui(f"상태: {status}", False)
        self.refresh_onenote_list()

    def refresh_onenote_list(self):
        if self._scanner_worker and self._scanner_worker.isRunning():
            return

        self.onenote_list_widget.clear()
        self.onenote_list_widget.addItem("OneNote 창을 검색 중입니다...")
        self.onenote_list_widget.setEnabled(False)
        self.refresh_button.setEnabled(False)

        self._scanner_worker = OneNoteWindowScanner(self.my_pid)
        self._scanner_worker.done.connect(self._on_onenote_list_ready)
        self._scanner_worker.start()

    def _on_onenote_list_ready(self, results: List[Dict]):
        self.onenote_windows_info = results
        self.onenote_list_widget.clear()

        if not results:
            self.onenote_list_widget.addItem("실행 중인 OneNote 창을 찾지 못했습니다.")
        else:
            items = [f'{r["title"]}  [{r["class_name"]}]' for r in results]
            self.onenote_list_widget.addItems(items)

        self.onenote_list_widget.setEnabled(True)
        self.refresh_button.setEnabled(True)

    def _cache_tree_control(self):
        self.tree_control = _find_tree_or_list(self.onenote_window)
        if self.tree_control:
            try:
                _ = self.tree_control.children()
            except Exception:
                pass

    # 연결 성공 여부 반환
    def _perform_connection(self, info: Dict) -> bool:
        try:
            ensure_pywinauto()
            self.onenote_window = Desktop(backend="uia").window(handle=info["handle"])
            if not self.onenote_window.is_visible():
                raise ElementNotFoundError

            window_title = self.onenote_window.window_text()
            save_connection_info(self.onenote_window)

            status_text = f"연결됨: '{window_title}'"
            self.update_status_and_ui(status_text, True)
            QTimer.singleShot(0, self._cache_tree_control)
            return True

        except ElementNotFoundError:
            self.update_status_and_ui("연결 실패: 선택한 창이 보이지 않습니다.", False)
            self.refresh_onenote_list()
            return False
        except Exception as e:
            self.update_status_and_ui(f"연결 실패: {e}", False)
            return False

    def connect_and_center_from_list_item(self, item):
        row = self.onenote_list_widget.currentRow()
        if 0 <= row < len(self.onenote_windows_info):
            info = self.onenote_windows_info[row]
            if self._perform_connection(info):
                QTimer.singleShot(50, self.center_selected_item_action)

    def select_other_window(self):
        dialog = OtherWindowSelectionDialog(self.my_pid, self)
        if dialog.exec():
            info = dialog.selected_info
            if info:
                self._perform_connection(info)

    def disconnect_and_clear_info(self):
        self.onenote_window = None
        self.tree_control = None

        status_msg = "연결 해제됨."
        if os.path.exists(PERSISTENCE_FILE):
            try:
                os.remove(PERSISTENCE_FILE)
                status_msg = "연결 해제 및 저장된 정보 삭제 완료."
            except OSError as e:
                status_msg = f"연결 해제. 정보 파일 삭제 실패: {e}"

        self.update_status_and_ui(status_msg, False)

    def _pre_action_check(self) -> bool:
        ensure_pywinauto()
        if not self.onenote_window:
            self.update_status_and_ui("오류: 앱에 연결되어 있지 않습니다.", False)
            return False
        try:
            if not self.onenote_window.is_visible():
                raise ElementNotFoundError
        except ElementNotFoundError:
            self.update_status_and_ui(
                "오류: 연결된 창을 찾을 수 없습니다. 연결을 해제합니다.", False
            )
            self.disconnect_and_clear_info()
            return False
        return True

    def center_selected_item_action(self):
        if not self._pre_action_check():
            return

        if not self.tree_control:
            self.tree_control = _find_tree_or_list(self.onenote_window)

        success, item_name = scroll_selected_item_to_center(
            self.onenote_window, self.tree_control
        )
        if not success:
            self.tree_control = _find_tree_or_list(self.onenote_window)
            success, item_name = scroll_selected_item_to_center(
                self.onenote_window, self.tree_control
            )

        if success:
            self.update_status_and_ui(f"성공: '{item_name}' 중앙 정렬 완료.", True)
        else:
            self.update_status_and_ui(
                "실패: 선택 항목을 찾거나 정렬하지 못했습니다.", True
            )

    # ----------------- 15. 즐겨찾기 로드/세이브 -----------------
    def _load_favorites(self):
        self.fav_tree.clear()
        if not os.path.exists(FAVORITES_FILE):
            return
        try:
            with open(FAVORITES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                for node in data:
                    self._append_fav_node(self.fav_tree.invisibleRootItem(), node)
            self.fav_tree.expandAll()
        except Exception as e:
            print(f"[ERROR] 즐겨찾기 로드 실패: {e}")

    def _save_favorites(self):
        try:
            data = []
            root = self.fav_tree.invisibleRootItem()
            for i in range(root.childCount()):
                data.append(self._serialize_fav_item(root.child(i)))
            with open(FAVORITES_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[ERROR] 즐겨찾기 저장 실패: {e}")

    def _serialize_fav_item(self, item: QTreeWidgetItem) -> Dict[str, Any]:
        node_type = item.data(0, ROLE_TYPE)
        payload = item.data(0, ROLE_DATA) or {}
        node = {
            "type": node_type,
            "id": payload.get("id") or str(uuid.uuid4()),
            "name": item.text(0),
        }
        if node_type == "section":
            node["target"] = payload.get("target", {})
        # children
        children = []
        for i in range(item.childCount()):
            children.append(self._serialize_fav_item(item.child(i)))
        if children:
            node["children"] = children
        return node

    def _append_fav_node(
        self, parent: QTreeWidgetItem, node: Dict[str, Any]
    ) -> QTreeWidgetItem:
        item = QTreeWidgetItem(parent)
        node_type = node.get("type", "group")
        name = node.get("name", "이름 없음")
        item.setText(0, name)
        item.setData(0, ROLE_TYPE, node_type)
        payload = {"id": node.get("id", str(uuid.uuid4()))}
        if node_type == "section":
            payload["target"] = node.get(
                "target", {}
            )  # {'sig':..., 'section_text':...}
            item.setIcon(
                0,
                self.style().standardIcon(
                    QApplication.style().StandardPixmap.SP_FileIcon
                ),
            )
        else:
            item.setIcon(
                0,
                self.style().standardIcon(
                    QApplication.style().StandardPixmap.SP_DirIcon
                ),
            )
        item.setData(0, ROLE_DATA, payload)
        item.setFlags(
            item.flags()
            | Qt.ItemFlag.ItemIsEditable
            | Qt.ItemFlag.ItemIsDragEnabled
            | Qt.ItemFlag.ItemIsDropEnabled
            | Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
        )
        for ch in node.get("children", []):
            self._append_fav_node(item, ch)
        return item

    # ----------------- 16. 즐겨찾기 조작 -----------------
    def _current_fav_item(self) -> Optional[QTreeWidgetItem]:
        items = self.fav_tree.selectedItems()
        return items[0] if items else None

    def _add_group(self):
        parent = self._current_fav_item()
        if parent and parent.data(0, ROLE_TYPE) == "section":
            parent = parent.parent()  # 섹션 아래에 그룹은 넣지 않음
        parent = parent or self.fav_tree.invisibleRootItem()
        node = {"type": "group", "name": "새 그룹", "children": []}
        item = self._append_fav_node(parent, node)
        self.fav_tree.editItem(item, 0)
        self._save_favorites()

    def _add_section_from_current(self):
        ensure_pywinauto()
        if not self.onenote_window:
            QMessageBox.information(self, "안내", "먼저 연결된 창이 있어야 합니다.")
            return

        title = ""
        try:
            title = self.onenote_window.window_text()
        except Exception:
            pass

        section_text = None
        try:
            tc = self.tree_control or _find_tree_or_list(self.onenote_window)
            if tc:
                sel = get_selected_tree_item_fast(tc)
                if sel:
                    section_text = sel.window_text()
        except Exception:
            pass

        default_name = section_text or title or "새 섹션"
        name, ok = QInputDialog.getText(
            self, "섹션 즐겨찾기 추가", "표시 이름:", text=default_name
        )
        if not ok or not name.strip():
            return

        try:
            sig = build_window_signature(self.onenote_window)
        except Exception:
            sig = {}

        target = {"sig": sig, "section_text": section_text}
        node = {"type": "section", "name": name.strip(), "target": target}

        parent = self._current_fav_item()
        if parent and parent.data(0, ROLE_TYPE) == "section":
            parent = parent.parent()
        parent = parent or self.fav_tree.invisibleRootItem()
        self._append_fav_node(parent, node)
        self._save_favorites()

    def _add_section_from_other_window(self):
        dialog = OtherWindowSelectionDialog(self.my_pid, self)
        if not dialog.exec():
            return
        info = dialog.selected_info
        if not info:
            return
        # 이름 입력
        default_name = (info.get("title") or "새 섹션").strip() or "새 섹션"
        name, ok = QInputDialog.getText(
            self, "섹션 즐겨찾기 추가", "표시 이름:", text=default_name
        )
        if not ok or not name.strip():
            return
        # sig 구성
        try:
            ensure_pywinauto()
            win = Desktop(backend="uia").window(handle=info["handle"])
            sig = build_window_signature(win)
        except Exception:
            sig = {
                "handle": info.get("handle"),
                "pid": info.get("pid"),
                "class_name": info.get("class_name"),
                "title": info.get("title"),
            }
        target = {"sig": sig, "section_text": None}
        node = {"type": "section", "name": name.strip(), "target": target}

        parent = self._current_fav_item()
        if parent and parent.data(0, ROLE_TYPE) == "section":
            parent = parent.parent()
        parent = parent or self.fav_tree.invisibleRootItem()
        self._append_fav_node(parent, node)
        self._save_favorites()

    def _rename_favorite_item(self):
        item = self._current_fav_item()
        if not item:
            return
        self.fav_tree.editItem(item, 0)
        # 저장은 itemChanged 시그널에서 처리

    def _delete_favorite_item(self):
        item = self._current_fav_item()
        if not item:
            return
        node_type = item.data(0, ROLE_TYPE)
        name = item.text(0)
        if node_type == "group" and item.childCount() > 0:
            ret = QMessageBox.question(
                self, "삭제 확인", f"그룹 '{name}'과(와) 모든 하위 항목을 삭제할까요?"
            )
            if ret != QMessageBox.StandardButton.Yes:
                return
        else:
            ret = QMessageBox.question(self, "삭제 확인", f"'{name}'을(를) 삭제할까요?")
            if ret != QMessageBox.StandardButton.Yes:
                return
        parent = item.parent() or self.fav_tree.invisibleRootItem()
        parent.removeChild(item)
        self._save_favorites()

    def _on_fav_context_menu(self, pos):
        item = self._current_fav_item()
        menu = QMenu(self)
        act_add_group = QAction("그룹 추가", self)
        act_add_group.triggered.connect(self._add_group)
        menu.addAction(act_add_group)

        act_add_curr = QAction("현재 전자필기장 추가", self)
        act_add_curr.triggered.connect(self._add_section_from_current)
        menu.addAction(act_add_curr)

        act_add_other = QAction("다른 창 추가", self)
        act_add_other.triggered.connect(self._add_section_from_other_window)
        menu.addAction(act_add_other)

        if item:
            menu.addSeparator()
            act_rename = QAction("이름 바꾸기", self)
            act_rename.triggered.connect(self._rename_favorite_item)
            menu.addAction(act_rename)

            act_delete = QAction("삭제", self)
            act_delete.triggered.connect(self._delete_favorite_item)
            menu.addAction(act_delete)

        menu.exec(self.fav_tree.viewport().mapToGlobal(pos))

    def _on_fav_item_double_clicked(self, item: QTreeWidgetItem, col: int):
        if not item:
            return
        node_type = item.data(0, ROLE_TYPE)
        if node_type != "section":
            # 그룹은 열고 닫기만
            item.setExpanded(
                item.isExpanded()
            )  # will fix below (Python doesn't support !)
            return
        payload = item.data(0, ROLE_DATA) or {}
        target = payload.get("target") or {}
        self._activate_favorite_section(target, display_name=item.text(0))

    def _activate_favorite_section(
        self, target: Dict[str, Any], display_name: str = ""
    ):
        sig = target.get("sig") or {}
        if not sig:
            QMessageBox.warning(
                self, "오류", "이 즐겨찾기에는 대상 창 정보가 없습니다."
            )
            return

        ensure_pywinauto()
        win = reacquire_window_by_signature(sig)
        if not win:
            QMessageBox.information(
                self, "실패", f"대상 창을 찾을 수 없습니다.\n'{display_name}'"
            )
            return

        try:
            win.set_focus()
        except Exception:
            pass

        # 앱 연결 갱신
        try:
            info = {
                "handle": win.handle,
                "title": win.window_text(),
                "class_name": win.class_name(),
                "pid": win.process_id(),
            }
            connected = self._perform_connection(info)
        except Exception:
            connected = False

        if connected and self._auto_center_after_activate:
            exe_name = (sig.get("exe_name") or "").lower()
            if "onenote" in exe_name or "onenote" in (sig.get("title") or "").lower():
                section_text = target.get("section_text")
                if section_text:
                    ok = select_section_by_text(
                        self.onenote_window, section_text, self.tree_control
                    )
                    if ok:
                        scroll_selected_item_to_center(
                            self.onenote_window, self.tree_control
                        )

        self.update_status_and_ui(f"활성화: '{display_name}'", True)


# ----------------- 17. 엔트리 포인트 -----------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = OneNoteScrollRemoconApp()
    ex.show()

    # fix: Python doesn't support '!' operator; patch group toggle here
    def _toggle_group(item, col):
        node_type = item.data(0, ROLE_TYPE)
        if node_type != "section":
            item.setExpanded(not item.isExpanded())

    ex.fav_tree.itemDoubleClicked.disconnect()
    ex.fav_tree.itemDoubleClicked.connect(_toggle_group)
    ex.fav_tree.itemDoubleClicked.connect(ex._on_fav_item_double_clicked)

    sys.exit(app.exec())

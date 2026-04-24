# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

# -*- coding: utf-8 -*-
import sys
import os
import time
from types import SimpleNamespace
from src.lazy_import import LazyAttr, LazyModule, lazy_class

if sys.platform.startswith("win"):
    try:
        import winreg
    except ImportError:  # pragma: no cover - Windows 전용
        winreg = None
else:
    winreg = None


json = LazyModule("json")
base64 = LazyModule("base64")
hashlib = LazyModule("hashlib")
copy = LazyModule("copy")
ctypes = LazyModule("ctypes")
unicodedata = LazyModule("unicodedata")
re = LazyModule("re")
uuid = LazyModule("uuid")
traceback = LazyModule("traceback")
wintypes = LazyModule("ctypes.wintypes")
subprocess = LazyModule("subprocess")
difflib = LazyModule("difflib")
html = LazyModule("html")
threading = LazyModule("threading")
_urllib_parse = LazyModule("urllib.parse")


def urlparse(*args, **kwargs): return _urllib_parse.urlparse(*args, **kwargs)
def parse_qs(*args, **kwargs): return _urllib_parse.parse_qs(*args, **kwargs)

from PyQt6.QtWidgets import (
    QApplication,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QPushButton,
    QLabel,
    QDialog,
    QListWidget,
    QGroupBox,
    QTreeWidget,
    QTreeWidgetItem,
    QTreeWidgetItemIterator,
    QToolButton,
    QSplitter,
    QMenu,
    QMenuBar,
    QInputDialog,
    QMessageBox,
    QAbstractItemView,
    QMainWindow,
    QFileDialog,
    QWidget,
    QLineEdit,
    QListWidgetItem,
    QTabWidget,
    QTextEdit,
    QComboBox,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
)
from PyQt6.QtCore import (
    QThread,
    pyqtSignal,
    QTimer,
    Qt,
    QSettings,
    QEvent,
    QRect,
    QByteArray,
)
from PyQt6.QtGui import QIcon, QAction, QBrush, QColor, QPixmap, QPainter, QPen
from src.app_version import APP_BUILD_VERSION, APP_VERSION
_MACOS_UI_MODULE = "src.macos_ui"
MacAutomationError = lazy_class(_MACOS_UI_MODULE, "MacAutomationError", Exception)
MacDesktop = lazy_class(_MACOS_UI_MODULE, "MacDesktop")
MacWindow = lazy_class(_MACOS_UI_MODULE, "MacWindow")
mac_recent_notebook_records_from_cache = LazyAttr(
    _MACOS_UI_MODULE, "_recent_notebook_records_from_cache"
)
mac_current_notebook_name = LazyAttr(_MACOS_UI_MODULE, "current_notebook_name")
mac_current_open_notebook_names = LazyAttr(
    _MACOS_UI_MODULE, "current_open_notebook_names"
)
mac_current_open_notebook_names_quick = LazyAttr(
    _MACOS_UI_MODULE, "current_open_notebook_names_quick"
)
macos_accessibility_is_trusted = LazyAttr(
    _MACOS_UI_MODULE, "macos_accessibility_is_trusted"
)
macos_last_ax_notebook_debug = LazyAttr(
    _MACOS_UI_MODULE, "macos_last_ax_notebook_debug"
)
mac_open_tab_notebook_records = LazyAttr(
    _MACOS_UI_MODULE, "open_tab_notebook_records"
)
mac_open_recent_notebook_record = LazyAttr(
    _MACOS_UI_MODULE, "open_recent_notebook_record"
)
mac_current_outline_context = LazyAttr(_MACOS_UI_MODULE, "current_outline_context")
enumerate_macos_windows = LazyAttr(_MACOS_UI_MODULE, "enumerate_macos_windows")
enumerate_macos_windows_quick = LazyAttr(
    _MACOS_UI_MODULE, "enumerate_macos_windows_quick"
)
is_macos_onenote_window_info = LazyAttr(_MACOS_UI_MODULE, "is_onenote_window_info")
macos_lookup_targets_json = LazyAttr(_MACOS_UI_MODULE, "macos_lookup_targets_json")
mac_pick_selected_row = LazyAttr(_MACOS_UI_MODULE, "pick_selected_row")
mac_recent_notebook_records = LazyAttr(_MACOS_UI_MODULE, "recent_notebook_records")
mac_select_page_row_by_text = LazyAttr(_MACOS_UI_MODULE, "select_page_row_by_text")
mac_select_open_notebook_by_name = LazyAttr(
    _MACOS_UI_MODULE, "select_open_notebook_by_name"
)
mac_select_row_by_text = LazyAttr(_MACOS_UI_MODULE, "select_row_by_text")
mac_center_selected_row = LazyAttr(_MACOS_UI_MODULE, "center_selected_row")
from src.platform_support import (
    IS_MACOS,
    IS_WINDOWS,
    ONENOTE_MAC_BUNDLE_ID,
    default_icon_path,
    open_path_in_system,
    open_url_in_system,
)


class WheelSafeComboBox(QComboBox):
    """Ignore wheel changes while collapsed so parent scroll areas keep scrolling."""

    def showPopup(self):
        try:
            width = max(self.width(), self.view().sizeHintForColumn(0) + 36)
            self.view().setMinimumWidth(width)
        except Exception:
            pass
        super().showPopup()

    def wheelEvent(self, event):
        try:
            popup_open = self.view().isVisible()
        except Exception:
            popup_open = False
        if popup_open:
            super().wheelEvent(event)
            return
        event.ignore()


FavoritesTree = lazy_class("src.ui.widgets", "FavoritesTree")
BufferTree = lazy_class("src.ui.widgets", "BufferTree")
TreeNameEditDelegate = lazy_class("src.ui.widgets", "TreeNameEditDelegate")

# ----------------- 0. 전역 상수 -----------------
SETTINGS_FILE = "OneNote_Remocon_Setting.json"
SETTINGS_PATH_POINTER_FILE = "OneNote_Remocon_Setting.path"
SETTINGS_PATH_ENV = "ONENOTE_REMOCON_SETTINGS_PATH"
APP_ICON_PATH = default_icon_path()

ONENOTE_CLASS_NAME = "ApplicationFrameWindow"
SCROLL_STEP_SENSITIVITY = 40
MACOS_GENERIC_ONENOTE_TITLES = {
    "microsoft onenote",
    "onenote",
    ONENOTE_MAC_BUNDLE_ID.casefold(),
}
MACOS_RECENT_NOTEBOOK_DIALOG_TITLE_TOKENS = (
    "최근 전자 필기장",
    "새 전자 필기장",
    "recent notebook",
    "new notebook",
)
CODEX_PLATFORM_WINDOWS = "windows"
CODEX_PLATFORM_MACOS = "macos"

ROLE_TYPE = Qt.ItemDataRole.UserRole + 1
ROLE_DATA = Qt.ItemDataRole.UserRole + 2
ROLE_OPEN_NOTEBOOK = Qt.ItemDataRole.UserRole + 3

# ----------------- 0.1 버퍼 트리 고정/가상 노드 -----------------
DEFAULT_GROUP_ID = "group-default-fixed"
DEFAULT_GROUP_NAME = "Default"
AGG_BUFFER_ID = "buffer-aggregate-all-sections"
AGG_BUFFER_NAME = "종합"
AGG_UNCLASSIFIED_GROUP_ID = "group-aggregate-uncategorized"
AGG_UNCLASSIFIED_GROUP_NAME = "분류 안 된 전자필기장"
AGG_CLASSIFIED_GROUP_ID = "group-aggregate-categorized"
AGG_CLASSIFIED_GROUP_NAME = "분류된 전자필기장"

# OneNote: 전체 전자필기장 자동등록 그룹
AUTO_ONENOTE_GROUP_ID = "group-onenote-auto"
AUTO_ONENOTE_GROUP_NAME = "OneNote(자동등록)"

# OneNote: 백스테이지(열기 화면) 자동화 텍스트/필터
ONENOTE_OPEN_NOTEBOOK_VIEW_TEXTS = (
    "전자 필기장 열기",
    "전자필기장 열기",
    "open notebook",
)
ONENOTE_FILE_MENU_TEXTS = ("파일", "file")
ONENOTE_OPEN_MENU_TEXTS = ("열기", "open")
ONENOTE_RECENT_MENU_TEXTS = ("최근 항목", "recent")
ONENOTE_SEARCH_TEXTS = ("검색", "search")
ONENOTE_NOTEBOOK_ITEM_CONTROL_TYPES = (
    "ListItem",
    "TreeItem",
    "DataItem",
    "Button",
    "Hyperlink",
    "Custom",
)
ONENOTE_NOTEBOOK_SKIP_EXACT_TEXTS = {
    "정보",
    "새로 만들기",
    "열기",
    "복사본 저장",
    "인쇄",
    "공유",
    "내보내기",
    "보내기",
    "최근 항목",
    "검색",
    "개인",
    "폴더",
    "info",
    "new",
    "open",
    "save a copy",
    "print",
    "share",
    "export",
    "send",
    "recent",
    "search",
    "personal",
    "folder",
}
ONENOTE_NOTEBOOK_SKIP_CONTAINS = (
    "shared with",
    "just me",
    "전자 필기장 관리",
    "manage notebooks",
    "onedrive(",
    "onedrive (",
    "switch account",
    "계정 전환",
)

# ----------------- 0.05 정렬 헬퍼 -----------------
def _name_sort_key(text: Any) -> str:
    "이름 기준 정렬용 키(유니코드 정규화 + casefold)."
    try:
        s = text if isinstance(text, str) else ("" if text is None else str(text))
        return unicodedata.normalize("NFKD", s).casefold()
    except Exception:
        try:
            return str(text)
        except Exception:
            return ""


def _platform_ui_font_stack(include_generic: bool = False) -> str:
    if IS_MACOS:
        fonts = [
            "'Apple SD Gothic Neo'",
            "'AppleGothic'",
            "'SF Pro Text'",
            "'Helvetica Neue'",
        ]
    elif IS_WINDOWS:
        fonts = ["'Malgun Gothic'", "'Segoe UI'"]
    else:
        fonts = ["'Noto Sans CJK KR'", "'Noto Sans'", "'DejaVu Sans'"]
    if include_generic:
        fonts.append("sans-serif")
    return ", ".join(fonts)


def _center_target_ui_name() -> str:
    return "섹션" if IS_MACOS else "전자필기장"


def _main_window_title() -> str:
    return "OneNote 빠른 이동" if IS_MACOS else "OneNote 전자필기장 위치정렬"


def _remocon_workspace_tab_title() -> str:
    return "빠른 이동" if IS_MACOS else "위치정렬"


def _current_add_button_label() -> str:
    return "현재 전자필기장 보기 추가" if IS_MACOS else "현재 전자필기장 추가"


def _favorite_activate_button_label() -> str:
    return "선택 전자필기장 보기" if IS_MACOS else "선택 항목 실행"


def _connection_group_title() -> str:
    return "OneNote 연결 대상" if IS_MACOS else "OneNote 창 목록"


def _current_actions_group_title() -> str:
    return "현재 전자필기장 보기 제어" if IS_MACOS else "현재 열린 항목 제어"


def _buffer_group_title() -> str:
    return "전자필기장 분류" if IS_MACOS else "프로젝트/등록 영역"


def _buffer_group_add_label() -> str:
    return "묶음" if IS_MACOS else "그룹"


def _buffer_item_add_label() -> str:
    return "분류함" if IS_MACOS else "버퍼"


def _rename_button_label() -> str:
    return "이름 변경" if IS_MACOS else "이름변경"


def _favorites_group_title() -> str:
    return "전자필기장 보기" if IS_MACOS else "모듈영역"


def _register_all_notebooks_button_label() -> str:
    return "열린 전자필기장 새로고침" if IS_MACOS else "종합 새로고침"


def _open_unchecked_notebooks_button_label(candidate_count: Optional[int] = None) -> str:
    base = "체크 없는 전자필기장 열기"
    if candidate_count is None:
        return base
    try:
        count = max(0, int(candidate_count))
    except Exception:
        return base
    return f"{base} ({count}개)"


def _open_unchecked_notebooks_tip() -> str:
    return "종합 목록에서 연두색 체크가 없는 전자필기장만 현재 열린 목록과 대조해서 엽니다."


_MAC_UI_OPEN_SOURCE_HINTS = {
    "MAC_RECENT",
    "MAC_RECENT_CACHE",
    "MAC_RECENT_DIALOG",
    "MAC_SHORTCUT",
    "MAC_OPEN_TAB",
    "RECENT",
    "RECENT_CACHE",
    "RECENT_DIALOG",
}
_APP_ONLY_NOTEBOOK_SOURCES = {
    "AGG_UNCHECKED",
    "AGG_UNCLASSIFIED",
    "SETTINGS_BUFFER",
}


def _notebook_record_source_hints(record: Dict[str, Any]) -> Set[str]:
    hints: Set[str] = set()
    raw_hints = (record or {}).get("_candidate_sources")
    if isinstance(raw_hints, (list, tuple, set)):
        for value in raw_hints:
            token = str(value or "").strip()
            if token:
                hints.add(token)
    raw_source = str((record or {}).get("source") or "").strip()
    if raw_source:
        for token in re.split(r"[+,/|]", raw_source):
            token = token.strip()
            if token:
                hints.add(token)
    return hints


def _mac_record_has_ui_open_hint(record: Dict[str, Any]) -> bool:
    return bool(_notebook_record_source_hints(record) & _MAC_UI_OPEN_SOURCE_HINTS)

_publish_context(globals())

# -*- coding: utf-8 -*-
import sys
import json
import os
import time
import uuid
import traceback
import ctypes
from ctypes import wintypes
from contextlib import contextmanager
from typing import Optional, List, Dict, Any, Set, Callable, Tuple
import base64
import hashlib
import copy
import unicodedata
import subprocess
import re
import difflib
import html
from types import SimpleNamespace
from urllib.parse import urlparse, parse_qs

try:
    import winreg
except ImportError:  # pragma: no cover - Windows 전용
    winreg = None

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
from PyQt6.QtGui import QIcon, QAction, QBrush, QColor


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


# widgets 모듈에서 커스텀 트리 위젯 임포트
from src.ui.widgets import FavoritesTree, BufferTree, TreeNameEditDelegate

# ----------------- 0. 전역 상수 -----------------
SETTINGS_FILE = "OneNote_Remocon_Setting.json"
SETTINGS_PATH_POINTER_FILE = "OneNote_Remocon_Setting.path"
SETTINGS_PATH_ENV = "ONENOTE_REMOCON_SETTINGS_PATH"
APP_ICON_PATH = "assets/app_icon.ico"

ONENOTE_CLASS_NAME = "ApplicationFrameWindow"
SCROLL_STEP_SENSITIVITY = 40

ROLE_TYPE = Qt.ItemDataRole.UserRole + 1
ROLE_DATA = Qt.ItemDataRole.UserRole + 2

# ----------------- 0.1 버퍼 트리 고정/가상 노드 -----------------
DEFAULT_GROUP_ID = "group-default-fixed"
DEFAULT_GROUP_NAME = "Default"
AGG_BUFFER_ID = "buffer-aggregate-all-sections"
AGG_BUFFER_NAME = "종합"
AGG_UNCLASSIFIED_GROUP_ID = "group-aggregate-uncategorized"
AGG_UNCLASSIFIED_GROUP_NAME = "미분류 카테고리"
AGG_CLASSIFIED_GROUP_ID = "group-aggregate-categorized"
AGG_CLASSIFIED_GROUP_NAME = "카테고리 분류된 전자필기장"

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

# ----------------- 0.0 설정 파일 경로 헬퍼 -----------------
def _get_app_base_path() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


def _get_default_settings_file_path() -> str:
    return os.path.join(_get_app_base_path(), SETTINGS_FILE)


def _settings_path_config_dir() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "OneNote_Remocon")


def _settings_path_config_file() -> str:
    return os.path.join(_settings_path_config_dir(), SETTINGS_PATH_POINTER_FILE)


def _expand_external_settings_path(raw: str, base_dir: Optional[str] = None) -> str:
    value = (raw or "").strip().strip('"').strip("'")
    if not value:
        return ""
    value = os.path.expandvars(os.path.expanduser(value))
    if not os.path.isabs(value) and base_dir:
        value = os.path.join(base_dir, value)
    return os.path.abspath(value)


def _read_settings_path_pointer(path: str) -> str:
    try:
        if not os.path.exists(path):
            return ""
        with open(path, "r", encoding="utf-8") as f:
            raw = f.read().strip()
        if not raw:
            return ""
        try:
            data = json.loads(raw)
            if isinstance(data, dict):
                raw = data.get("settings_path") or data.get("path") or ""
        except Exception:
            pass
        return _expand_external_settings_path(raw, os.path.dirname(path))
    except Exception:
        return ""


def _get_external_settings_file_path() -> str:
    env_path = _expand_external_settings_path(os.environ.get(SETTINGS_PATH_ENV, ""))
    if env_path:
        return env_path

    user_pointer = _read_settings_path_pointer(_settings_path_config_file())
    if user_pointer:
        return user_pointer

    local_pointer = _read_settings_path_pointer(
        os.path.join(_get_app_base_path(), SETTINGS_PATH_POINTER_FILE)
    )
    if local_pointer:
        return local_pointer

    return ""


def _set_external_settings_file_path(path: str) -> None:
    resolved = _expand_external_settings_path(path)
    if not resolved:
        raise ValueError("설정 JSON 경로가 비어 있습니다.")
    os.makedirs(_settings_path_config_dir(), exist_ok=True)
    with open(_settings_path_config_file(), "w", encoding="utf-8") as f:
        f.write(resolved)


def _clear_external_settings_file_path() -> None:
    try:
        os.remove(_settings_path_config_file())
    except FileNotFoundError:
        pass


def _settings_path_mode_label() -> str:
    external = _get_external_settings_file_path()
    if external:
        return f"공용 JSON: {external}"
    return f"기본 JSON: {_get_default_settings_file_path()}"


def _get_settings_file_path() -> str:
    """
    설정 파일(쓰기 가능)의 경로를 반환합니다.
    우선순위:
    1. ONENOTE_REMOCON_SETTINGS_PATH 환경변수
    2. 사용자 공용 포인터(%APPDATA%/OneNote_Remocon/OneNote_Remocon_Setting.path)
    3. 실행 위치의 로컬 포인터(OneNote_Remocon_Setting.path)
    4. 기본값: EXE 위치 또는 프로젝트 루트의 OneNote_Remocon_Setting.json
    """
    external = _get_external_settings_file_path()
    if external:
        return external

    return _get_default_settings_file_path()


def _find_settings_seed_file(primary_path: str) -> Optional[str]:
    """EXE 첫 실행 시 복사해 쓸 초기 설정 파일 후보를 찾습니다."""
    candidates = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(os.path.join(meipass, SETTINGS_FILE))
    candidates.append(_get_default_settings_file_path())
    candidates.append(os.path.join(os.path.abspath("."), SETTINGS_FILE))
    candidates.append(
        os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", SETTINGS_FILE))
    )

    primary_abs = os.path.abspath(primary_path)
    seen = set()
    for path in candidates:
        try:
            path_abs = os.path.abspath(path)
        except Exception:
            continue
        if path_abs in seen or path_abs == primary_abs:
            continue
        seen.add(path_abs)
        if os.path.exists(path_abs):
            return path_abs
    return None


def _settings_has_user_buffers(settings: Dict[str, Any]) -> bool:
    """Default/종합만 있는 빈 설정인지, 실제 프로젝트 데이터가 있는지 구분합니다."""
    nodes = settings.get("favorites_buffers")
    if not isinstance(nodes, list):
        return False

    stack = list(nodes)
    while stack:
        node = stack.pop()
        if not isinstance(node, dict):
            continue
        node_type = node.get("type")
        if node_type == "group":
            if node.get("id") != DEFAULT_GROUP_ID:
                return True
            children = node.get("children")
            if isinstance(children, list):
                stack.extend(children)
        elif node_type == "buffer":
            data = node.get("data")
            if node.get("id") != AGG_BUFFER_ID:
                return True
            if isinstance(data, list) and data:
                return True
        elif node_type in ("notebook", "section"):
            return True
    return False


# ----------------- 0.0 설정 파일 로드/저장 유틸리티 (즐겨찾기 버퍼 구조 추가) -----------------
DEFAULT_SETTINGS = {
    "window_geometry": {"x": 200, "y": 180, "width": 960, "height": 540},
    "splitter_states": None,  # 새 설정 항목 추가
    "connection_signature": None,
    "favorites_buffers": [],  # List 형태로 변경됨
    "active_buffer_id": None, # ID 기반으로 변경
    "debug_hotpaths": False,
    "debug_perf_logs": False,
}

_JSON_TEXT_CACHE: Dict[str, Dict[str, Any]] = {}
_SETTINGS_OBJECT_CACHE: Dict[str, Dict[str, Any]] = {}
_PROCESS_IMAGE_PATH_CACHE: Dict[int, Dict[str, Any]] = {}
_PROCESS_IMAGE_PATH_CACHE_TTL_SEC = 5.0
_OPEN_NOTEBOOK_RECORDS_CACHE: Dict[str, Any] = {
    "expires_at": 0.0,
    "records": [],
}
_OPEN_NOTEBOOK_RECORDS_CACHE_TTL_SEC = 10.0


def _get_file_signature(path: str) -> Optional[tuple]:
    try:
        st = os.stat(path)
        return (st.st_mtime_ns, st.st_size)
    except OSError:
        return None


def _update_json_text_cache(
    path: str,
    text: str,
    *,
    file_sig: Optional[tuple] = None,
) -> None:
    _JSON_TEXT_CACHE[path] = {
        "text": text,
        "sig": file_sig if file_sig is not None else _get_file_signature(path),
    }


def _update_settings_object_cache(path: str, data: Dict[str, Any]) -> None:
    _SETTINGS_OBJECT_CACHE[path] = {
        "sig": _get_file_signature(path),
        "data": copy.deepcopy(data),
    }


def _clear_open_notebook_records_cache() -> None:
    _OPEN_NOTEBOOK_RECORDS_CACHE["expires_at"] = 0.0
    _OPEN_NOTEBOOK_RECORDS_CACHE["records"] = []


def _dump_json_text(obj: Dict[str, Any]) -> str:
    return json.dumps(obj, ensure_ascii=False, indent=2)

def _write_json_text(path: str, text: str) -> bool:
    """내용이 바뀐 경우에만 .bak 백업 후 원자적으로 저장합니다."""
    file_sig = _get_file_signature(path)
    cache_entry = _JSON_TEXT_CACHE.get(path) or {}
    cached_text = cache_entry.get("text")
    cached_sig = cache_entry.get("sig")
    if cached_text == text and cached_sig == file_sig:
        return False

    old_text = cached_text if cached_sig == file_sig else None
    if old_text is None:
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    old_text = f.read()
                _update_json_text_cache(path, old_text, file_sig=file_sig)
        except Exception:
            old_text = None
    if old_text == text:
        return False
    # 기존 파일은 .bak으로 백업 (마이그레이션 실패/되돌리기 대비)
    try:
        if os.path.exists(path):
            import shutil
            shutil.copy2(path, path + ".bak")
    except Exception:
        pass
    tmp_path = f"{path}.tmp.{os.getpid()}"
    with open(tmp_path, "w", encoding="utf-8") as f:
        f.write(text)
    os.replace(tmp_path, path)
    _update_json_text_cache(path, text)
    return True

def _write_json(path: str, obj: Dict[str, Any]) -> bool:
    """UTF-8(한글 유지)로 설정 파일을 저장합니다."""
    return _write_json_text(path, _dump_json_text(obj))


def _migrate_favorites_buffers_inplace(data: Dict[str, Any]) -> bool:
    """
    1패널(버퍼 트리) 도입 이후에도 예전 설정/즐겨찾기 JSON을 그대로 인식하도록 마이그레이션합니다.

    지원하는 레거시 형태:
      - favorites: [...]                     (아주 구버전)
      - favorites_buffers: {name: [...]}     (버퍼=이름 딕셔너리)
      - favorites_buffers: [...]             (버퍼가 없고, 그룹/섹션 목록만 있는 리스트)
    """
    migrated = False

    # (A) favorites -> favorites_buffers(dict) (구버전)
    if "favorites" in data and "favorites_buffers" not in data:
        data["favorites_buffers"] = {"기본 즐겨찾기 버퍼": data.get("favorites") or []}
        data["active_buffer"] = "기본 즐겨찾기 버퍼"
        data.pop("favorites", None)
        migrated = True

    raw = data.get("favorites_buffers")

    # (B) dict -> list[buffer] (이전 버전: {name: [data...]})
    if isinstance(raw, dict):
        new_list = []
        name_to_id = {}
        for name, fav_data in raw.items():
            buf = {
                "type": "buffer",
                "id": str(uuid.uuid4()),
                "name": name,
                "data": fav_data if isinstance(fav_data, list) else [],
            }
            new_list.append(buf)
            name_to_id[name] = buf["id"]
        data["favorites_buffers"] = new_list

        # active_buffer(name) -> active_buffer_id
        legacy_name = data.get("active_buffer")
        if legacy_name and legacy_name in name_to_id:
            data["active_buffer_id"] = name_to_id[legacy_name]
        migrated = True

    # (C) list인데 buffer 노드가 하나도 없으면: "즐겨찾기 트리 데이터"만 있던 구버전으로 간주
    # ⚠️ 주의: 신규 구조에서는 최상단이 group이고 buffer는 children에 들어갈 수 있다.
    # 따라서 has_buffer 검사는 반드시 재귀적으로 수행해야 한다.
    def _has_any_buffer_node(nodes: Any) -> bool:
        if not isinstance(nodes, list):
            return False
        for n in nodes:
            if not isinstance(n, dict):
                continue
            if n.get("type") == "buffer":
                return True
            if n.get("type") == "group":
                if _has_any_buffer_node(n.get("children") or []):
                    return True
        return False

    def _collect_buffer_ids(nodes: Any, out: list) -> None:
        if not isinstance(nodes, list):
            return
        for n in nodes:
            if not isinstance(n, dict):
                continue
            if n.get("type") == "buffer" and n.get("id"):
                out.append(n.get("id"))
            elif n.get("type") == "group":
                _collect_buffer_ids(n.get("children") or [], out)

    raw2 = data.get("favorites_buffers")
    if isinstance(raw2, list):
        has_buffer = _has_any_buffer_node(raw2)
        if (not has_buffer) and raw2:
            data["favorites_buffers"] = [{
                "type": "buffer",
                "id": str(uuid.uuid4()),
                "name": "기본 즐겨찾기 버퍼",
                "data": raw2,
            }]
            migrated = True

    # (D) active_buffer_id 유효성 검사 (버퍼가 group 아래에 있을 수 있으므로 재귀 수집)
    buf_ids: list = []
    _collect_buffer_ids(data.get("favorites_buffers", []), buf_ids)
    if buf_ids:
        if data.get("active_buffer_id") not in buf_ids:
            data["active_buffer_id"] = buf_ids[0]
            migrated = True
    else:
        # 버퍼가 하나도 없으면 active_id도 None
        if data.get("active_buffer_id") is not None:
            data["active_buffer_id"] = None
            migrated = True

    # 더 이상 쓰지 않는 레거시 키 정리
    if "active_buffer" in data:
        data.pop("active_buffer", None)
        migrated = True

    return migrated


def load_settings() -> Dict[str, Any]:
    # 설정 파일 경로를 실행 파일 위치 기준으로 가져옴
    settings_path = _get_settings_file_path()

    file_sig = _get_file_signature(settings_path)
    cache_entry = _SETTINGS_OBJECT_CACHE.get(settings_path)
    if file_sig is not None and cache_entry and cache_entry.get("sig") == file_sig:
        cached = copy.deepcopy(cache_entry.get("data") or DEFAULT_SETTINGS)
        _ensure_default_and_aggregate_inplace(cached)
        return cached

    if not os.path.exists(settings_path):
        seed_path = _find_settings_seed_file(settings_path)
        if seed_path:
            try:
                with open(seed_path, "r", encoding="utf-8") as f:
                    raw_text = f.read()
                data = json.loads(raw_text)
                _migrate_favorites_buffers_inplace(data)
                settings = DEFAULT_SETTINGS.copy()
                settings.update(data)
                _ensure_default_and_aggregate_inplace(settings)
                try:
                    _write_json(settings_path, settings)
                except Exception as e:
                    print(f"[WARN] 초기 설정 복사 실패(메모리 로드는 계속): {e}")
                _update_settings_object_cache(settings_path, settings)
                return settings
            except Exception as e:
                print(f"[WARN] 초기 설정 파일 로드 실패({seed_path}): {e}")

        settings = DEFAULT_SETTINGS.copy()
        _ensure_default_and_aggregate_inplace(settings)
        try:
            _write_json(settings_path, settings)
        except Exception as e:
            print(f"[WARN] 기본 설정 파일 생성 실패(메모리 로드는 계속): {e}")
        _update_settings_object_cache(settings_path, settings)
        return settings
    try:
        with open(settings_path, "r", encoding="utf-8") as f:
            raw_text = f.read()
        _update_json_text_cache(settings_path, raw_text, file_sig=file_sig)
        data = json.loads(raw_text)

        # 하위 호환성을 위한 마이그레이션 로직
        migrated = _migrate_favorites_buffers_inplace(data)

        settings = DEFAULT_SETTINGS.copy()
        settings.update(data)
        # ✅ 로드 직후에도 Default/종합 구조 강제
        _ensure_default_and_aggregate_inplace(settings)
        if not _settings_has_user_buffers(settings):
            seed_path = _find_settings_seed_file(settings_path)
            if seed_path:
                try:
                    with open(seed_path, "r", encoding="utf-8") as f:
                        seed_data = json.loads(f.read())
                    _migrate_favorites_buffers_inplace(seed_data)
                    seed_settings = DEFAULT_SETTINGS.copy()
                    seed_settings.update(seed_data)
                    _ensure_default_and_aggregate_inplace(seed_settings)
                    if _settings_has_user_buffers(seed_settings):
                        settings = seed_settings
                        migrated = True
                        print(f"[INFO] 빈 EXE 설정 대신 초기 설정 사용: {seed_path}")
                except Exception as e:
                    print(f"[WARN] 초기 설정 재로드 실패({seed_path}): {e}")
        if migrated:
            try:
                _write_json(settings_path, settings)
                print(f"[INFO] 설정 마이그레이션 완료: {settings_path}")
            except Exception as e:
                print(f"[WARN] 마이그레이션 저장 실패(무시): {e}")

        _update_settings_object_cache(settings_path, settings)
        return settings
    except Exception as e:
        print(f"[ERROR] 설정 파일 로드 실패: {e}")
        settings = DEFAULT_SETTINGS.copy()
        _ensure_default_and_aggregate_inplace(settings)
        return settings


def save_settings(data: Dict[str, Any]) -> bool:
    # 설정 파일 경로를 실행 파일 위치 기준으로 가져옴
    settings_path = _get_settings_file_path()
    try:
        payload = dict(data)
        payload.pop("favorites", None)
        # ✅ 저장 직전에 항상 Default/종합 구조 강제 보정
        _ensure_default_and_aggregate_inplace(payload)
        changed = _write_json(settings_path, payload)
        _update_settings_object_cache(settings_path, payload)
        return changed
    except Exception as e:
        print(f"[ERROR] 설정 파일 저장 실패: {e}")
        return False


def _ensure_default_and_aggregate_inplace(settings: Dict[str, Any]) -> None:
    """
    1패널(버퍼 트리)에
      - Default 그룹(최상단 고정)
      - 종합(가상 버퍼)
    를 항상 보장합니다.
    """
    bufs = settings.get("favorites_buffers")
    if not isinstance(bufs, list):
        bufs = []
        settings["favorites_buffers"] = bufs

    # Default 그룹 찾기/생성
    default_idx = None
    default_node = None
    for i, n in enumerate(bufs):
        if isinstance(n, dict) and n.get("type") == "group" and n.get("id") == DEFAULT_GROUP_ID:
            default_idx = i
            default_node = n
            break
    if default_node is None:
        default_node = {
            "type": "group",
            "id": DEFAULT_GROUP_ID,
            "name": DEFAULT_GROUP_NAME,
            "locked": True,
            "children": []
        }
        bufs.insert(0, default_node)
    else:
        # 항상 0번으로 이동 + 이름 강제
        default_node["name"] = DEFAULT_GROUP_NAME
        default_node["locked"] = True
        if default_idx != 0:
            bufs.pop(default_idx)
            bufs.insert(0, default_node)

    # Default 그룹 children 보정
    children = default_node.get("children")
    if not isinstance(children, list):
        children = []
        default_node["children"] = children

    # 종합(가상 버퍼) 찾기/생성 (Default 그룹의 첫 번째로 고정)
    agg_idx = None
    agg_node = None
    for i, c in enumerate(children):
        if isinstance(c, dict) and c.get("type") == "buffer" and c.get("id") == AGG_BUFFER_ID:
            agg_idx = i
            agg_node = c
            break
    if agg_node is None:
        agg_node = {
            "type": "buffer",
            "id": AGG_BUFFER_ID,
            "name": AGG_BUFFER_NAME,
            "virtual": "aggregate",
            "locked": True,
            "data": []
        }
        children.insert(0, agg_node)
    else:
        # 항상 children[0]으로 이동 + 이름/속성 강제
        agg_node["name"] = AGG_BUFFER_NAME
        agg_node["locked"] = True
        agg_node["virtual"] = "aggregate"
        if agg_idx != 0:
            children.pop(agg_idx)
            children.insert(0, agg_node)

    # active_buffer_id가 종합/기타 버퍼에 대해 유효하지 않으면 첫 "일반 buffer"로 보정
    # (종합으로 시작하고 싶으면 이 부분을 제거하면 됨)
    if settings.get("active_buffer_id") is None:
        # 기본은 종합이 아니라, 첫 일반 buffer를 찾는다
        first_normal = _find_first_normal_buffer_id(bufs)
        if first_normal:
            settings["active_buffer_id"] = first_normal


def _find_first_normal_buffer_id(nodes: List[Dict[str, Any]]) -> Optional[str]:
    def _walk(lst):
        for n in lst:
            if not isinstance(n, dict):
                continue
            if n.get("type") == "buffer":
                if n.get("id") != AGG_BUFFER_ID:
                    return n.get("id")
            if n.get("type") == "group":
                cid = _walk(n.get("children") or [])
                if cid:
                    return cid
        return None
    return _walk(nodes)


def _find_buffer_node_by_id(nodes: Any, buffer_id: Optional[str]) -> Optional[Dict[str, Any]]:
    if not buffer_id or not isinstance(nodes, list):
        return None

    stack = list(reversed(nodes))
    while stack:
        node = stack.pop()
        if not isinstance(node, dict):
            continue
        node_type = node.get("type")
        if node_type == "buffer" and node.get("id") == buffer_id:
            return node
        if node_type == "group":
            children = node.get("children")
            if isinstance(children, list) and children:
                stack.extend(reversed(children))
    return None


def _collect_all_sections_dedup(settings: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    모든 버퍼의 data에서 section/notebook들을 수집해서 (sig + text) 기준 중복 제거 후 반환.
    - 반환 형태: 중앙트리(_load_favorites_into_center_tree)에 바로 넣을 수 있는 node 리스트
    """
    bufs = settings.get("favorites_buffers", [])
    out: List[Dict[str, Any]] = []
    seen = set()

    def _freeze_key(value: Any):
        if isinstance(value, dict):
            try:
                items = sorted(value.items(), key=lambda pair: str(pair[0]))
            except Exception:
                items = value.items()
            return tuple((str(k), _freeze_key(v)) for k, v in items)
        if isinstance(value, list):
            return tuple(_freeze_key(v) for v in value)
        if isinstance(value, (str, int, float, bool, type(None))):
            return value
        return str(value)

    def _section_key(node: Dict[str, Any]):
        t = node.get("target") or {}
        sig = t.get("sig") or {}
        ty = node.get("type", "section")
        # 섹션명 또는 노트북명 수집
        text = t.get("section_text") or t.get("notebook_text") or ""
        return (ty, _freeze_key(sig), text)

    def _walk_fav_nodes(nodes: Any):
        if not isinstance(nodes, list):
            return
        for n in nodes:
            if not isinstance(n, dict):
                continue
            ty = n.get("type")
            if ty in ("section", "notebook"):
                k = _section_key(n)
                if k not in seen:
                    seen.add(k)
                    # 종합은 납작하게(flat) 보여주기
                    out.append({
                        "type": ty,
                        "id": n.get("id") or str(uuid.uuid4()),
                        "name": n.get("name") or (n.get("target") or {}).get("section_text") or (n.get("target") or {}).get("notebook_text") or "항목",
                        "target": n.get("target") or {}
                    })
            # 그룹 아래 children 순회
            ch = n.get("children")
            if isinstance(ch, list):
                _walk_fav_nodes(ch)

    def _walk_buffers(nodes: Any):
        if not isinstance(nodes, list):
            return
        for b in nodes:
            if not isinstance(b, dict):
                continue
            if b.get("type") == "buffer":
                # ✅ 더 이상 aggregate를 스킵하지 않음 (그 안에 직접 등록한 노트북 등 보존 목적)
                _walk_fav_nodes(b.get("data") or [])
            elif b.get("type") == "group":
                _walk_buffers(b.get("children") or [])

    _walk_buffers(bufs)
    # ✅ 종합은 중앙트리에서 '이름순'으로 보이도록 정렬
    try:
        out.sort(key=lambda n: _name_sort_key((n or {}).get("name", "")))
    except Exception:
        pass
    return out


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
_pwa_import_error = ""


def ensure_pywinauto():
    global _pwa_ready, _pwa_import_error, Desktop, WindowNotFoundError, ElementNotFoundError, TimeoutError, UIAWrapper, UIAElementInfo, mouse, keyboard
    # NameError 수정: _ppa_ready -> _pwa_ready
    if _pwa_ready:
        return
    try:
        from pywinauto import (
            Desktop as _Desktop,
            mouse as _mouse,
            keyboard as _keyboard,
        )
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
        _pwa_import_error = ""
    except ImportError as e:
        _pwa_import_error = str(e)
        print(f"[WARN][PWA] import failed: {_pwa_import_error}")


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


# ----------------- 0.3 리소스 경로 헬퍼 (PyInstaller 호환) -----------------
def resource_path(relative_path):
    """
    PyInstaller에서 묶인 리소스 파일을 찾는 경로를 반환합니다.
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# ----------------- 1. 프로세스 실행 파일 경로 얻기 -----------------
def get_process_image_path(pid: int) -> Optional[str]:
    if not pid:
        return None

    now = time.monotonic()
    cached = _PROCESS_IMAGE_PATH_CACHE.get(pid)
    if cached and now < float(cached.get("expires_at", 0.0)):
        return cached.get("path")

    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

    # 64비트 안전: use_last_error로 WinAPI 에러 사용 가능
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    OpenProcess = kernel32.OpenProcess
    OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    OpenProcess.restype = wintypes.HANDLE

    QueryFullProcessImageNameW = kernel32.QueryFullProcessImageNameW
    QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE,
        wintypes.DWORD,
        wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD),
    ]
    QueryFullProcessImageNameW.restype = wintypes.BOOL

    CloseHandle = kernel32.CloseHandle
    CloseHandle.argtypes = [wintypes.HANDLE]
    CloseHandle.restype = wintypes.BOOL

    hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not hProcess:
        _PROCESS_IMAGE_PATH_CACHE[pid] = {
            "expires_at": now + _PROCESS_IMAGE_PATH_CACHE_TTL_SEC,
            "path": None,
        }
        return None
    try:
        # 1차 버퍼
        size = 512
        while True:
            buf_len = wintypes.DWORD(size)
            buf = ctypes.create_unicode_buffer(buf_len.value)
            ok = QueryFullProcessImageNameW(hProcess, 0, buf, ctypes.byref(buf_len))
            if ok:
                path = buf.value
                _PROCESS_IMAGE_PATH_CACHE[pid] = {
                    "expires_at": now + _PROCESS_IMAGE_PATH_CACHE_TTL_SEC,
                    "path": path,
                }
                return path
            # 버퍼 부족 시 한 번 정도 키워 봄
            err = ctypes.get_last_error()
            # ERROR_INSUFFICIENT_BUFFER = 122
            if err == 122 and size < 4096:
                size *= 2
                continue
            _PROCESS_IMAGE_PATH_CACHE[pid] = {
                "expires_at": now + _PROCESS_IMAGE_PATH_CACHE_TTL_SEC,
                "path": None,
            }
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

    # 1. Classic Desktop (OMain*) - 레거시 OneNote
    if "omain" in (cls or "").lower():
        return True

    # 2. Modern App (ApplicationFrameWindow) + 타이틀 키워드
    if cls == ONENOTE_CLASS_NAME and (
        "onenote" in title_lower or "원노트" in title_lower
    ):
        return True

    # 3. Fallback: 제목에 키워드 + EXE 확인
    if "onenote" in title_lower or "원노트" in title_lower:
        exe_path = get_process_image_path(pid)
        if exe_path:
            exe_name = os.path.basename(exe_path).lower()
            if "onenote.exe" in exe_name or "onenoteim.exe" in exe_name:
                return True

    return False


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
    ensure_pywinauto()
    if not _pwa_ready:
        return False
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
def _wrapper_identity_key(ctrl):
    try:
        rect = ctrl.rectangle()
        return (
            _safe_window_text(ctrl),
            _safe_control_type(ctrl),
            rect.left,
            rect.top,
            rect.right,
            rect.bottom,
        )
    except Exception:
        return (id(ctrl),)


def _control_depth_within_tree(ctrl, tree_control) -> int:
    depth = 0
    current = ctrl
    for _ in range(20):
        current = _safe_parent(current)
        if current is None:
            break
        depth += 1
        if current == tree_control:
            break
    return depth


def _pick_best_tree_item_candidate(tree_control, candidates):
    best = None
    best_score = None
    for item in candidates:
        if item is None:
            continue
        try:
            focus = 1 if item.has_keyboard_focus() else 0
        except Exception:
            focus = 0
        try:
            selected = 1 if item.is_selected() else 0
        except Exception:
            selected = 0
        depth = _control_depth_within_tree(item, tree_control)
        try:
            rect = item.rectangle()
            height = max(1, rect.bottom - rect.top)
        except Exception:
            height = 9999
        score = (focus, depth, selected, -height)
        if best_score is None or score > best_score:
            best = item
            best_score = score
    return best


def get_selected_tree_item_fast(tree_control):
    ensure_pywinauto()
    if not _pwa_ready:
        return None
    candidates = []
    seen = set()

    def _push(item):
        if item is None:
            return
        key = _wrapper_identity_key(item)
        if key in seen:
            return
        seen.add(key)
        candidates.append(item)

    def _best_candidate():
        if not candidates:
            return None
        return _pick_best_tree_item_candidate(tree_control, candidates)

    try:
        if hasattr(tree_control, "get_selection"):
            sel = tree_control.get_selection()
            if sel:
                for item in sel:
                    _push(item)
    except Exception:
        pass

    best = _best_candidate()
    if best is not None:
        return best

    try:
        iface_sel = getattr(tree_control, "iface_selection", None)
        if iface_sel:
            arr = iface_sel.GetSelection()
            length = getattr(arr, "Length", 0)
            if length and length > 0:
                for idx in range(length):
                    try:
                        el = arr.GetElement(idx)
                        _push(UIAWrapper(UIAElementInfo(el)))
                    except Exception:
                        continue
    except Exception:
        pass

    best = _best_candidate()
    if best is not None:
        return best

    try:
        for item in tree_control.children():
            try:
                if item.is_selected():
                    _push(item)
            except Exception:
                pass
    except Exception:
        pass

    best = _best_candidate()
    if best is not None:
        return best

    try:
        for control_type in ("TreeItem", "ListItem"):
            for item in tree_control.descendants(control_type=control_type):
                try:
                    if item.is_selected() or item.has_keyboard_focus():
                        _push(item)
                except Exception:
                    pass
    except Exception:
        pass

    return _best_candidate()


# ----------------- 8. 페이지/섹션 컨테이너(Tree/List) 찾기 - ensure 호출 -----------------
def _find_tree_or_list(onenote_window):
    ensure_pywinauto()
    if not _pwa_ready:
        return None
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


def _normalize_notebook_name_key(s: Optional[str]) -> str:
    text = unicodedata.normalize("NFKC", s or "").casefold()
    return re.sub(r"[\s\-_]+", "", text)


def _normalize_project_search_key(s: Optional[str]) -> str:
    text = unicodedata.normalize("NFKC", s or "").casefold()
    return re.sub(r"\s+", "", text)


def _strip_stale_favorite_prefix(text: Optional[str]) -> str:
    raw = (text or "").strip()
    for prefix in ("(구) ", "(old) "):
        if raw.startswith(prefix):
            return raw[len(prefix) :].strip()
    return raw


def select_section_by_text(
    onenote_window, text: str, tree_control: Optional[object] = None
) -> bool:
    ensure_pywinauto()
    if not _pwa_ready:
        return False
    try:
        tree_control = tree_control or _find_tree_or_list(onenote_window)
        if not tree_control:
            return False

        target_norm = _normalize_text(text)

        def _scan(types: List[str]):
            for t in types:
                try:
                    for itm in tree_control.descendants(control_type=t):
                        try:
                            if _normalize_text(itm.window_text()) == target_norm:
                                try:
                                    itm.select()
                                    return True
                                except Exception:
                                    try:
                                        itm.click_input()
                                        return True
                                    except Exception:
                                        pass
                        except Exception:
                            pass
                except Exception:
                    pass
            return False

        if _scan(["TreeItem"]):
            return True
        if _scan(["ListItem"]):
            return True
        return False
    except Exception as e:
        print(f"[ERROR] 섹션 선택 실패: {e}")
        return False


def select_notebook_item_by_text(
    onenote_window,
    text: str,
    tree_control: Optional[object] = None,
    *,
    center_after_select: bool = False,
):
    """
    전자필기장(노트북) 이름으로 찾고 선택합니다.
    - root children 우선 탐색(전자필기장은 보통 루트에 있음)
    - 실패하면 descendants(TreeItem/ListItem)로 fallback
    """
    ensure_pywinauto()
    if not _pwa_ready:
        return False
    try:
        tree_control = tree_control or _find_tree_or_list(onenote_window)
        if not tree_control:
            return False
        target_norm = _normalize_text(text)

        def _select_and_center(item):
            try:
                item.select()
            except Exception:
                try:
                    item.click_input()
                except Exception:
                    return None

            if center_after_select:
                _center_element_in_view(
                    item,
                    tree_control,
                    placement=_get_scroll_placement_for_selected_item(item, tree_control),
                )
            return item

        # 1) root-level children 우선
        try:
            for item in (tree_control.children() or []):
                try:
                    if _normalize_text(item.window_text()) == target_norm:
                        return _select_and_center(item)
                except Exception:
                    continue
        except Exception:
            pass

        # 2) descendants fallback
        for control_type in ["TreeItem", "ListItem"]:
            try:
                for item in tree_control.descendants(control_type=control_type):
                    try:
                        if _normalize_text(item.window_text()) == target_norm:
                            return _select_and_center(item)
                    except Exception:
                        pass
            except Exception:
                pass

        return None
    except Exception as e:
        print(f"[ERROR] 전자필기장 선택 실패: {e}")
        return None


def select_notebook_by_text(
    onenote_window, text: str, tree_control: Optional[object] = None, *, center_after_select: bool = False
) -> bool:
    return (
        select_notebook_item_by_text(
            onenote_window,
            text,
            tree_control,
            center_after_select=center_after_select,
        )
        is not None
    )


def _safe_window_text(ctrl) -> str:
    try:
        return ctrl.window_text() or ""
    except Exception:
        return ""


def _safe_control_type(ctrl) -> str:
    try:
        return ctrl.element_info.control_type or ""
    except Exception:
        return ""


def _safe_rectangle(ctrl):
    try:
        return ctrl.rectangle()
    except Exception:
        return None


def _make_rect_proxy(left: int, top: int, right: int, bottom: int):
    return SimpleNamespace(left=left, top=top, right=right, bottom=bottom)


def _safe_parent(ctrl):
    try:
        return ctrl.parent()
    except Exception:
        return None


def _find_scrollable_ancestor(ctrl, max_depth: int = 8):
    current = ctrl
    fallback = None
    for _ in range(max(1, max_depth)):
        current = _safe_parent(current)
        if current is None:
            break
        fallback = current
        try:
            if getattr(current, "iface_scroll", None) is not None:
                return current
        except Exception:
            pass
        if _safe_control_type(current) in ("List", "Tree"):
            return current
    return fallback


def _find_center_anchor_element(selected_item):
    item_rect = _safe_rectangle(selected_item)
    if item_rect is None:
        return selected_item, "item_rect_missing"

    item_text = _extract_primary_accessible_text(_safe_window_text(selected_item))
    target_norm = _normalize_text(item_text)
    if not target_norm:
        return selected_item, "item_no_text"

    best = None
    best_score = None
    search_types = ("Text", "ListItem", "TreeItem", "Button", "Custom", "DataItem")
    top_limit = item_rect.top + min(260, max(120, (item_rect.bottom - item_rect.top) // 3))

    for cand in _iter_descendants_by_types(selected_item, search_types):
        if cand is None or cand == selected_item:
            continue
        rect = _safe_rectangle(cand)
        if rect is None:
            continue

        height = rect.bottom - rect.top
        width = rect.right - rect.left
        if height < 14 or height > 120 or width < 60:
            continue
        if rect.top < item_rect.top or rect.top > top_limit:
            continue

        cand_text = _extract_primary_accessible_text(_safe_window_text(cand))
        cand_norm = _normalize_text(cand_text)
        if not cand_norm:
            continue

        exact = 1 if cand_norm == target_norm else 0
        partial = 1 if (target_norm in cand_norm or cand_norm in target_norm) else 0
        if not exact and not partial:
            continue

        top_gap = abs(rect.top - item_rect.top)
        score = (exact, partial, -top_gap, width, -height)
        if best_score is None or score > best_score:
            best = cand
            best_score = score

    if best is not None:
        return best, "descendant_text"

    return selected_item, "item"


def _is_root_notebook_tree_item(selected_item, tree_control) -> bool:
    if selected_item is None or tree_control is None:
        return False

    try:
        if _safe_parent(selected_item) == tree_control:
            return True
    except Exception:
        pass

    try:
        for child in tree_control.children() or []:
            if child == selected_item:
                return True
    except Exception:
        pass

    try:
        return _control_depth_within_tree(selected_item, tree_control) <= 1
    except Exception:
        return False


def _get_scroll_placement_for_selected_item(selected_item, tree_control) -> str:
    if _is_root_notebook_tree_item(selected_item, tree_control):
        return "upper"
    return "center"


def _is_already_well_placed_in_view(
    rect_container,
    rect_item,
    rect_anchor,
    *,
    placement: str,
) -> bool:
    if rect_container is None or rect_item is None:
        return False

    container_height = max(1, rect_container.bottom - rect_container.top)
    top_bias = min(48, max(0, int(container_height * 0.08)))
    bottom_bias = min(24, max(0, int(container_height * 0.04)))
    visible_top = rect_container.top + top_bias
    visible_bottom = rect_container.bottom - bottom_bias

    if placement == "upper":
        anchor_top = rect_anchor.top if rect_anchor is not None else rect_item.top
        upper_band_bottom = visible_top + min(96, max(44, int(container_height * 0.22)))
        return visible_top <= anchor_top <= upper_band_bottom

    item_center_y = (
        (rect_anchor.top + rect_anchor.bottom) / 2
        if rect_anchor is not None
        else (rect_item.top + rect_item.bottom) / 2
    )
    container_center_y = (visible_top + visible_bottom) / 2
    return abs(item_center_y - container_center_y) <= 10


def _resolve_alignment_target_for_selected_item(selected_item, tree_control):
    placement = _get_scroll_placement_for_selected_item(selected_item, tree_control)
    if placement == "upper":
        return selected_item, "item_upper_fast", placement

    anchor_element, anchor_source = _find_center_anchor_element(selected_item)
    return anchor_element, anchor_source, placement


def _wait_until(predicate, timeout: float = 1.5, interval: float = 0.05):
    deadline = time.monotonic() + max(0.05, timeout)
    while time.monotonic() < deadline:
        try:
            value = predicate()
            if value:
                return value
        except Exception:
            pass
        time.sleep(interval)
    return None


def _iter_descendants_by_types(root, control_types):
    for control_type in control_types:
        try:
            for item in root.descendants(control_type=control_type):
                yield item
        except Exception:
            continue


def _extract_primary_accessible_text(text: Any) -> str:
    raw = "" if text is None else str(text)
    raw = raw.replace("\r", "\n")
    parts = [part.strip() for part in raw.split("\n") if part.strip()]
    if not parts:
        return raw.strip()
    for part in parts:
        norm = _normalize_text(part)
        if norm in ("shared with: just me", "shared with just me", "just me"):
            continue
        return part
    return parts[0]


def _find_descendant_by_text(root, texts, control_types=None):
    targets = {_normalize_text(t) for t in (texts or []) if t}
    if not targets:
        return None

    search_types = control_types or (
        "Text",
        "Button",
        "MenuItem",
        "ListItem",
        "TreeItem",
        "Hyperlink",
        "Group",
        "Pane",
        "Custom",
    )
    for item in _iter_descendants_by_types(root, search_types):
        txt = _normalize_text(_safe_window_text(item))
        if not txt:
            continue
        if txt in targets:
            return item
        for target in targets:
            if target and (target in txt or txt in target):
                return item
    return None


def _activate_ui_element(ctrl) -> bool:
    try:
        iface_invoke = getattr(ctrl, "iface_invoke", None)
        if iface_invoke is not None:
            iface_invoke.Invoke()
            return True
    except Exception:
        pass

    try:
        if hasattr(ctrl, "double_click_input"):
            ctrl.double_click_input()
            return True
    except Exception:
        pass

    try:
        if hasattr(ctrl, "click_input"):
            ctrl.click_input()
            return True
    except Exception:
        pass

    try:
        if hasattr(ctrl, "select"):
            ctrl.select()
            keyboard.send_keys("{ENTER}")
            return True
    except Exception:
        pass

    return False


def _is_onenote_open_notebook_view(onenote_window) -> bool:
    title = _find_descendant_by_text(
        onenote_window,
        ONENOTE_OPEN_NOTEBOOK_VIEW_TEXTS,
        control_types=("Text", "Button", "Pane", "Group", "Custom"),
    )
    if title:
        return True

    recent = _find_descendant_by_text(
        onenote_window,
        ONENOTE_RECENT_MENU_TEXTS,
        control_types=("Button", "ListItem", "Text", "Hyperlink"),
    )
    search = _find_descendant_by_text(
        onenote_window,
        ONENOTE_SEARCH_TEXTS,
        control_types=("Edit", "Button", "Text", "Custom"),
    )
    return bool(recent and search)


def _ensure_onenote_open_notebook_view(onenote_window) -> bool:
    ensure_pywinauto()
    if not _pwa_ready:
        return False

    try:
        onenote_window.set_focus()
    except Exception:
        pass

    if _is_onenote_open_notebook_view(onenote_window):
        return True

    try:
        keyboard.send_keys("^o")
    except Exception:
        pass
    if _wait_until(lambda: _is_onenote_open_notebook_view(onenote_window), timeout=2.5, interval=0.1):
        return True

    file_item = _find_descendant_by_text(
        onenote_window,
        ONENOTE_FILE_MENU_TEXTS,
        control_types=("Button", "MenuItem", "TabItem", "Hyperlink", "Text"),
    )
    if file_item and _activate_ui_element(file_item):
        time.sleep(0.25)
        open_item = _find_descendant_by_text(
            onenote_window,
            ONENOTE_OPEN_MENU_TEXTS,
            control_types=("Button", "MenuItem", "ListItem", "Hyperlink", "Text"),
        )
        if open_item:
            _activate_ui_element(open_item)
            if _wait_until(lambda: _is_onenote_open_notebook_view(onenote_window), timeout=2.0, interval=0.1):
                return True

    try:
        keyboard.send_keys("^o")
    except Exception:
        pass
    return bool(
        _wait_until(
            lambda: _is_onenote_open_notebook_view(onenote_window),
            timeout=1.5,
            interval=0.1,
        )
    )


def _collect_open_notebook_candidates(onenote_window):
    ensure_pywinauto()
    if not _pwa_ready:
        return [], None

    win_rect = _safe_rectangle(onenote_window)
    if not win_rect:
        return [], None

    left_min = win_rect.left + 60
    left_max = win_rect.left + int((win_rect.right - win_rect.left) * 0.85)
    top_min = win_rect.top + 80
    bottom_max = min(
        win_rect.bottom - 10,
        win_rect.top + int((win_rect.bottom - win_rect.top) * 0.72),
    )

    raw_candidates = []
    seen = set()
    for item in _iter_descendants_by_types(onenote_window, ONENOTE_NOTEBOOK_ITEM_CONTROL_TYPES):
        raw_text = _safe_window_text(item)
        primary_text = _extract_primary_accessible_text(raw_text)
        norm_text = _normalize_text(primary_text)
        if not norm_text:
            continue
        if norm_text in ONENOTE_NOTEBOOK_SKIP_EXACT_TEXTS:
            continue
        if any(skip in norm_text for skip in ONENOTE_NOTEBOOK_SKIP_CONTAINS):
            continue

        rect = _safe_rectangle(item)
        if not rect:
            continue
        width = rect.right - rect.left
        height = rect.bottom - rect.top
        if width < 120 or height < 16 or height > 120:
            continue
        if rect.left < left_min or rect.left > left_max:
            continue
        if rect.top < top_min or rect.bottom > bottom_max:
            continue

        dedupe_key = (norm_text, rect.left // 6, rect.top // 6)
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)

        parent = _safe_parent(item)
        container = _find_scrollable_ancestor(item) or parent
        container_rect = _safe_rectangle(container) if container else None
        container_type = _safe_control_type(container) if container else ""
        raw_candidates.append(
            {
                "item": item,
                "text": primary_text,
                "norm_text": norm_text,
                "rect": rect,
                "parent": parent,
                "container": container,
                "container_rect": container_rect,
                "container_type": container_type,
            }
        )

    if not raw_candidates:
        return [], None

    grouped = {}
    for cand in raw_candidates:
        container_rect = cand.get("container_rect")
        if container_rect is None:
            key = ("", 0, 0, 0, 0)
        else:
            key = (
                cand.get("container_type") or "",
                container_rect.left // 10,
                container_rect.top // 10,
                container_rect.right // 10,
                container_rect.bottom // 10,
            )
        grouped.setdefault(key, []).append(cand)

    best_group = None
    best_score = None
    for group in grouped.values():
        uniq_count = len({cand["norm_text"] for cand in group})
        list_like = sum(
            1
            for cand in group
            if cand.get("container_type") in ("List", "Tree", "Pane", "Group", "Custom")
        )
        score = (uniq_count, list_like, -min(cand["rect"].top for cand in group))
        if best_score is None or score > best_score:
            best_score = score
            best_group = group

    final_candidates = best_group if best_group else raw_candidates
    final_candidates.sort(key=lambda cand: (cand["rect"].top, cand["rect"].left, cand["text"]))
    container = final_candidates[0].get("container") if final_candidates else None
    return final_candidates, container


def _scroll_open_notebook_candidates(container) -> bool:
    if container is None:
        return False

    try:
        container.set_focus()
    except Exception:
        pass

    if _scroll_vertical_via_pattern(container, "down", small=True, repeats=4):
        time.sleep(0.2)
        return True

    try:
        _safe_wheel(container, -4)
        time.sleep(0.2)
        return True
    except Exception:
        pass

    try:
        keyboard.send_keys("{PGDN}")
        time.sleep(0.2)
        return True
    except Exception:
        return False


def _find_next_unattempted_open_notebook(onenote_window, attempted_norms: Set[str]):
    tried_snapshots = set()
    last_candidates = []
    container = None

    for _ in range(24):
        candidates, container = _collect_open_notebook_candidates(onenote_window)
        last_candidates = candidates
        if not candidates:
            return None, [], container

        for cand in candidates:
            if cand["norm_text"] not in attempted_norms:
                return cand, candidates, container

        snapshot = tuple(cand["norm_text"] for cand in candidates[:20])
        if not snapshot or snapshot in tried_snapshots:
            break
        tried_snapshots.add(snapshot)

        if not _scroll_open_notebook_candidates(container):
            break

    return None, last_candidates, container


def _open_one_notebook_from_backstage(
    onenote_window, attempted_norms: Optional[Set[str]] = None
) -> Optional[Dict[str, Any]]:
    if not _ensure_onenote_open_notebook_view(onenote_window):
        return None

    attempted_norms = attempted_norms or set()
    target, candidates, _container = _find_next_unattempted_open_notebook(
        onenote_window, attempted_norms
    )
    if not candidates:
        return {
            "done": False,
            "opened_text": "",
            "visible_names": [],
            "error": "전자 필기장 목록 항목을 찾지 못했습니다.",
        }

    snapshot = [cand["text"] for cand in candidates[:10]]
    if target is None:
        return {"done": True, "opened_text": "", "visible_names": snapshot}

    if not _activate_ui_element(target["item"]):
        return {
            "done": False,
            "opened_text": target["text"],
            "visible_names": snapshot,
            "attempted_norm": target["norm_text"],
            "error": f"전자필기장 열기 실패: '{target['text']}'",
        }

    _wait_until(
        lambda: not _is_onenote_open_notebook_view(onenote_window),
        timeout=2.5,
        interval=0.1,
    )
    time.sleep(0.35)
    return {
        "done": False,
        "opened_text": target["text"],
        "visible_names": snapshot,
        "attempted_norm": target["norm_text"],
    }


def _ps_quote(text: str) -> str:
    return "'" + (text or "").replace("'", "''") + "'"


def _run_powershell(script: str, timeout: int = 30) -> str:
    full_script = (
        "$ErrorActionPreference='Stop';"
        "[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new($false);"
        "$OutputEncoding=[Console]::OutputEncoding;"
        + script
    )
    startupinfo = None
    creationflags = 0
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    completed = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-WindowStyle",
            "Hidden",
            "-Command",
            full_script,
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=max(5, timeout),
        startupinfo=startupinfo,
        creationflags=creationflags,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            (completed.stderr or completed.stdout or "").strip()
            or f"PowerShell exit code {completed.returncode}"
        )
    return (completed.stdout or "").strip()


def _load_json_output(raw_text: str):
    text = (raw_text or "").strip()
    if not text:
        return None
    return json.loads(text)


def _iter_onedrive_notebook_shortcut_dirs() -> List[str]:
    roots = []
    candidates = [
        os.environ.get("OneDrive"),
        os.path.join(os.path.expanduser("~"), "OneDrive"),
    ]
    for base in candidates:
        if not base or not os.path.isdir(base):
            continue
        for rel in ("문서", "Documents"):
            path = os.path.join(base, rel)
            if os.path.isdir(path) and path not in roots:
                roots.append(path)
    return roots


def _read_internet_shortcut_url(path: str) -> str:
    for encoding in ("utf-8-sig", "cp949", "utf-8", "latin-1"):
        try:
            with open(path, "r", encoding=encoding) as f:
                for line in f:
                    if line.startswith("URL="):
                        return line[4:].strip()
        except Exception:
            continue
    return ""


def _looks_like_onenote_shortcut_url(url: str) -> bool:
    lower = (url or "").strip().lower()
    if not lower:
        return False
    return (
        "callerscenarioid=onenote-prod" in lower
        or lower.startswith("onenote:")
        or ("onedrive.live.com/redir.aspx" in lower and "onenote" in lower)
    )


def _extract_onedrive_cid(url: str) -> str:
    try:
        parsed = urlparse(url or "")
        query = parse_qs(parsed.query)
        cid = (query.get("cid") or [""])[0].strip()
        return cid.lower()
    except Exception:
        return ""


def _encode_onenote_protocol_segment(text: str) -> str:
    segment = (text or "").strip()
    return (
        segment.replace("%", "%25")
        .replace("#", "%23")
        .replace(" ", "%20")
    )


def _build_onenote_protocol_url(shortcut_path: str, web_url: str, notebook_name: str) -> str:
    cid = _extract_onedrive_cid(web_url)
    if not cid:
        return ""

    root_label = "문서"
    for root in _iter_onedrive_notebook_shortcut_dirs():
        try:
            if os.path.commonpath([os.path.abspath(shortcut_path), os.path.abspath(root)]) == os.path.abspath(root):
                root_label = os.path.basename(root.rstrip("\\/")) or root_label
                break
        except Exception:
            continue

    encoded_root = _encode_onenote_protocol_segment(root_label)
    encoded_name = _encode_onenote_protocol_segment(notebook_name)
    return f"onenote:https://d.docs.live.net/{cid}/{encoded_root}/{encoded_name}/"


def _get_onenote_exe_path() -> str:
    if winreg is None:
        return ""
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"onenote\shell\Open\command") as key:
            command, _ = winreg.QueryValueEx(key, None)
    except Exception:
        return ""

    command = str(command or "").strip()
    if not command:
        return ""

    if command.startswith('"'):
        end = command.find('"', 1)
        if end > 1:
            exe_path = command[1:end]
        else:
            exe_path = ""
    else:
        exe_path = command.split(" ", 1)[0].strip()

    return exe_path if exe_path and os.path.isfile(exe_path) else ""


def _collect_onenote_notebook_shortcuts() -> List[Dict[str, str]]:
    results: Dict[str, Dict[str, str]] = {}
    for root in _iter_onedrive_notebook_shortcut_dirs():
        try:
            names = sorted(os.listdir(root), key=_name_sort_key)
        except Exception:
            continue
        for name in names:
            if not name.lower().endswith(".url"):
                continue
            path = os.path.join(root, name)
            if not os.path.isfile(path):
                continue
            url = _read_internet_shortcut_url(path)
            if not _looks_like_onenote_shortcut_url(url):
                continue
            display_name = os.path.splitext(name)[0].strip()
            norm_name = _normalize_notebook_name_key(display_name)
            if not norm_name or norm_name in results:
                continue
            results[norm_name] = {
                "name": display_name,
                "path": path,
                "url": url,
            }
    return list(results.values())


def _get_open_notebook_names_via_com(
    refresh: bool = False,
    max_age_sec: float = _OPEN_NOTEBOOK_RECORDS_CACHE_TTL_SEC,
) -> List[str]:
    records = _get_open_notebook_records_via_com(
        refresh=refresh, max_age_sec=max_age_sec
    )
    return [
        str(record.get("name") or "").strip()
        for record in records
        if str(record.get("name") or "").strip()
    ]


def _normalize_notebook_record(raw: Any) -> Optional[Dict[str, str]]:
    if not isinstance(raw, dict):
        return None

    notebook_id = str(raw.get("id") or raw.get("ID") or "").strip()
    name = str(raw.get("name") or "").strip()
    path = str(raw.get("path") or "").strip()

    if not notebook_id and not name and not path:
        return None

    return {
        "id": notebook_id,
        "name": name,
        "path": path,
    }


def _get_open_notebook_records_via_com(
    refresh: bool = False,
    max_age_sec: float = _OPEN_NOTEBOOK_RECORDS_CACHE_TTL_SEC,
) -> List[Dict[str, str]]:
    now = time.monotonic()
    cache_records = _OPEN_NOTEBOOK_RECORDS_CACHE.get("records") or []
    cache_expires_at = float(_OPEN_NOTEBOOK_RECORDS_CACHE.get("expires_at") or 0.0)
    if not refresh and now < cache_expires_at:
        return [dict(record) for record in cache_records]

    script = """
$one = New-Object -ComObject OneNote.Application
$xml = ''
$one.GetHierarchy('', 4, [ref]$xml)
[xml]$doc = $xml
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace('one', 'http://schemas.microsoft.com/office/onenote/2013/onenote')
$items = @(
  $doc.SelectNodes('//one:Notebook', $ns) | ForEach-Object {
    [pscustomobject]@{
      id = $_.GetAttribute('ID')
      name = $_.GetAttribute('name')
      path = $_.GetAttribute('path')
    }
  }
)
$items | ConvertTo-Json -Compress
"""
    data = _load_json_output(_run_powershell(script, timeout=30))
    if data is None:
        _OPEN_NOTEBOOK_RECORDS_CACHE["records"] = []
        _OPEN_NOTEBOOK_RECORDS_CACHE["expires_at"] = now + max(0.0, max_age_sec)
        return []
    if isinstance(data, dict):
        data = [data]

    records: List[Dict[str, str]] = []
    if not isinstance(data, list):
        _OPEN_NOTEBOOK_RECORDS_CACHE["records"] = []
        _OPEN_NOTEBOOK_RECORDS_CACHE["expires_at"] = now + max(0.0, max_age_sec)
        return records

    for entry in data:
        record = _normalize_notebook_record(entry)
        if record and record.get("name"):
            records.append(record)

    _OPEN_NOTEBOOK_RECORDS_CACHE["records"] = [dict(record) for record in records]
    _OPEN_NOTEBOOK_RECORDS_CACHE["expires_at"] = now + max(0.0, max_age_sec)
    return [dict(record) for record in records]


def _pick_notebook_name_suggestion(
    requested_name: str, records: List[Dict[str, str]]
) -> str:
    requested_key = _normalize_notebook_name_key(requested_name)
    if not requested_key:
        return ""

    best_name = ""
    best_score = 0.0
    for record in records:
        name = str(record.get("name") or "").strip()
        name_key = _normalize_notebook_name_key(name)
        if not name_key:
            continue

        score = difflib.SequenceMatcher(None, requested_key, name_key).ratio()
        if requested_key in name_key or name_key in requested_key:
            score = max(score, 0.93)
        if score > best_score:
            best_score = score
            best_name = name

    return best_name if best_score >= 0.72 else ""


def _collect_root_notebook_names_from_tree(tree_control, limit: int = 32) -> List[str]:
    names: List[str] = []
    seen = set()
    if not tree_control:
        return names
    try:
        roots = list(tree_control.children() or [])
    except Exception:
        roots = []

    for item in roots:
        try:
            name = _extract_primary_accessible_text(item.window_text()).strip()
        except Exception:
            name = ""
        key = _normalize_notebook_name_key(name)
        if not key or key in seen:
            continue
        seen.add(key)
        names.append(name)
        if len(names) >= max(1, limit):
            break
    return names


def _build_notebook_not_found_error(
    requested_name: str, candidate_names: List[str]
) -> str:
    shown_name = requested_name or "알 수 없는 전자필기장"
    records = [{"name": name} for name in (candidate_names or []) if name]
    suggestion = _pick_notebook_name_suggestion(requested_name, records)
    if suggestion:
        return (
            f"전자필기장 '{shown_name}'을(를) 찾지 못했습니다. "
            f"이름이 바뀌었을 수 있습니다. 현재 열려 있는 비슷한 이름: '{suggestion}'."
        )
    return (
        f"전자필기장 '{shown_name}'을(를) 찾지 못했습니다. "
        "현재 연결된 OneNote에서 해당 전자필기장이 보이지 않습니다."
    )


def _resolve_notebook_target_for_activation(
    target: Dict[str, Any], fallback_name: str = ""
) -> Dict[str, Any]:
    requested_name = _strip_stale_favorite_prefix(
        str((target or {}).get("notebook_text") or fallback_name or "").strip()
    )
    requested_id = str((target or {}).get("notebook_id") or "").strip()
    result = {
        "requested_name": requested_name,
        "resolved_name": requested_name,
        "notebook_id": requested_id,
        "renamed": False,
        "should_abort": False,
        "com_failed": False,
        "error": "",
    }

    try:
        records = _get_open_notebook_records_via_com()
    except Exception as e:
        result["com_failed"] = True
        result["error"] = str(e)
        return result

    if not records:
        return result

    requested_key = _normalize_notebook_name_key(requested_name)
    records_by_key: Dict[str, List[Dict[str, str]]] = {}
    for record in records:
        record_key = _normalize_notebook_name_key(record.get("name"))
        if record_key:
            records_by_key.setdefault(record_key, []).append(record)

    matched_record = None
    if requested_id:
        matched_record = next(
            (record for record in records if (record.get("id") or "") == requested_id),
            None,
        )

    if matched_record is None and requested_key:
        exact_matches = records_by_key.get(requested_key) or []
        if exact_matches:
            matched_record = exact_matches[0]

    if matched_record is not None:
        resolved_name = str(matched_record.get("name") or requested_name).strip()
        resolved_id = str(matched_record.get("id") or requested_id).strip()
        result["resolved_name"] = resolved_name or requested_name
        result["notebook_id"] = resolved_id
        if requested_name and resolved_name:
            result["renamed"] = (
                _normalize_notebook_name_key(requested_name)
                != _normalize_notebook_name_key(resolved_name)
            )
        return result

    suggestion = _pick_notebook_name_suggestion(requested_name, records)
    shown_name = requested_name or fallback_name or "알 수 없는 전자필기장"
    if suggestion:
        hint = (
            f"전자필기장 '{shown_name}'을(를) 찾지 못했습니다. "
            f"이름이 바뀌었을 수 있습니다. 현재 열려 있는 비슷한 이름: '{suggestion}'."
        )
    else:
        hint = (
            f"전자필기장 '{shown_name}'을(를) 찾지 못했습니다. "
            "이름이 바뀌었거나 현재 열려 있지 않을 수 있습니다. "
            "이름을 바꿨다면 다시 등록해 주세요."
        )

    result["should_abort"] = True
    result["error"] = hint
    return result


def _resolve_favorite_activation_target(
    target: Dict[str, Any], display_name: str
) -> Dict[str, Any]:
    notebook_text = str((target or {}).get("notebook_text") or "").strip()
    section_text = str((target or {}).get("section_text") or "").strip()
    result = {
        "ok": True,
        "target_kind": None,
        "expected_center_text": "",
        "resolved_name": "",
        "resolved_notebook_id": "",
        "error": "",
    }

    if notebook_text:
        notebook_info = _resolve_notebook_target_for_activation(target, display_name)
        result["target_kind"] = "notebook"
        result["resolved_name"] = notebook_info.get("resolved_name") or notebook_text
        result["resolved_notebook_id"] = notebook_info.get("notebook_id") or ""
        if notebook_info.get("should_abort"):
            result["ok"] = False
            result["error"] = notebook_info.get("error") or ""
            return result
        result["expected_center_text"] = result["resolved_name"] or notebook_text
        return result

    if section_text:
        result["target_kind"] = "section"
        result["expected_center_text"] = section_text
        return result

    result["target_kind"] = "notebook"
    result["expected_center_text"] = _strip_stale_favorite_prefix(display_name)
    return result


def _open_notebook_shortcut_via_shell(
    shortcut_path: str,
    url: str,
    expected_name: str,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]:
    expected_key = _normalize_notebook_name_key(expected_name)
    protocol_url = _build_onenote_protocol_url(shortcut_path, url, expected_name)
    try:
        open_keys = {
            _normalize_notebook_name_key(name)
            for name in _get_open_notebook_names_via_com(refresh=True)
            if _normalize_notebook_name_key(name)
        }
    except Exception:
        open_keys = set()

    if expected_key and expected_key in open_keys:
        return {"ok": True, "already": True, "name": expected_name, "error": ""}

    launch_errors = []
    if progress_callback:
        try:
            progress_callback(f"OneNote 실행 요청 중... {expected_name}")
        except Exception:
            pass

    if protocol_url:
        exe_path = _get_onenote_exe_path()
        _clear_open_notebook_records_cache()
        try:
            if exe_path:
                subprocess.Popen([exe_path, "/hyperlink", protocol_url])
            else:
                os.startfile(protocol_url)
        except Exception as e:
            launch_errors.append(str(e))
            try:
                if exe_path:
                    os.startfile(protocol_url)
                else:
                    _run_powershell(
                        f"Start-Process -FilePath {_ps_quote(protocol_url)}",
                        timeout=10,
                    )
            except Exception as e2:
                launch_errors.append(str(e2))
    else:
        launch_errors.append("OneNote 앱 프로토콜 URL 생성 실패")

    for _ in range(20):
        time.sleep(0.5)
        wait_round = _ + 1
        if progress_callback and (wait_round == 1 or wait_round % 4 == 0):
            try:
                progress_callback(
                    f"OneNote 응답 대기 중... {expected_name} ({wait_round / 2:.1f}초)"
                )
            except Exception:
                pass
        try:
            open_keys = {
                _normalize_notebook_name_key(name)
                for name in _get_open_notebook_names_via_com(refresh=True)
                if _normalize_notebook_name_key(name)
            }
        except Exception:
            continue
        if expected_key and expected_key in open_keys:
            return {
                "ok": True,
                "already": False,
                "name": expected_name,
                "error": "",
            }

    return {
        "ok": False,
        "already": False,
        "name": expected_name,
        "error": "; ".join(msg for msg in launch_errors if msg) or "OneNote 앱에서 열기 실패",
    }


# ----------------- 9. 요소를 중앙으로 위치시키는 함수(최적화) - ensure 호출 -----------------
def _center_element_in_view(
    element_to_center,
    scroll_container,
    *,
    anchor_element=None,
    placement: str = "center",
):
    ensure_pywinauto()
    if not _pwa_ready:
        return
    try:
        try:
            element_to_center.iface_scroll_item.ScrollIntoView()
        except AttributeError:
            return

        settle_timeout = 0.03 if placement == "upper" else 0.1
        settle_interval = 0.01 if placement == "upper" else 0.015
        _wait_rect_settle(
            lambda: element_to_center.rectangle(),
            timeout=settle_timeout,
            interval=settle_interval,
        )

        effective_container = (
            _find_scrollable_ancestor(element_to_center) or scroll_container
        )

        def _anchor_metrics(rect_container, rect_item, rect_anchor):
            if rect_container is None or rect_item is None:
                return rect_item, 0.0, "full"

            if rect_anchor is not None:
                anchor_height = max(1, rect_anchor.bottom - rect_anchor.top)
                if 14 <= anchor_height <= 140:
                    return (
                        rect_anchor,
                        (rect_anchor.top + rect_anchor.bottom) / 2,
                        "anchor_element",
                    )

            container_height = max(1, rect_container.bottom - rect_container.top)
            item_height = max(1, rect_item.bottom - rect_item.top)
            row_height = max(44, min(88, int(container_height * 0.055)))
            if item_height > max(220, int(container_height * 0.85)):
                anchor_bottom = min(rect_item.bottom, rect_item.top + row_height)
                anchor = _make_rect_proxy(
                    rect_item.left,
                    rect_item.top,
                    rect_item.right,
                    anchor_bottom,
                )
                return anchor, (rect_item.top + anchor_bottom) / 2, "top_slice"

            return rect_item, (rect_item.top + rect_item.bottom) / 2, "full"

        def _calc_offset():
            rect_container = _safe_rectangle(effective_container)
            rect_item = _safe_rectangle(element_to_center)
            rect_anchor = _safe_rectangle(anchor_element) if anchor_element else None
            if rect_container is None or rect_item is None:
                return None, None, None, 0.0, "full"
            anchor_rect, item_center_y, anchor_mode = _anchor_metrics(
                rect_container, rect_item, rect_anchor
            )
            container_height = max(1, rect_container.bottom - rect_container.top)
            top_bias = min(48, max(0, int(container_height * 0.08)))
            bottom_bias = min(24, max(0, int(container_height * 0.04)))
            visible_top = rect_container.top + top_bias
            visible_bottom = rect_container.bottom - bottom_bias
            if placement == "upper":
                anchor_top = anchor_rect.top if anchor_rect is not None else rect_item.top
                target_y = visible_top + min(28, max(10, int(container_height * 0.03)))
                offset = anchor_top - target_y
            else:
                container_center_y = (visible_top + visible_bottom) / 2
                offset = item_center_y - container_center_y
            return (
                rect_container,
                rect_item,
                anchor_rect,
                offset,
                anchor_mode,
            )

        rect_container, rect_item, anchor_rect, offset, anchor_mode = _calc_offset()
        print(
            "[DBG][CENTER][GEOM]",
            f"phase=initial",
            f"placement={placement}",
            f"anchor={anchor_mode}",
            f"offset={offset:.1f}",
            f"container={rect_container}",
            f"item={rect_item}",
            f"anchor_rect={anchor_rect}",
        )

        if _is_already_well_placed_in_view(
            rect_container, rect_item, anchor_rect, placement=placement
        ):
            return

        if abs(offset) <= 10:
            return

        def step_for(dy):
            item_height = 28
            if anchor_rect is not None:
                item_height = min(96, max(20, anchor_rect.bottom - anchor_rect.top))
            elif rect_item is not None:
                item_height = min(96, max(20, rect_item.bottom - rect_item.top))
            return max(1, min(8, int(abs(dy) / max(item_height, 20))))

        max_loops = 3 if placement == "upper" else 6
        for _ in range(max_loops):
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

            time.sleep(0.005 if placement == "upper" else 0.01)

            (
                rect_container,
                rect_item,
                anchor_rect,
                offset,
                anchor_mode,
            ) = _calc_offset()

            if _is_already_well_placed_in_view(
                rect_container, rect_item, anchor_rect, placement=placement
            ):
                break

        print(
            "[DBG][CENTER][GEOM]",
            f"phase=final",
            f"placement={placement}",
            f"anchor={anchor_mode}",
            f"offset={offset:.1f}",
            f"container={rect_container}",
            f"item={rect_item}",
            f"anchor_rect={anchor_rect}",
        )

    except Exception as e:
        print(f"[WARN] 중앙 정렬 중 오류: {e}")


# ----------------- 10. 선택된 항목을 중앙으로 스크롤 -----------------
def scroll_selected_item_to_center(
    onenote_window,
    tree_control: Optional[object] = None,
    *,
    selected_item=None,
):
    ensure_pywinauto()
    if not _pwa_ready:
        return False, None
    try:
        tree_control = tree_control or _find_tree_or_list(onenote_window)
        if not tree_control:
            return False, None

        selected_item = selected_item or get_selected_tree_item_fast(tree_control)
        if not selected_item:
            print("[DBG][CENTER][TARGET] selected_item=None")
            return False, None

        item_name = selected_item.window_text()
        try:
            has_focus = bool(selected_item.has_keyboard_focus())
        except Exception:
            has_focus = False
        try:
            is_selected = bool(selected_item.is_selected())
        except Exception:
            is_selected = False
        depth = _control_depth_within_tree(selected_item, tree_control)
        rect = _safe_rectangle(selected_item)
        height = None if rect is None else max(1, rect.bottom - rect.top)
        anchor_element, anchor_source, placement = _resolve_alignment_target_for_selected_item(
            selected_item, tree_control
        )
        anchor_text = _safe_window_text(anchor_element)
        anchor_rect = _safe_rectangle(anchor_element)
        anchor_height = None if anchor_rect is None else max(1, anchor_rect.bottom - anchor_rect.top)
        print(
            "[DBG][CENTER][TARGET]",
            f"text={item_name!r}",
            f"type={_safe_control_type(selected_item)!r}",
            f"depth={depth}",
            f"height={height}",
            f"placement={placement}",
            f"anchor_source={anchor_source}",
            f"anchor_text={anchor_text!r}",
            f"anchor_height={anchor_height}",
            f"selected={is_selected}",
            f"focus={has_focus}",
        )
        _center_element_in_view(
            selected_item,
            tree_control,
            anchor_element=anchor_element,
            placement=placement,
        )
        return True, item_name
    except (ElementNotFoundError, TimeoutError):
        return False, None
    except Exception:
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
        current_settings = load_settings()
        if current_settings.get("connection_signature") == info:
            return
        current_settings["connection_signature"] = info
        save_settings(current_settings)
    except Exception as e:
        print(f"[ERROR] 연결 정보 저장 실패: {e}")


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
    if not _pwa_ready:
        return None
    h = sig.get("handle")
    if h:
        try:
            w = Desktop(backend="uia").window(handle=h)
            if w.is_visible():
                return w
        except Exception:
            pass

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
    settings = load_settings()
    sig = settings.get("connection_signature")
    if not sig:
        return None, "연결되지 않음"
    try:
        win = reacquire_window_by_signature(sig)
        if win and win.is_visible():
            window_title = win.window_text()
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
            win, status = load_connection_info_and_reconnect()
            if win:
                payload = {
                    "ok": True,
                    "status": status,
                    "sig": build_window_signature(win),
                }
            else:
                payload = {"ok": False, "status": status}
        except Exception as e:
            payload = {"ok": False, "status": f"연결되지 않음 (오류: {e})"}
        self.finished.emit(payload)


# ----------------- 3-A. OneNote 창 목록 스캔 워커 -----------------
class OneNoteWindowScanner(QThread):
    done = pyqtSignal(list)

    def __init__(self, my_pid: int, parent=None):
        super().__init__(parent)
        self.my_pid = my_pid

    def run(self):
        results = []
        try:
            wins = enum_windows_fast(filter_title_substr=None)
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

    def run(self):
        try:
            ensure_pywinauto()
            if not _pwa_ready:
                self.done.emit(False, "")
                return

            win = reacquire_window_by_signature(self.sig)
            if not win:
                self.done.emit(False, "")
                return

            tree = _find_tree_or_list(win)
            if not tree:
                self.done.emit(False, "")
                return

            expected_norm = _normalize_text(self.expected_text)
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
            if not _pwa_ready:
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
            if not tree:
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


class OpenAllNotebooksWorker(QThread):
    progress = pyqtSignal(str)
    done = pyqtSignal(dict)

    def __init__(self, sig: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.sig = copy.deepcopy(sig or {})

    def run(self):
        result = {
            "ok": False,
            "window_info": None,
            "opened_count": 0,
            "opened_names": [],
            "remaining_names": [],
            "error": "",
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

            try:
                result["window_info"] = {
                    "handle": win.handle,
                    "title": win.window_text(),
                    "class_name": win.class_name(),
                    "pid": win.process_id(),
                }
            except Exception:
                result["window_info"] = None

            shortcut_targets = _collect_onenote_notebook_shortcuts()
            if shortcut_targets:
                try:
                    open_names = {
                        _normalize_notebook_name_key(name)
                        for name in _get_open_notebook_names_via_com(refresh=True)
                        if _normalize_notebook_name_key(name)
                    }
                except Exception as e:
                    print(f"[WARN][OPEN_ALL_NOTEBOOKS][COM][LIST] {e}")
                    open_names = set()

                pending_targets = [
                    t
                    for t in shortcut_targets
                    if _normalize_notebook_name_key(t.get("name")) not in open_names
                ]

                total_targets = len(pending_targets)
                self.progress.emit(
                    f"실제 OneNote 전체 열기 준비 완료... 대상 {total_targets}개"
                )

                if not pending_targets:
                    result["ok"] = True
                    result["remaining_names"] = []
                    self.done.emit(result)
                    return

                failed_names = []
                failed_details = []

                for index, target in enumerate(pending_targets, start=1):
                    if self.isInterruptionRequested():
                        result["error"] = "사용자 중단"
                        self.done.emit(result)
                        return

                    name = (target.get("name") or "").strip()
                    path = (target.get("path") or "").strip()
                    url = (target.get("url") or "").strip()
                    if not name or (not path and not url):
                        failed_names.append(name or "이름 없음")
                        failed_details.append(f"{name or '이름 없음'}: 바로가기 정보 없음")
                        continue

                    self.progress.emit(
                        f"실제 OneNote 전체 열기 진행 중... {index}/{total_targets} 시도 - {name}"
                    )
                    try:
                        step = _open_notebook_shortcut_via_shell(
                            path,
                            url,
                            name,
                            progress_callback=lambda msg, idx=index, total=total_targets: self.progress.emit(
                                f"실제 OneNote 전체 열기 진행 중... {idx}/{total} - {msg}"
                            ),
                        )
                    except Exception as e:
                        step = {"ok": False, "already": False, "name": name, "error": str(e)}

                    if step.get("ok"):
                        result["opened_names"].append(name)
                        result["opened_count"] += 1
                        self.progress.emit(
                            f"실제 OneNote 전체 열기 진행 중... {index}/{total_targets} 완료 - {name}"
                        )
                    else:
                        failed_names.append(name)
                        failed_details.append(
                            f"{name}: {step.get('error') or '열기 실패'}"
                        )

                if failed_names:
                    result["error"] = (
                        f"바로가기 기반 열기 실패 {len(failed_names)}개 - "
                        + "; ".join(failed_details[:3])
                    )
                    result["remaining_names"] = failed_names
                    self.done.emit(result)
                    return

                result["ok"] = True
                result["remaining_names"] = []
                self.done.emit(result)
                return

            last_snapshot = None
            stale_rounds = 0
            attempted_norms: Set[str] = set()

            for _ in range(200):
                if self.isInterruptionRequested():
                    result["error"] = "사용자 중단"
                    self.done.emit(result)
                    return

                step = _open_one_notebook_from_backstage(win, attempted_norms)
                if step is None:
                    result["error"] = "OneNote의 '전자 필기장 열기' 화면을 찾지 못했습니다."
                    self.done.emit(result)
                    return

                visible_names = step.get("visible_names") or []
                if step.get("done"):
                    result["ok"] = True
                    result["remaining_names"] = []
                    self.done.emit(result)
                    return

                attempted_norm = (step.get("attempted_norm") or "").strip()
                current_snapshot = tuple(_normalize_text(name) for name in visible_names[:8])
                if (not attempted_norm) and current_snapshot and current_snapshot == last_snapshot:
                    stale_rounds += 1
                else:
                    stale_rounds = 0
                last_snapshot = current_snapshot

                if attempted_norm:
                    attempted_norms.add(attempted_norm)

                opened_text = (step.get("opened_text") or "").strip()
                if opened_text:
                    result["opened_names"].append(opened_text)
                    result["opened_count"] += 1
                    self.progress.emit(
                        f"실제 OneNote 전체 열기 중... {result['opened_count']}개 - {opened_text}"
                    )

                if step.get("error"):
                    result["error"] = step["error"]
                    result["remaining_names"] = visible_names
                    self.done.emit(result)
                    return

                if stale_rounds >= 2:
                    result["error"] = "전자필기장 목록이 더 이상 변하지 않아 중단했습니다."
                    result["remaining_names"] = visible_names
                    self.done.emit(result)
                    return

            result["error"] = "안전 제한(200개)에 도달해 중단했습니다."
            self.done.emit(result)
        except Exception as e:
            print(f"[WARN][OPEN_ALL_NOTEBOOKS][WORKER] {e}")
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


# ----------------- 14-A. 즐겨찾기 트리 위젯 (삭제 - src.ui.widgets에서 임포트) -----------------
# class FavoritesTree(QTreeWidget):
#     ... (삭제됨) ...


# ----------------- 14. PyQt GUI -----------------
class OneNoteScrollRemoconApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self._t_boot = time.perf_counter()
        self._boot_marks = []

        def _mark(label: str):
            now = time.perf_counter()
            self._boot_marks.append((label, (now - self._t_boot) * 1000.0))

        self._boot_mark = _mark
        self._boot_mark("QMainWindow.__init__ done")
        # 1. 설정 로드 및 창 위치/상태 복원
        self.settings = load_settings()
        self._boot_mark("load_settings done")
        self.onenote_window = None
        self.tree_control = None
        self._reconnect_worker = None
        self._scanner_worker = None
        self._pending_center_timer: Optional[QTimer] = None
        self._center_worker: Optional[CenterAfterActivateWorker] = None
        self._favorite_activation_worker: Optional[FavoriteActivationWorker] = None
        self._open_all_notebooks_worker: Optional[OpenAllNotebooksWorker] = None
        self._codex_location_lookup_worker: Optional[CodexLocationLookupWorker] = None
        self._retained_qthreads: List[QThread] = []
        self._center_request_seq = 0
        self._last_list_connect_key = None
        self._last_list_connect_at = 0.0
        self._pending_onenote_list_selection_key = None
        self._last_onenote_list_refresh_at = 0.0
        self.onenote_windows_info: List[Dict] = []
        self.my_pid = os.getpid()
        self._auto_center_after_activate = True
        self.active_buffer_id = None
        # 현재 활성 버퍼의 데이터(payload) 및 해당 트리 아이템
        # NOTE: PyQt의 item.data()로 꺼낸 dict는 수정해도 item 내부에 반영되지 않는 경우가 있어,
        #       저장 시 반드시 item.setData()로 payload를 다시 주입한다.
        self.active_buffer_node = None  # Dict payload
        self.active_buffer_item = None  # QTreeWidgetItem
        self._active_buffer_settings_node = None  # Dict node in self.settings
        self._last_loaded_center_buffer_id = None
        self._buffer_item_index: Dict[str, QTreeWidgetItem] = {}
        self._first_buffer_item: Optional[QTreeWidgetItem] = None
        self._buffer_search_highlight_bg = QBrush(QColor("#6d5a1f"))
        self._buffer_search_highlight_fg = QBrush(QColor("#fff3bf"))
        self._buffer_search_clear_bg = QBrush()
        self._buffer_search_clear_fg = QBrush()
        self._buffer_search_index: List[Dict[str, Any]] = []
        self._buffer_search_last_match_records: List[Dict[str, Any]] = []
        self._buffer_search_highlighted_by_id: Dict[int, QTreeWidgetItem] = {}
        self._module_search_index: List[Dict[str, Any]] = []
        self._module_search_last_match_records: List[Dict[str, Any]] = []
        self._module_search_highlighted_by_id: Dict[int, QTreeWidgetItem] = {}
        self._buffer_search_match_count = 0
        self._module_search_match_count = 0
        self._buffer_search_pending_text = ""
        self._buffer_search_pending_key = ""
        self._buffer_search_last_applied_key = ""
        self._buffer_search_last_first_match_id = 0
        self._module_search_last_first_match_id = 0
        self._buffer_search_timer = QTimer(self)
        self._buffer_search_timer.setSingleShot(True)
        self._buffer_search_timer.timeout.connect(
            lambda: self._highlight_project_buffers_from_module_search(
                self._buffer_search_pending_text,
                precomputed_query=self._buffer_search_pending_key,
            )
        )
        self._buffer_save_timer = QTimer(self)
        self._buffer_save_timer.setSingleShot(True)
        self._buffer_save_timer.timeout.connect(self._save_buffer_structure)
        self._buffer_save_interval_ms = 120
        self._aggregate_cache_valid = False
        self._aggregate_cache = []
        self._aggregate_display_cache_sig = None
        self._aggregate_display_cache = []
        self._aggregate_display_cache_kind = None
        self._aggregate_display_cache_source_id = 0
        self._aggregate_classified_keys_cache_valid = False
        self._aggregate_classified_keys_cache: Set[str] = set()
        self._aggregate_reclassify_in_progress = False
        self._settings_save_timer = QTimer(self)
        self._settings_save_timer.setSingleShot(True)
        self._settings_save_timer.timeout.connect(self._flush_pending_settings_save)
        self._settings_save_interval_ms = 180
        self._settings_save_pending = False
        self._settings_save_in_progress = False
        self._onenote_list_refresh_timer = QTimer(self)
        self._onenote_list_refresh_timer.setSingleShot(True)
        self._onenote_list_refresh_timer.timeout.connect(
            self._refresh_onenote_list_from_click
        )

        # --- [START] 창 위치 복원 및 유효성 검사 로직 (수정됨) ---
        geo_settings = self.settings.get(
            "window_geometry", DEFAULT_SETTINGS["window_geometry"]
        )

        # 주 모니터의 사용 가능한 영역 가져오기 (작업 표시줄 제외)
        primary_screen = QApplication.primaryScreen()
        if not primary_screen:  # 헤드리스 환경 등 예외 처리
            # 기본 가상 화면 크기 설정
            screen_rect = QRect(0, 0, 1920, 1080)
        else:
            screen_rect = primary_screen.availableGeometry()

        # 저장된 창 위치 QRect 객체로 생성
        window_rect = QRect(
            geo_settings.get("x", 200),
            geo_settings.get("y", 180),
            geo_settings.get("width", 960),
            geo_settings.get("height", 540),
        )

        # 창이 화면에 보이는지 확인 (최소 100x50 픽셀이 보여야 함)
        intersection = screen_rect.intersected(window_rect)
        is_visible = intersection.width() >= 100 and intersection.height() >= 50

        if not is_visible:
            # 창이 화면 밖에 있으면 화면 중앙으로 이동
            # 창 크기는 유지하되, 화면 크기보다 크지 않도록 조정
            window_rect.setWidth(min(window_rect.width(), screen_rect.width()))
            window_rect.setHeight(min(window_rect.height(), screen_rect.height()))
            # 중앙 정렬
            window_rect.moveCenter(screen_rect.center())

        self.setGeometry(window_rect)
        # --- [END] 창 위치 복원 및 유효성 검사 로직 ---

        # 즐겨찾기 복사 데이터 임시 저장소 (클립보드 역할)
        self.clipboard_data: Optional[Dict] = None

        # 즐겨찾기 버퍼 복사 데이터 임시 저장소
        self.buffer_clipboard_data: Optional[Dict] = None

        # --- FavoritesTree Undo/Redo (Ctrl+Z / Ctrl+Shift+Z / Ctrl+X) ---
        self._fav_undo_stack: List[str] = []
        self._fav_redo_stack: List[str] = []
        self._fav_last_snapshot: Optional[str] = None
        self._fav_undo_batch_final_snapshot: Optional[str] = None
        self._fav_undo_suspended: bool = False
        self._fav_last_persisted_hash: Optional[str] = None
        self._last_center_payload_hash: Optional[str] = None
        self._last_center_payload_snapshot: Optional[str] = None
        self._last_center_payload_source_id: int = 0
        self._last_saved_buffer_structure_sig: Optional[str] = None
        self._fav_save_timer = QTimer(self)
        self._fav_save_timer.setSingleShot(True)
        self._fav_save_timer.timeout.connect(self._flush_pending_favorites_save)
        self._fav_save_interval_ms: int = 120
        self._fav_save_pending: bool = False
        self._fav_undo_max: int = 80
        # bulk operation에서 (다중 붙여넣기/삭제/잘라내기 등) Ctrl+Z가 "한 개씩" 되돌아가는 문제를 막기 위해
        # Undo/Redo를 "트랜잭션"처럼 한 번에 묶어 처리한다.
        self._fav_undo_batch_depth: int = 0
        self._fav_undo_batch_base_snapshot: Optional[str] = None
        self._fav_undo_batch_reason: str = ""
        self._debug_hotpaths = bool(self.settings.get("debug_hotpaths", False))
        self._debug_perf_logs = bool(self.settings.get("debug_perf_logs", False))

        # 1.1 애플리케이션 아이콘 설정
        icon_path = resource_path(APP_ICON_PATH)
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.init_ui("로딩 중...")
        self._boot_mark("init_ui done")

        # NOTE:
        #   실제 체감 지체는 w.show() 내부(첫 레이아웃/폴리시/폰트 계산 등)에서 크게 발생할 수 있다.
        #   따라서 무거운 초기화는 "첫 show로 1프레임 그린 뒤" 실행한다.
        self._bootstrap_scheduled = False

    def _dbg_hot(self, *args, **kwargs):
        if self._debug_hotpaths:
            print(*args, **kwargs)

    def _dbg_perf(self, *args, **kwargs):
        if self._debug_perf_logs:
            print(*args, **kwargs)

    def showEvent(self, e):
        super().showEvent(e)
        if not getattr(self, "_bootstrap_scheduled", False):
            self._bootstrap_scheduled = True
            QTimer.singleShot(10, self._deferred_bootstrap)  # allow first paint

    def _deferred_bootstrap(self):
        # 첫 페인트 이후에 돌리되, 작업 중 불필요한 리페인트/레이아웃을 줄인다.
        try:
            # 2. 즐겨찾기 버퍼 및 즐겨찾기 로드
            t0 = time.perf_counter()
            self._load_buffers_and_favorites()
            self._boot_mark(f"_load_buffers_and_favorites done (+{(time.perf_counter()-t0)*1000.0:.1f}ms)")

            # OneNote/pywinauto 쪽은 여기서부터 시작해도 충분 (필요 시 내부에서 ensure_pywinauto()가 또 호출됨)
            QTimer.singleShot(0, self.refresh_onenote_list)
            QTimer.singleShot(0, self._start_auto_reconnect)
            self._boot_mark("timers scheduled")

            # FIX: 앱 시작 시 저장된 버퍼 기준으로 2패널 강제 리빌드
            QTimer.singleShot(50, self._finish_boot_sequence)
        except Exception as e:
            print(f"[BOOT][ERROR] deferred bootstrap failed: {e}")
            traceback.print_exc()
            try:
                self.connection_status_label.setText(f"부팅 로드 실패: {e}")
            except Exception:
                pass
        finally:
            self.setUpdatesEnabled(True)
            self.update()

        # 부팅 구간 로그 출력
        try:
            self._dbg_perf("[BOOT][PERF] ---- startup marks ----")
            for label, ms in self._boot_marks:
                self._dbg_perf(f"[BOOT][PERF] {ms:8.1f} ms | {label}")
            self._dbg_perf("[BOOT][PERF] ------------------------")
        except Exception:
            pass

        self.connection_status_label.setText("준비됨 (자동 재연결 중...)")

    def _ps_single_quoted(self, value: str) -> str:
        return "'" + (value or "").replace("'", "''") + "'"

    def _codex_codegen_values(self) -> Dict[str, str]:
        profile = self._codex_target_from_fields()
        title_input = getattr(self, "codex_request_title_input", None)
        body_editor = getattr(self, "codex_request_body_editor", None)
        target_input = getattr(self, "codex_request_target_input", None)

        title = title_input.text().strip() if title_input is not None else ""
        body = body_editor.toPlainText().strip() if body_editor is not None else ""
        target = target_input.text().strip() if target_input is not None else ""

        return {
            "title": title or "코덱스 작업",
            "body": body or "코덱스가 작성한 메모입니다.",
            "target": target or profile.get("path", ""),
            "notebook": profile.get("notebook", ""),
            "section_group": profile.get("section_group", ""),
            "section": profile.get("section", ""),
            "section_group_id": (
                profile.get("section_group_id")
                or "{2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}"
            ),
            "section_id": (
                profile.get("section_id")
                or "{175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}"
            ),
        }

    def _codex_target_profile_from_json_text(self, text: str) -> Dict[str, str]:
        raw = (text or "").strip()
        if not raw:
            raise ValueError("클립보드에 작업 위치 정보가 없습니다.")

        try:
            data = json.loads(raw)
        except Exception:
            starts = [idx for idx in (raw.find("{"), raw.find("[")) if idx >= 0]
            ends = [idx for idx in (raw.rfind("}"), raw.rfind("]")) if idx >= 0]
            if not starts or not ends:
                raise ValueError("작업 위치 정보의 시작과 끝을 찾지 못했습니다.")
            data = json.loads(raw[min(starts): max(ends) + 1])

        if isinstance(data, dict) and isinstance(data.get("targets"), list):
            data = data["targets"][0] if data["targets"] else {}
        elif isinstance(data, list):
            data = data[0] if data else {}
        if not isinstance(data, dict):
            raise ValueError("작업 위치 정보 형식이 올바르지 않습니다.")

        current = self._codex_target_from_fields()

        def pick(key: str, fallback_key: Optional[str] = None) -> str:
            value = data.get(key)
            if value is None and fallback_key:
                value = data.get(fallback_key)
            if value is None:
                value = current.get(key, "")
            return str(value or "").strip()

        return {
            "name": pick("name") or "새 대상",
            "path": pick("path"),
            "notebook": pick("notebook"),
            "section_group": pick("section_group", "sectionGroup"),
            "section": pick("section"),
            "section_group_id": pick("section_group_id", "sectionGroupId"),
            "section_id": pick("section_id", "sectionId"),
        }

    def _save_codex_target_from_clipboard_json(self) -> None:
        try:
            profile = self._codex_target_profile_from_json_text(
                QApplication.clipboard().text()
            )
            self._populate_codex_target_fields(profile)
            self._apply_codex_target_to_request()
            self._save_codex_target_profile()
        except Exception as e:
            QMessageBox.warning(self, "작업 위치 적용 실패", str(e))

    def _codex_onenote_location_lookup_script(self) -> str:
        return """# OneNote COM: 작업 위치 상세 조회
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$xml = ""
$one.GetHierarchy(
    "",
    [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections,
    [ref]$xml
)

[xml]$doc = $xml
$nsUri = $doc.DocumentElement.NamespaceURI
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("one", $nsUri)

$targets = New-Object System.Collections.Generic.List[object]

function Join-TargetPath {
    param([string[]]$Parts)
    $clean = @($Parts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    return [string]::Join(" > ", [string[]]$clean)
}

function Add-Target {
    param(
        [string]$Kind,
        [string]$DisplayName,
        [string]$Path,
        [string]$Notebook,
        [string]$SectionGroup,
        [string]$Section,
        [string]$SectionGroupId,
        [string]$SectionId
    )
    $targets.Add([ordered]@{
        kind = $Kind
        name = $DisplayName
        path = $Path
        notebook = $Notebook
        section_group = $SectionGroup
        section = $Section
        section_group_id = $SectionGroupId
        section_id = $SectionId
    })
}

function Visit-OneNoteNode {
    param(
        [System.Xml.XmlNode]$Node,
        [string[]]$PathParts,
        [string]$NotebookName,
        [string]$SectionGroupPath,
        [string]$ParentContainerId
    )

    $local = $Node.LocalName
    if ($local -notin @("Notebook", "SectionGroup", "Section")) {
        return
    }

    $nodeName = $Node.GetAttribute("name")
    if ([string]::IsNullOrWhiteSpace($nodeName)) {
        $nodeName = "(이름 없음)"
    }
    $nodeId = $Node.GetAttribute("ID")

    if ($local -eq "Notebook") {
        $nextPathParts = @($nodeName)
        $path = Join-TargetPath -Parts $nextPathParts
        Add-Target `
            -Kind "notebook" `
            -DisplayName "전자필기장 - $nodeName" `
            -Path $path `
            -Notebook $nodeName `
            -SectionGroup "" `
            -Section "" `
            -SectionGroupId $nodeId `
            -SectionId ""

        foreach ($child in $Node.ChildNodes) {
            if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                Visit-OneNoteNode `
                    -Node $child `
                    -PathParts $nextPathParts `
                    -NotebookName $nodeName `
                    -SectionGroupPath "" `
                    -ParentContainerId $nodeId
            }
        }
        return
    }

    $nextPathParts = @($PathParts + $nodeName)
    $path = Join-TargetPath -Parts $nextPathParts

    if ($local -eq "SectionGroup") {
        $groupPath = if ([string]::IsNullOrWhiteSpace($SectionGroupPath)) {
            $nodeName
        } else {
            "$SectionGroupPath > $nodeName"
        }
        Add-Target `
            -Kind "section_group" `
            -DisplayName "그룹 - $nodeName" `
            -Path $path `
            -Notebook $NotebookName `
            -SectionGroup $groupPath `
            -Section "" `
            -SectionGroupId $nodeId `
            -SectionId ""

        foreach ($child in $Node.ChildNodes) {
            if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                Visit-OneNoteNode `
                    -Node $child `
                    -PathParts $nextPathParts `
                    -NotebookName $NotebookName `
                    -SectionGroupPath $groupPath `
                    -ParentContainerId $nodeId
            }
        }
        return
    }

    if ($local -eq "Section") {
        Add-Target `
            -Kind "section" `
            -DisplayName "섹션 - $nodeName" `
            -Path $path `
            -Notebook $NotebookName `
            -SectionGroup $SectionGroupPath `
            -Section $nodeName `
            -SectionGroupId $ParentContainerId `
            -SectionId $nodeId
    }
}

foreach ($notebook in @($doc.SelectNodes("//one:Notebook", $ns))) {
    Visit-OneNoteNode `
        -Node $notebook `
        -PathParts @() `
        -NotebookName "" `
        -SectionGroupPath "" `
        -ParentContainerId ""
}

$result = [ordered]@{
    generated_at = (Get-Date).ToString("s")
    count = $targets.Count
    targets = $targets
}

$result | ConvertTo-Json -Depth 8 -Compress
"""

    def _codex_location_lookup_targets_from_json_text(
        self, text: str
    ) -> List[Dict[str, str]]:
        data = self._codex_json_from_text(text)
        raw_targets = data.get("targets") if isinstance(data, dict) else data
        if not isinstance(raw_targets, list):
            raise ValueError("OneNote 위치 조회 결과에서 위치 목록을 찾지 못했습니다.")

        targets: List[Dict[str, str]] = []
        for item in raw_targets:
            if not isinstance(item, dict):
                continue
            kind = str(item.get("kind", "") or "").strip()
            path = str(item.get("path", "") or "").strip()
            notebook = str(item.get("notebook", "") or "").strip()
            section_group = str(item.get("section_group", "") or "").strip()
            section = str(item.get("section", "") or "").strip()
            section_group_id = str(item.get("section_group_id", "") or "").strip()
            section_id = str(item.get("section_id", "") or "").strip()
            name = str(item.get("name", "") or "").strip()
            if not path:
                continue
            targets.append(
                {
                    "kind": kind,
                    "name": name or path,
                    "path": path,
                    "notebook": notebook,
                    "section_group": section_group,
                    "section": section,
                    "section_group_id": section_group_id,
                    "section_id": section_id,
                }
            )
        return targets

    def _codex_location_lookup_label(self, profile: Dict[str, str]) -> str:
        kind = profile.get("kind", "")
        kind_label = {
            "notebook": "전자필기장",
            "section_group": "그룹",
            "section": "섹션",
        }.get(kind, "위치")
        return f"{kind_label} | {profile.get('path', '')}"

    def _codex_location_cache_path(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "onenote-location-cache.json")

    def _codex_location_targets_from_saved_targets(self) -> List[Dict[str, str]]:
        targets: List[Dict[str, str]] = []
        seen: Set[str] = set()
        for item in self._load_codex_targets():
            if not isinstance(item, dict):
                continue
            path = str(item.get("path", "") or "").strip()
            notebook = str(item.get("notebook", "") or "").strip()
            section_group = str(item.get("section_group", "") or "").strip()
            section = str(item.get("section", "") or "").strip()
            if not path:
                continue
            entries = [
                {
                    "kind": "notebook",
                    "name": notebook,
                    "path": notebook,
                    "notebook": notebook,
                    "section_group": "",
                    "section": "",
                    "section_group_id": "",
                    "section_id": "",
                },
                {
                    "kind": "section_group",
                    "name": section_group,
                    "path": " > ".join(part for part in [notebook, section_group] if part),
                    "notebook": notebook,
                    "section_group": section_group,
                    "section": "",
                    "section_group_id": str(item.get("section_group_id", "") or ""),
                    "section_id": "",
                },
                {
                    "kind": "section",
                    "name": section or path,
                    "path": path,
                    "notebook": notebook,
                    "section_group": section_group,
                    "section": section,
                    "section_group_id": str(item.get("section_group_id", "") or ""),
                    "section_id": str(item.get("section_id", "") or ""),
                },
            ]
            for profile in entries:
                if not profile.get("path"):
                    continue
                key = "|".join(
                    [
                        profile.get("kind", ""),
                        profile.get("path", ""),
                        profile.get("section_id", ""),
                    ]
                )
                if key in seen:
                    continue
                seen.add(key)
                targets.append(profile)
        return targets

    def _load_codex_location_lookup_cache(self) -> List[Dict[str, str]]:
        path = self._codex_location_cache_path()
        try:
            with open(path, "r", encoding="utf-8") as f:
                targets = self._codex_location_lookup_targets_from_json_text(f.read())
            if targets:
                return targets
        except Exception:
            pass
        return self._codex_location_targets_from_saved_targets()

    def _save_codex_location_lookup_cache(self, targets: List[Dict[str, str]]) -> None:
        payload = {
            "version": 1,
            "cached_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "targets": targets,
        }
        self._write_json_file_atomic(self._codex_location_cache_path(), payload)

    def _load_codex_location_lookup_cache_into_ui(self, selected_path: str = "") -> bool:
        targets = self._load_codex_location_lookup_cache()
        if not targets:
            return False
        self._codex_location_lookup_targets = targets
        self._populate_codex_location_lookup_combo(selected_path)
        return True

    def _codex_location_first_profile(
        self, *, kind: str = "", notebook: str = "", section_group: Optional[str] = None
    ) -> Optional[Dict[str, str]]:
        for profile in getattr(self, "_codex_location_lookup_targets", []):
            if not isinstance(profile, dict):
                continue
            if kind and profile.get("kind") != kind:
                continue
            if notebook and profile.get("notebook") != notebook:
                continue
            if section_group is not None and profile.get("section_group", "") != section_group:
                continue
            return profile
        return None

    def _codex_location_selected_notebook(self) -> str:
        combo = getattr(self, "codex_location_notebook_combo", None)
        if combo is None:
            return ""
        return str(combo.currentData() or "").strip()

    def _codex_location_selected_group(self) -> str:
        combo = getattr(self, "codex_location_group_combo", None)
        if combo is None:
            return ""
        data = combo.currentData()
        if isinstance(data, dict):
            return str(data.get("section_group", "") or "").strip()
        return ""

    def _configure_codex_lookup_combo(self, combo: QComboBox) -> None:
        combo.setMinimumWidth(0)
        combo.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        try:
            combo.setSizeAdjustPolicy(
                QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon
            )
            combo.setMinimumContentsLength(8)
            combo.setMaxVisibleItems(16)
        except Exception:
            pass

    def _populate_codex_location_group_combo(
        self, selected_group: str = "", selected_section: str = ""
    ) -> None:
        combo = getattr(self, "codex_location_group_combo", None)
        if combo is None:
            return

        notebook = self._codex_location_selected_notebook()
        seen: Set[str] = set()
        groups: List[Dict[str, str]] = []
        for profile in getattr(self, "_codex_location_lookup_targets", []):
            if not isinstance(profile, dict):
                continue
            if notebook and profile.get("notebook") != notebook:
                continue
            group_name = str(profile.get("section_group", "") or "").strip()
            if not group_name or group_name in seen:
                continue
            if profile.get("kind") == "section_group":
                group_profile = dict(profile)
            else:
                group_profile = {
                    "kind": "section_group",
                    "name": group_name,
                    "path": " > ".join(
                        part for part in [profile.get("notebook", ""), group_name] if part
                    ),
                    "notebook": profile.get("notebook", ""),
                    "section_group": group_name,
                    "section": "",
                    "section_group_id": profile.get("section_group_id", ""),
                    "section_id": "",
                }
            groups.append(group_profile)
            seen.add(group_name)

        combo.blockSignals(True)
        combo.clear()
        combo.addItem("섹션 그룹 없음", {
            "kind": "section_group",
            "name": "섹션 그룹 없음",
            "path": notebook,
            "notebook": notebook,
            "section_group": "",
            "section": "",
            "section_group_id": "",
            "section_id": "",
        })
        selected_idx = 0
        for idx, profile in enumerate(groups, start=1):
            label = profile.get("section_group", "") or profile.get("path", "")
            combo.addItem(label, profile)
            if selected_group and profile.get("section_group") == selected_group:
                selected_idx = idx
        combo.setCurrentIndex(selected_idx)
        combo.blockSignals(False)
        self._populate_codex_location_section_combo(selected_section)

    def _populate_codex_location_section_combo(self, selected_section: str = "") -> None:
        combo = getattr(self, "codex_location_section_combo", None)
        if combo is None:
            return

        notebook = self._codex_location_selected_notebook()
        group_name = self._codex_location_selected_group()
        sections = [
            profile
            for profile in getattr(self, "_codex_location_lookup_targets", [])
            if isinstance(profile, dict)
            and profile.get("kind") == "section"
            and (not notebook or profile.get("notebook") == notebook)
            and profile.get("section_group", "") == group_name
        ]

        combo.blockSignals(True)
        combo.clear()
        combo.addItem("섹션 선택", None)
        selected_idx = 0
        for idx, profile in enumerate(sections, start=1):
            label = profile.get("section", "") or profile.get("path", "")
            combo.addItem(label, profile)
            if selected_section and profile.get("section") == selected_section:
                selected_idx = idx
        combo.setCurrentIndex(selected_idx)
        combo.blockSignals(False)

    def _populate_codex_location_lookup_combo(
        self, selected_path: str = ""
    ) -> None:
        targets = getattr(self, "_codex_location_lookup_targets", [])
        notebook_combo = getattr(self, "codex_location_notebook_combo", None)
        if notebook_combo is None:
            return

        selected_profile = None
        matched_selected_path = False
        if selected_path:
            selected_profile = next(
                (
                    profile
                    for profile in targets
                    if isinstance(profile, dict) and profile.get("path") == selected_path
                ),
                None,
            )
            matched_selected_path = selected_profile is not None
        if selected_profile is None:
            selected_profile = self._codex_location_first_profile(kind="section")
        if selected_profile is None:
            selected_profile = self._codex_location_first_profile()

        notebooks: List[str] = []
        seen: Set[str] = set()
        for profile in targets:
            if not isinstance(profile, dict):
                continue
            notebook = str(profile.get("notebook", "") or "").strip()
            if notebook and notebook not in seen:
                notebooks.append(notebook)
                seen.add(notebook)

        selected_notebook = (selected_profile or {}).get("notebook", "")
        notebook_combo.blockSignals(True)
        notebook_combo.clear()
        for notebook in notebooks:
            notebook_combo.addItem(notebook, notebook)
        selected_idx = max(0, notebook_combo.findData(selected_notebook))
        notebook_combo.setCurrentIndex(selected_idx if notebook_combo.count() else -1)
        notebook_combo.blockSignals(False)

        self._populate_codex_location_group_combo(
            (selected_profile or {}).get("section_group", ""),
            (selected_profile or {}).get("section", ""),
        )

        if selected_profile is not None and matched_selected_path:
            self._apply_codex_location_profile(selected_profile)

    def _set_codex_location_lookup_enabled(self, enabled: bool) -> None:
        toggle = getattr(self, "codex_location_lookup_toggle", None)
        lookup_widgets = (
            getattr(self, "codex_location_notebook_combo", None),
            getattr(self, "codex_location_group_combo", None),
            getattr(self, "codex_location_section_combo", None),
        )
        refresh_btn = getattr(self, "codex_location_lookup_refresh_btn", None)

        if toggle is not None:
            if toggle.isChecked() != enabled:
                toggle.blockSignals(True)
                toggle.setChecked(enabled)
                toggle.blockSignals(False)
            toggle.setText("OneNote 조회 ON" if enabled else "OneNote 조회 OFF")

        for widget in (*lookup_widgets, refresh_btn):
            if widget is not None:
                widget.setVisible(enabled)

        notebook_combo = getattr(self, "codex_location_notebook_combo", None)
        if enabled and notebook_combo is not None and notebook_combo.count() == 0:
            current_path = ""
            path_input = getattr(self, "codex_target_path_input", None)
            if path_input is not None:
                current_path = path_input.text().strip()
            if self._load_codex_location_lookup_cache_into_ui(current_path):
                try:
                    count = len(getattr(self, "_codex_location_lookup_targets", []))
                    self.connection_status_label.setText(
                        f"저장된 OneNote 위치 {count}개를 불러왔습니다."
                    )
                except Exception:
                    pass
            else:
                try:
                    self.connection_status_label.setText(
                        "저장된 OneNote 위치가 없습니다. 조회를 눌러 한 번 갱신하세요."
                    )
                except Exception:
                    pass

    def _refresh_codex_location_lookup(self) -> None:
        worker = getattr(self, "_codex_location_lookup_worker", None)
        try:
            if worker is not None and worker.isRunning():
                self.connection_status_label.setText("OneNote 위치 조회가 이미 진행 중입니다.")
                return
        except Exception:
            pass

        toggle = getattr(self, "codex_location_lookup_toggle", None)
        refresh_btn = getattr(self, "codex_location_lookup_refresh_btn", None)
        lookup_widgets = (
            getattr(self, "codex_location_notebook_combo", None),
            getattr(self, "codex_location_group_combo", None),
            getattr(self, "codex_location_section_combo", None),
        )
        current_path = ""
        path_input = getattr(self, "codex_target_path_input", None)
        if path_input is not None:
            current_path = path_input.text().strip()

        if refresh_btn is not None:
            refresh_btn.setEnabled(False)
            refresh_btn.setText("조회 중")
        if toggle is not None:
            toggle.setEnabled(False)
        try:
            self.connection_status_label.setText(
                "OneNote 위치 조회 중... 저장된 위치 목록은 계속 사용할 수 있습니다."
            )
        except Exception:
            pass

        worker = CodexLocationLookupWorker(
            self._codex_onenote_location_lookup_script(),
            timeout=60,
            parent=self,
        )
        self._codex_location_lookup_worker = worker
        self._retain_qthread_until_finished(worker, "_codex_location_lookup_worker")
        worker.done.connect(
            lambda result, selected_path=current_path, lookup_widgets=lookup_widgets, refresh_btn=refresh_btn, toggle=toggle, worker=worker: self._on_codex_location_lookup_done(
                result,
                selected_path,
                lookup_widgets,
                refresh_btn,
                toggle,
                worker,
            )
        )
        worker.start()

    def _on_codex_location_lookup_done(
        self,
        result: Dict[str, Any],
        selected_path: str,
        lookup_widgets: Tuple[Optional[QWidget], ...],
        refresh_btn: Optional[QToolButton],
        toggle: Optional[QToolButton],
        worker: CodexLocationLookupWorker,
    ) -> None:
        active_worker = getattr(self, "_codex_location_lookup_worker", None)
        if active_worker is not None and active_worker is not worker:
            return

        if refresh_btn is not None:
            refresh_btn.setEnabled(True)
            refresh_btn.setText("조회")
        if toggle is not None:
            toggle.setEnabled(True)
        for widget in lookup_widgets:
            if widget is not None:
                widget.setEnabled(True)

        if not result.get("ok"):
            QMessageBox.warning(
                self,
                "OneNote 위치 조회 실패",
                str(result.get("error") or "알 수 없는 오류"),
            )
            return

        try:
            targets = self._codex_location_lookup_targets_from_json_text(
                str(result.get("raw") or "")
            )
            self._codex_location_lookup_targets = targets
            self._save_codex_location_lookup_cache(targets)
            self._populate_codex_location_lookup_combo(selected_path)
            elapsed = max(0.0, float(result.get("elapsed_ms") or 0) / 1000.0)
            try:
                self.connection_status_label.setText(
                    f"OneNote 위치 조회 완료: {len(targets)}개 저장 ({elapsed:.1f}초)"
                )
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "OneNote 위치 조회 결과 처리 실패", str(e))

    def _apply_codex_location_profile(self, profile: Dict[str, str]) -> None:
        if not isinstance(profile, dict):
            return
        self._populate_codex_target_fields(profile)
        self._apply_codex_target_to_request()
        try:
            self._update_codex_status_summary()
            self.connection_status_label.setText(
                f"코덱스 세부 위치 반영: {profile.get('path', '')}"
            )
        except Exception:
            pass

    def _on_codex_location_notebook_selected(self) -> None:
        self._populate_codex_location_group_combo()
        profile = self._codex_location_first_profile(
            kind="notebook",
            notebook=self._codex_location_selected_notebook(),
        )
        if profile is not None:
            self._apply_codex_location_profile(profile)

    def _on_codex_location_group_selected(self) -> None:
        self._populate_codex_location_section_combo()
        combo = getattr(self, "codex_location_group_combo", None)
        profile = combo.currentData() if combo is not None else None
        if isinstance(profile, dict):
            self._apply_codex_location_profile(profile)

    def _on_codex_location_section_selected(self) -> None:
        combo = getattr(self, "codex_location_section_combo", None)
        profile = combo.currentData() if combo is not None else None
        if isinstance(profile, dict):
            self._apply_codex_location_profile(profile)

    def _codex_target_profile_json_text(self) -> str:
        return json.dumps(
            {
                "version": 1,
                "targets": [self._codex_target_from_fields()],
            },
            ensure_ascii=False,
            indent=2,
        )

    def _copy_codex_target_profile_json_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_current_target_copy_text())
        try:
            self.connection_status_label.setText(
                "현재 작업 위치를 한국어로 복사했습니다."
            )
        except Exception:
            pass

    def _codex_all_targets_json_text(self) -> str:
        return json.dumps(
            {
                "version": 1,
                "exported_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "targets": self._load_codex_targets(),
            },
            ensure_ascii=False,
            indent=2,
        )

    def _copy_codex_all_targets_json_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_all_targets_copy_text())
        try:
            self.connection_status_label.setText("저장된 작업 위치 목록을 한국어로 복사했습니다.")
        except Exception:
            pass

    def _codex_all_targets_copy_text(self) -> str:
        targets = self._load_codex_targets()
        lines = ["저장된 OneNote 작업 위치 목록"]
        if not targets:
            lines.append("- 저장된 작업 위치가 없습니다.")
            return "\n".join(lines)
        for idx, target in enumerate(targets, start=1):
            path = target.get("path") or target.get("name") or "경로 미지정"
            notebook = target.get("notebook") or "미지정"
            section_group = target.get("section_group") or "미지정"
            section = target.get("section") or "미지정"
            lines.extend(
                [
                    f"{idx}. {path}",
                    f"   - 전자필기장: {notebook}",
                    f"   - 섹션 그룹: {section_group}",
                    f"   - 섹션: {section}",
                ]
            )
        return "\n".join(lines)

    def _codex_targets_from_inventory_json_text(self, text: str) -> List[Dict[str, str]]:
        raw = (text or "").strip()
        if not raw:
            raise ValueError("클립보드에 OneNote 위치 조회 결과가 없습니다.")

        try:
            data = json.loads(raw)
        except Exception:
            starts = [idx for idx in (raw.find("{"), raw.find("[")) if idx >= 0]
            ends = [idx for idx in (raw.rfind("}"), raw.rfind("]")) if idx >= 0]
            if not starts or not ends:
                raise ValueError("위치 조회 결과의 시작과 끝을 찾지 못했습니다.")
            data = json.loads(raw[min(starts): max(ends) + 1])

        items = data.get("items") if isinstance(data, dict) else data
        if not isinstance(items, list):
            raise ValueError("위치 조회 결과에서 작업 위치 목록을 찾지 못했습니다.")

        section_group_ids: Dict[str, str] = {}
        for item in items:
            if not isinstance(item, dict):
                continue
            if str(item.get("type", "")).casefold() == "sectiongroup":
                path = str(item.get("path", "")).strip()
                item_id = str(item.get("id", "")).strip()
                if path and item_id:
                    section_group_ids[path] = item_id

        targets: List[Dict[str, str]] = []
        seen: Set[str] = set()
        for item in items:
            if not isinstance(item, dict):
                continue
            if str(item.get("type", "")).casefold() != "section":
                continue
            path = str(item.get("path", "")).strip()
            section_id = str(item.get("id", "")).strip()
            if not path or not section_id:
                continue
            parts = [part.strip() for part in path.split(">") if part.strip()]
            if not parts:
                continue
            notebook = parts[0]
            section = parts[-1]
            section_group = " > ".join(parts[1:-1])
            group_path = " > ".join(parts[:-1])
            dedupe_key = f"{path}|{section_id}"
            if dedupe_key in seen:
                continue
            seen.add(dedupe_key)
            targets.append(
                {
                    "name": f"조회 위치 - {section}",
                    "path": path,
                    "notebook": notebook,
                    "section_group": section_group,
                    "section": section,
                    "section_group_id": section_group_ids.get(group_path, ""),
                    "section_id": section_id,
                }
            )
        return targets

    def _codex_inventory_target_preview_text(self) -> str:
        targets = self._codex_targets_from_inventory_json_text(QApplication.clipboard().text())
        rows = [
            f"| {idx} | {target.get('name', '')} | {target.get('path', '')} | "
            f"{target.get('section_id', '')} |"
            for idx, target in enumerate(targets, start=1)
        ]
        if not rows:
            rows.append("| - | 후보 없음 | - | - |")
        return "\n".join(
            [
                "# OneNote 작업 위치 후보",
                "",
                f"후보 수: {len(targets)}",
                "",
                "| 번호 | 이름 | 경로 | Section ID |",
                "| ---: | --- | --- | --- |",
                *rows,
                "",
            ]
        )

    def _copy_codex_inventory_target_preview_to_clipboard(self) -> None:
        try:
            QApplication.clipboard().setText(self._codex_inventory_target_preview_text())
            try:
                self.connection_status_label.setText(
                    "OneNote 작업 위치 후보를 클립보드에 복사했습니다."
                )
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "작업 위치 후보 생성 실패", str(e))

    def _import_codex_targets_from_clipboard_inventory(self) -> None:
        try:
            incoming = self._codex_targets_from_inventory_json_text(
                QApplication.clipboard().text()
            )
            if not incoming:
                QMessageBox.information(self, "작업 위치 후보 없음", "등록할 섹션 후보를 찾지 못했습니다.")
                return

            targets = self._load_codex_targets()
            existing_keys = {
                (target.get("path", ""), target.get("section_id", ""))
                for target in targets
                if isinstance(target, dict)
            }
            added = 0
            for target in incoming:
                key = (target.get("path", ""), target.get("section_id", ""))
                if key in existing_keys:
                    continue
                targets.append(target)
                existing_keys.add(key)
                added += 1

            self._write_codex_targets(targets)
            self._refresh_codex_target_combo()
            self._update_codex_status_summary()
            QMessageBox.information(
                self,
                "작업 위치 등록 완료",
                f"새 작업 위치 {added}개를 등록했습니다.\n전체 후보: {len(incoming)}개",
            )
        except Exception as e:
            QMessageBox.warning(self, "작업 위치 등록 실패", str(e))

    def _codex_onenote_inventory_script(self) -> str:
        return """# OneNote COM: 전체 구조 위치 조회 JSON
# 전자필기장/섹션그룹/섹션/페이지 목록을 평탄화해서 클립보드에 JSON으로 복사합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$xml = ""
$one.GetHierarchy(
    "",
    [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages,
    [ref]$xml
)

[xml]$doc = $xml
$nsUri = $doc.DocumentElement.NamespaceURI
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("one", $nsUri)

$items = New-Object System.Collections.Generic.List[object]

function Add-OneNoteNode {
    param(
        [System.Xml.XmlNode]$Node,
        [string[]]$PathParts
    )

    $local = $Node.LocalName
    if ($local -notin @("Notebook", "SectionGroup", "Section", "Page")) {
        return
    }

    $name = $Node.GetAttribute("name")
    if ([string]::IsNullOrWhiteSpace($name) -and $local -eq "Page") {
        $name = $Node.GetAttribute("name")
    }
    if ([string]::IsNullOrWhiteSpace($name)) {
        $name = "(이름 없음)"
    }

    $nextPath = @($PathParts + $name)
    $items.Add([ordered]@{
        type = $local
        name = $name
        id = $Node.GetAttribute("ID")
        path = [string]::Join(" > ", [string[]]$nextPath)
        isUnread = $Node.GetAttribute("isUnread")
        lastModifiedTime = $Node.GetAttribute("lastModifiedTime")
    })

    foreach ($child in $Node.ChildNodes) {
        if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            Add-OneNoteNode -Node $child -PathParts $nextPath
        }
    }
}

foreach ($notebook in @($doc.SelectNodes("//one:Notebook", $ns))) {
    Add-OneNoteNode -Node $notebook -PathParts @()
}

$summary = [ordered]@{
    generated_at = (Get-Date).ToString("s")
    total = $items.Count
    notebooks = @($items | Where-Object { $_.type -eq "Notebook" }).Count
    section_groups = @($items | Where-Object { $_.type -eq "SectionGroup" }).Count
    sections = @($items | Where-Object { $_.type -eq "Section" }).Count
    pages = @($items | Where-Object { $_.type -eq "Page" }).Count
    items = $items
}

$json = $summary | ConvertTo-Json -Depth 8
$json | Set-Clipboard
$json
"""

    def _copy_codex_onenote_inventory_script_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_onenote_location_request_text())
        try:
            self.connection_status_label.setText(
                "OneNote 위치 조회 요청을 한국어로 복사했습니다."
            )
        except Exception:
            pass

    def _codex_onenote_location_request_text(self) -> str:
        return """작업:
OneNote에서 현재 열려 있는 전자필기장, 섹션 그룹, 섹션을 조회해줘.

정리 방식:
- 전자필기장별로 섹션 그룹과 섹션을 계층형으로 정리한다.
- 작업 위치로 쓸 수 있는 섹션 후보를 따로 표시한다.
- 사용자가 복사해서 작업 지시로 쓸 수 있게 경로를 `전자필기장 > 섹션 그룹 > 섹션` 형식으로 적는다.

보고:
- 조회한 전자필기장 수
- 섹션 그룹 수
- 섹션 수
- 바로 작업 위치로 지정할 만한 후보 목록
"""

    def _codex_page_reader_script(self) -> str:
        profile = self._codex_target_from_fields()
        section_id = self._ps_single_quoted(profile.get("section_id", ""))
        section_path = self._ps_single_quoted(profile.get("path", ""))
        return f"""# OneNote COM: 현재 대상 섹션 페이지 읽기
# 현재 코덱스 작업 위치의 Section ID 기준으로 페이지 목록과 선택 페이지 내용을 조회합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$sectionId = {section_id}
$sectionPath = {section_path}
$PageTitleContains = ""

if ([string]::IsNullOrWhiteSpace($sectionId)) {{
    throw "Section ID가 비어 있습니다. OneNote 조회 ON으로 세부 위치를 선택하거나 왼쪽 패널에서 섹션을 선택하세요. 대상: $sectionPath"
}}

$hierarchyXml = ""
$one.GetHierarchy(
    $sectionId,
    [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages,
    [ref]$hierarchyXml
)

[xml]$doc = $hierarchyXml
$nsUri = $doc.DocumentElement.NamespaceURI
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("one", $nsUri)

$pages = @($doc.SelectNodes("//one:Page", $ns)) | ForEach-Object {{
    [ordered]@{{
        title = $_.GetAttribute("name")
        id = $_.GetAttribute("ID")
        lastModifiedTime = $_.GetAttribute("lastModifiedTime")
    }}
}}

$selectedPage = $null
if (-not [string]::IsNullOrWhiteSpace($PageTitleContains)) {{
    $selectedPage = $pages |
        Where-Object {{ $_.title -like "*$PageTitleContains*" }} |
        Select-Object -First 1
}}

$pageContent = ""
if ($null -ne $selectedPage) {{
    $one.GetPageContent($selectedPage.id, [ref]$pageContent)
}}

$result = [ordered]@{{
    generated_at = (Get-Date).ToString("s")
    section_id = $sectionId
    section_path = $sectionPath
    page_count = @($pages).Count
    pages = $pages
    selected_page = $selectedPage
    selected_page_xml = $pageContent
}}

$json = $result | ConvertTo-Json -Depth 8
$json | Set-Clipboard
$json
"""

    def _copy_codex_page_reader_script_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_page_reader_request_text())
        try:
            self.connection_status_label.setText(
                "현재 섹션의 페이지 읽기 요청을 한국어로 복사했습니다."
            )
        except Exception:
            pass

    def _codex_page_reader_request_text(self) -> str:
        try:
            target = self._codex_target_from_fields()
        except Exception:
            target = {}
        path = target.get("path") or "현재 선택된 작업 위치"
        return f"""작업:
아래 OneNote 작업 위치의 페이지 목록을 읽어줘.

작업 위치:
{path}

정리 방식:
- 페이지 제목 목록을 먼저 적는다.
- 최근 수정된 페이지가 있으면 표시한다.
- 사용자가 이어서 작업할 만한 페이지 후보를 알려준다.
- 특정 페이지 내용을 읽어야 하면 어떤 제목을 골라야 하는지 묻는다.

보고:
- 페이지 수
- 페이지 제목 목록
- 다음에 읽을 만한 페이지 후보
"""

    def _codex_json_from_text(self, text: str) -> Any:
        raw = (text or "").strip()
        if not raw:
            raise ValueError("조회 결과가 비어 있습니다.")
        try:
            return json.loads(raw)
        except Exception:
            starts = [idx for idx in (raw.find("{"), raw.find("[")) if idx >= 0]
            ends = [idx for idx in (raw.rfind("}"), raw.rfind("]")) if idx >= 0]
            if not starts or not ends:
                raise
            return json.loads(raw[min(starts): max(ends) + 1])

    def _codex_text_from_page_xml(self, xml_text: str) -> str:
        if not xml_text:
            return ""
        chunks = re.findall(r"(?is)<[^:>]*:?T[^>]*>(.*?)</[^:>]*:?T>", xml_text)
        if not chunks:
            chunks = [re.sub(r"(?is)<[^>]+>", " ", xml_text)]
        lines = []
        for chunk in chunks:
            plain = re.sub(r"(?is)<[^>]+>", " ", chunk)
            plain = html.unescape(plain)
            plain = re.sub(r"\s+", " ", plain).strip()
            if plain:
                lines.append(plain)
        return "\n".join(lines)

    def _codex_page_reader_result_summary_text(self, text: str = "") -> str:
        data = self._codex_json_from_text(text or QApplication.clipboard().text())
        if not isinstance(data, dict):
            raise ValueError("페이지 읽기 결과 형식이 올바르지 않습니다.")

        pages = data.get("pages", [])
        if not isinstance(pages, list):
            pages = []

        rows = []
        for idx, page in enumerate(pages, start=1):
            if not isinstance(page, dict):
                continue
            rows.append(
                f"| {idx} | {page.get('title', '') or '-'} | "
                f"{page.get('lastModifiedTime', '') or '-'} | `{page.get('id', '')}` |"
            )
        if not rows:
            rows.append("| - | 페이지 없음 | - | - |")

        selected = data.get("selected_page")
        selected_title = ""
        if isinstance(selected, dict):
            selected_title = str(selected.get("title", "") or "")
        selected_text = self._codex_text_from_page_xml(str(data.get("selected_page_xml", "") or ""))

        return "\n".join(
            [
                "# OneNote 페이지 읽기 결과 요약",
                "",
                f"생성 시각: {data.get('generated_at', '') or '-'}",
                f"대상 섹션: {data.get('section_path', '') or data.get('section_id', '') or '-'}",
                f"페이지 수: {data.get('page_count', len(pages))}",
                f"선택 페이지: {selected_title or '-'}",
                "",
                "## 페이지 목록",
                "",
                "| 번호 | 제목 | 수정 시각 | Page ID |",
                "| ---: | --- | --- | --- |",
                *rows,
                "",
                "## 선택 페이지 텍스트",
                "",
                selected_text or "- 선택된 페이지 XML이 없거나 텍스트를 추출하지 못했습니다.",
                "",
            ]
        )

    def _copy_codex_page_reader_result_summary_to_clipboard(self) -> None:
        try:
            QApplication.clipboard().setText(self._codex_page_reader_result_summary_text())
            try:
                self.connection_status_label.setText("OneNote 페이지 읽기 결과 요약을 클립보드에 복사했습니다.")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "페이지 읽기 결과 요약 실패", str(e))

    def _append_codex_page_reader_result_to_request_body(self) -> None:
        body_editor = getattr(self, "codex_request_body_editor", None)
        if body_editor is None:
            return
        try:
            summary = self._codex_page_reader_result_summary_text()
        except Exception as e:
            QMessageBox.warning(self, "페이지 읽기 결과 추가 실패", str(e))
            return
        current = body_editor.toPlainText().rstrip()
        body_editor.setPlainText((current + "\n\n" + summary).strip())
        self._update_codex_codegen_previews()
        try:
            self.connection_status_label.setText("페이지 읽기 결과 요약을 요청 본문에 추가했습니다.")
        except Exception:
            pass

    def _codex_onenote_templates(self) -> Dict[str, str]:
        values = self._codex_codegen_values()
        title = self._ps_single_quoted(values["title"])
        body = self._ps_single_quoted(values["body"])
        target = self._ps_single_quoted(values["target"])
        section_group_id = self._ps_single_quoted(values["section_group_id"])
        section_id = self._ps_single_quoted(values["section_id"])
        return {
            "add_page": f"""# OneNote COM: 페이지 추가
# 대상: {values["target"]}
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$sectionId = {section_id}
$title = {title}
$body = {body}
$pageId = ""
$one.CreateNewPage(
    $sectionId,
    [ref]$pageId,
    [Microsoft.Office.Interop.OneNote.NewPageStyle]::npsBlankPageWithTitle
)

$pageXml = ""
$one.GetPageContent($pageId, [ref]$pageXml)

[xml]$doc = $pageXml
$nsUri = $doc.DocumentElement.NamespaceURI
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("one", $nsUri)

$titleNode = $doc.SelectSingleNode("//one:Title//one:T", $ns)
if ($titleNode -ne $null) {{
    $titleNode.InnerText = $title
}}

$outline = $doc.CreateElement("one", "Outline", $nsUri)
$oeChildren = $doc.CreateElement("one", "OEChildren", $nsUri)
foreach ($line in ($body -split "`r?`n")) {{
    $oe = $doc.CreateElement("one", "OE", $nsUri)
    $t = $doc.CreateElement("one", "T", $nsUri)
    $t.InnerText = $line
    [void]$oe.AppendChild($t)
    [void]$oeChildren.AppendChild($oe)
}}
[void]$outline.AppendChild($oeChildren)
[void]$doc.DocumentElement.AppendChild($outline)

$one.UpdatePageContent($doc.OuterXml)
$verifyXml = ""
$one.GetPageContent($pageId, [ref]$verifyXml)
if ($verifyXml -notlike "*$title*") {{
    throw "페이지 생성 검증 실패: 제목을 찾지 못했습니다."
}}
$one.NavigateTo($pageId)
""",
            "add_section": f"""# OneNote COM: 새 섹션 생성
# 대상: {values["target"]}
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$sectionGroupId = {section_group_id}
$sectionName = {title}
$newSectionId = ""
$safeName = $sectionName -replace '[<>:"/\\\\|?*]', '-'
if (-not $safeName.EndsWith(".one")) {{
    $safeName = "$safeName.one"
}}

$one.OpenHierarchy(
    $safeName,
    $sectionGroupId,
    [ref]$newSectionId,
    [Microsoft.Office.Interop.OneNote.CreateFileType]::cftSection
)

$verifyXml = ""
$one.GetHierarchy($sectionGroupId, [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections, [ref]$verifyXml)
if ($verifyXml -notlike "*$sectionName*") {{
    throw "섹션 생성 검증 실패: 새 섹션 이름을 찾지 못했습니다."
}}
$one.NavigateTo($newSectionId)
""",
            "add_section_group": f"""# OneNote COM: 새 섹션 그룹 생성
# 대상: {values["target"]}
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$parentSectionGroupId = {section_group_id}
$sectionGroupName = {title}
$newSectionGroupId = ""

$one.OpenHierarchy(
    $sectionGroupName,
    $parentSectionGroupId,
    [ref]$newSectionGroupId,
    [Microsoft.Office.Interop.OneNote.CreateFileType]::cftFolder
)

$verifyXml = ""
$one.GetHierarchy($parentSectionGroupId, [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections, [ref]$verifyXml)
if ($verifyXml -notlike "*$sectionGroupName*") {{
    throw "섹션 그룹 생성 검증 실패: 새 섹션 그룹 이름을 찾지 못했습니다."
}}
$one.NavigateTo($newSectionGroupId)
""",
            "add_notebook": f"""# OneNote COM: 새 전자필기장 생성
# 주의: 새 전자필기장은 저장 위치 경로가 필요하다. OneDrive 동기화 폴더 또는 로컬 경로를 먼저 확정한다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$notebookPath = {target}
$newNotebookId = ""
if ([string]::IsNullOrWhiteSpace($notebookPath)) {{
    throw "새 전자필기장을 만들 파일 시스템 경로가 필요합니다."
}}

$one.OpenHierarchy(
    $notebookPath,
    "",
    [ref]$newNotebookId,
    [Microsoft.Office.Interop.OneNote.CreateFileType]::cftNotebook
)

$one.NavigateTo($newNotebookId)
""",
        }

    def _selected_codex_template_text(self) -> str:
        combo = getattr(self, "codex_template_combo", None)
        if combo is None:
            return ""
        key = combo.currentData()
        return self._codex_onenote_templates().get(key, "")

    def _update_codex_template_preview(self) -> None:
        preview = getattr(self, "codex_template_preview", None)
        if preview is None:
            return
        self._set_plain_text_if_changed(preview, self._selected_codex_template_text())

    def _set_plain_text_if_changed(self, widget: Optional[QTextEdit], text: str) -> None:
        if widget is None:
            return
        if widget.toPlainText() != text:
            widget.setPlainText(text)

    def _set_label_text_if_changed(self, widget: Optional[QLabel], text: str) -> None:
        if widget is None:
            return
        if widget.text() != text:
            widget.setText(text)

    def _schedule_codex_codegen_previews(self, *args) -> None:
        timer = getattr(self, "_codex_codegen_preview_timer", None)
        if timer is None:
            timer = QTimer(self)
            timer.setSingleShot(True)
            timer.setInterval(120)
            timer.timeout.connect(self._update_codex_codegen_previews)
            self._codex_codegen_preview_timer = timer
        timer.start()

    def _schedule_codex_skill_call_preview(self, *args) -> None:
        timer = getattr(self, "_codex_skill_preview_timer", None)
        if timer is None:
            timer = QTimer(self)
            timer.setSingleShot(True)
            timer.setInterval(120)
            timer.timeout.connect(self._update_codex_skill_call_preview)
            self._codex_skill_preview_timer = timer
        timer.start()

    def _update_codex_codegen_previews(self) -> None:
        self._update_codex_request_preview()
        self._update_codex_template_preview()
        self._update_codex_work_order_preview()
        self._update_codex_skill_package_preview()
        self._update_codex_status_summary()

    def _copy_codex_template_to_clipboard(self) -> None:
        text = self._selected_codex_template_text()
        if not text:
            return
        QApplication.clipboard().setText(text)
        try:
            self.connection_status_label.setText("코덱스 OneNote 작업 양식을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _codex_skills_dir(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "skills")

    def _codex_skill_packages_dir(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "skill-packages")

    def _codex_instructions_dir(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "instructions")

    def _codex_internal_instructions_path(self) -> str:
        return os.path.join(self._codex_instructions_dir(), "onenote-com-internal.md")

    def _codex_internal_reference_text(self) -> str:
        return (
            "OneNote 조작 방식, 대상 ID, PowerShell 패턴, 검증 기준은 "
            "`docs/codex/instructions/`와 `docs/codex/onenote-targets.json`에서 "
            "필요할 때 조회한다. 사용자 요청문이나 작업 주문서에는 내부 지침 전문을 붙이지 않는다."
        )

    def _codex_requests_dir(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "requests")

    def _codex_request_draft_path(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "request-draft.json")

    def _codex_skill_slug(self, name: str) -> str:
        raw = (name or "").strip()
        raw = re.sub(r'[<>:"/\\\\|?*]+', "-", raw)
        raw = re.sub(r"\s+", "-", raw)
        raw = raw.strip(".-")
        return raw or "codex-skill"

    def _codex_skill_order_index_path(self) -> str:
        return os.path.join(self._codex_skills_dir(), "skill-order-index.md")

    def _write_text_file_atomic(self, path: str, text: str) -> bool:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        try:
            with open(path, "r", encoding="utf-8") as f:
                if f.read() == text:
                    return False
        except FileNotFoundError:
            pass
        except Exception:
            pass

        tmp_path = f"{path}.{uuid.uuid4().hex}.tmp"
        try:
            with open(tmp_path, "w", encoding="utf-8", newline="") as f:
                f.write(text)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp_path, path)
        finally:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass
        return True

    def _write_json_file_atomic(self, path: str, payload: Any) -> bool:
        text = json.dumps(payload, ensure_ascii=False, indent=2)
        return self._write_text_file_atomic(path, text)

    def _codex_builtin_internal_instructions_text(self) -> str:
        return """# 코덱스 전용 OneNote 조작 지침

이 문서는 사용자 스킬이 아니다. OneNote 작업을 수행하는 Codex가 항상 전제로 삼는 내부 실행 지침이다.

## 적용 순서

1. 사용자 요청과 선택된 사용자 스킬에서 목표, 대상, 작성 형식만 추출한다.
2. OneNote 조작 방식은 이 폴더의 작업별 내부 지침에서 고른다.
3. 저장된 대상 ID가 있으면 전체 탐색보다 ID 직접 호출을 우선한다.
4. 완료 후 화면 캡처가 아니라 `GetHierarchy` 또는 `GetPageContent` 결과로 검증한다.
5. 최종 보고에는 변경 항목과 검증 결과만 짧게 남긴다.

## 사용자 스킬과의 경계

- 사용자 스킬에서는 `## Instructions`만 작업에 맞게 적용한다.
- 사용자 요청문, 작업 주문서, 스킬 호출문에는 이 문서 전문이나 PowerShell 템플릿을 붙이지 않는다.
- 글쓰기 형식, 정리 방식, 이름 규칙처럼 요청마다 달라지는 기준은 사용자 스킬에 둔다.
- COM 호출 순서, 대상 ID 우선순위, XML 수정 방식, 검증 절차는 코덱스 전용 지침에 둔다.
- 사용자가 내부 구현 설명을 요청하지 않았다면 COM 세부 호출을 길게 설명하지 않는다.

## 공통 실행 원칙

- OneNote 조작은 화면 클릭 자동화보다 Windows OneNote COM API를 우선 사용한다.
- PowerShell에서는 `New-Object -ComObject OneNote.Application`으로 연결한다.
- 구조 탐색은 기본적으로 `GetHierarchy('', hsSections, ref xml)`까지만 사용한다.
- `hsPages` 전체 조회는 페이지가 많으면 느리므로 페이지 복제, 페이지 목록 조회, 최종 페이지 수 검증처럼 필요한 경우에만 쓴다.
- `docs/codex/onenote-targets.json`과 `docs/codex/onenote-location-cache.json`에 대상 ID가 있으면 먼저 사용한다.
- 이름이 중복될 수 있으므로 전자필기장, 섹션 그룹, 섹션을 상위 경로까지 함께 제한해 찾는다.
- 새 ID를 찾거나 자주 쓸 대상이 생기면 대상 캐시에 저장한다.
- XML은 문자열 치환보다 XML 파서로 수정한다.
- `OpenHierarchy` 또는 `CreateNewPage` 직후에는 필요하면 잠시 대기하고 `GetHierarchy` 또는 `GetPageContent`로 새 ID가 실제 사용 가능한지 확인한다.

## 작업별 내부 문서

- 페이지 추가: `원노트-페이지-추가.md`
- 섹션 생성: `원노트-섹션-생성.md`
- 섹션 그룹 생성: `원노트-섹션그룹-생성.md`
- 전자필기장 생성/열기: `원노트-전자필기장-생성.md`
- 전자필기장 복제: `원노트-전자필기장-복제.md`
- 대상 ID 찾기: `원노트-대상ID-찾기.md`
- 작업 템플릿 기준: `onenote-com-templates.md`

## 빠른 라우팅

- 페이지 추가는 대상 `Section ID`로 `CreateNewPage`, `GetPageContent`, `UpdatePageContent`를 사용한다.
- 섹션 생성은 대상 `SectionGroup ID`로 `OpenHierarchy("섹션명.one", sectionGroupId, ref newId, cftSection)`를 사용한다.
- 섹션 그룹 생성은 전자필기장 또는 섹션 그룹 ID로 `OpenHierarchy("그룹명", parentId, ref newId, cftFolder)`를 사용한다.
- 새 전자필기장은 OneNote가 열 수 있는 로컬/동기화 경로를 확인한 뒤 `OpenHierarchy(notebookPath, "", ref newId, cftNotebook)`를 사용한다.
- 전자필기장 복제는 대상 이름을 `코덱스-{원본 전자필기장명}`으로 잡고, 활성 섹션 그룹/섹션/페이지 수가 원본과 같은지 확인한다.

## 검증 기준

- 페이지 추가/수정: `GetPageContent(pageId)`에서 제목과 본문 일부를 확인한다.
- 섹션 생성: `GetHierarchy(sectionGroupId, hsSections, ref xml)`에서 새 섹션 이름을 확인한다.
- 섹션 그룹 생성: `GetHierarchy(parentId, hsSections, ref xml)`에서 새 그룹 이름을 확인한다.
- 전자필기장 생성/열기: `GetHierarchy('', hsNotebooks, ref xml)` 또는 `hsSections`에서 열린 전자필기장을 확인한다.
- 전자필기장 복제: 내부 휴지통을 제외한 활성 섹션 그룹 수, 섹션 수, 페이지 수가 원본과 대상에서 일치하는지 확인한다.

## 보고 기준

- 성공하면 만든/수정한 OneNote 항목과 검증 결과만 간단히 보고한다.
- 실패하면 대상 경로, 대상 ID, 실패한 단계, 검증 결과를 짧게 보고한다.
- 사용자가 화면 확인을 요청한 경우에만 마지막에 `NavigateTo(...)`를 호출했다고 언급한다.
"""

    def _codex_internal_instructions_text(self) -> str:
        path = self._codex_internal_instructions_path()
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    text = f.read().strip()
                if text:
                    return text
        except Exception:
            pass
        return self._codex_builtin_internal_instructions_text()

    def _ensure_codex_internal_instructions_file(self) -> str:
        path = self._codex_internal_instructions_path()
        if not os.path.exists(path):
            self._write_text_file_atomic(
                path,
                self._codex_builtin_internal_instructions_text(),
            )
        return path

    def _save_codex_internal_instructions(self) -> None:
        editor = getattr(self, "codex_internal_instructions_editor", None)
        if editor is None:
            return
        try:
            path = self._codex_internal_instructions_path()
            text = editor.toPlainText().strip()
            if not text:
                text = self._codex_builtin_internal_instructions_text()
                editor.setPlainText(text)
            self._write_text_file_atomic(path, text + "\n")
            self._update_codex_codegen_previews()
            try:
                self.connection_status_label.setText(f"코덱스 전용 지침 저장 완료: {path}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "코덱스 전용 지침 저장 실패", str(e))

    def _reload_codex_internal_instructions(self) -> None:
        editor = getattr(self, "codex_internal_instructions_editor", None)
        if editor is None:
            return
        try:
            self._ensure_codex_internal_instructions_file()
            editor.setPlainText(self._codex_internal_instructions_text())
            self._update_codex_codegen_previews()
            try:
                self.connection_status_label.setText("코덱스 전용 지침을 다시 불러왔습니다.")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "코덱스 전용 지침 불러오기 실패", str(e))

    def _copy_codex_internal_instructions_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_internal_instructions_text())
        try:
            self.connection_status_label.setText("코덱스 전용 지침을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _open_codex_instructions_folder(self) -> None:
        try:
            os.makedirs(self._codex_instructions_dir(), exist_ok=True)
            os.startfile(self._codex_instructions_dir())
        except Exception as e:
            QMessageBox.warning(self, "코덱스 전용 지침 폴더 열기 실패", str(e))

    def _codex_skill_templates(self) -> Dict[str, Dict[str, str]]:
        return {
            "writing": {
                "name": "나의 기본 글쓰기 형식",
                "trigger": "사용자가 정리된 글쓰기 형식으로 OneNote 페이지 작성을 요청할 때",
                "body": """목표:
- 사용자의 거친 메모를 읽기 쉬운 글로 정리한다.

형식:
- 제목: 핵심 주제를 짧게 쓴다.
- 한 줄 요약: 결론을 먼저 쓴다.
- 본문: 배경, 핵심 내용, 다음 행동 순서로 정리한다.
- 체크포인트: 다시 확인할 항목을 bullet로 남긴다.

작성 기준:
- 사용자의 원래 표현을 최대한 보존하되 문장 흐름만 정리한다.
- 불확실한 내용은 단정하지 말고 확인 필요로 표시한다.
- OneNote에는 제목과 본문을 분리해서 작성한다.

검증:
- 제목과 본문이 위 형식을 충족하는지 확인한다.
""",
            },
            "daily_log": {
                "name": "업무일지 정리",
                "trigger": "오늘 작업 내용, 업무 기록, 진행 상황을 OneNote에 남길 때",
                "body": """목표:
- 하루 동안 한 작업과 다음 행동을 빠르게 되돌아볼 수 있게 정리한다.

형식:
- 날짜:
- 오늘 한 일:
- 결정한 것:
- 막힌 것:
- 다음 행동:

작성 기준:
- 작업 결과와 미완료 항목을 분리한다.
- 다음 행동은 바로 실행 가능한 동사형으로 쓴다.
- 중요한 파일/프로젝트/스킬 주문번호가 있으면 함께 적는다.

검증:
- 페이지 제목에 날짜 또는 작업명이 들어갔는지 확인한다.
""",
            },
            "meeting": {
                "name": "회의록 정리",
                "trigger": "회의 내용, 통화 내용, 논의 결과를 OneNote에 정리할 때",
                "body": """목표:
- 회의에서 결정된 내용과 후속 작업을 빠르게 찾을 수 있게 정리한다.

형식:
- 회의명:
- 참석자:
- 논의 요약:
- 결정 사항:
- 액션 아이템:
- 보류/리스크:

작성 기준:
- 결정 사항과 의견을 섞지 않는다.
- 액션 아이템에는 담당자, 마감, 확인 방법을 포함한다.
- 근거가 부족한 내용은 보류/리스크로 분리한다.

검증:
- 액션 아이템이 별도 목록으로 남아 있는지 확인한다.
""",
            },
            "idea": {
                "name": "아이디어 정리",
                "trigger": "아이디어, 기획, 개선안을 OneNote에서 발전시킬 때",
                "body": """목표:
- 떠오른 아이디어를 실행 가능한 형태로 바꾼다.

형식:
- 아이디어:
- 왜 필요한가:
- 사용자/상황:
- 가능한 구현:
- 예상 문제:
- 다음 실험:

작성 기준:
- 막연한 표현을 구체적인 실험이나 작업 단위로 바꾼다.
- 구현 아이디어와 검증 아이디어를 분리한다.
- 당장 할 수 없는 것은 후보로만 남긴다.

검증:
- 다음 실험이 1개 이상 있는지 확인한다.
""",
            },
            "study": {
                "name": "학습 노트 정리",
                "trigger": "강의, 문서, 공부 내용을 OneNote 학습 노트로 만들 때",
                "body": """목표:
- 학습 내용을 다시 보기 쉬운 구조로 압축한다.

형식:
- 주제:
- 핵심 개념:
- 예시:
- 헷갈린 점:
- 적용할 곳:
- 복습 질문:

작성 기준:
- 정의, 예시, 적용을 분리한다.
- 이해가 불확실한 부분은 헷갈린 점에 남긴다.
- 복습 질문은 나중에 바로 테스트할 수 있게 작성한다.

검증:
- 핵심 개념과 복습 질문이 둘 다 있는지 확인한다.
""",
            },
            "checklist": {
                "name": "작업 체크리스트",
                "trigger": "반복 작업, 배포 전 점검, 정리 절차를 체크리스트로 만들 때",
                "body": """목표:
- 반복 작업을 빠뜨리지 않도록 체크리스트로 만든다.

형식:
- 목적:
- 사전 확인:
- 실행 순서:
- 검증:
- 실패 시 복구:

작성 기준:
- 각 항목은 체크 가능한 문장으로 쓴다.
- 실행 순서와 검증 순서를 분리한다.
- 위험한 작업은 복구 방법을 같이 남긴다.

검증:
- 실행 전/후 확인 항목이 모두 있는지 확인한다.
""",
            },
            "onenote_cleanup": {
                "name": "OneNote 정리 작업",
                "trigger": "OneNote의 메모, 섹션, 섹션 그룹을 정리하거나 재배치할 때",
                "body": """목표:
- OneNote 구조를 안전하게 정리하고 변경 내역을 검증한다.

형식:
- 현재 위치:
- 바꿀 위치:
- 변경 작업:
- 보존할 항목:
- 검증 방법:

작업 기준:
- 삭제보다 이동/이름 변경을 우선한다.
- 전자필기장/섹션/페이지 계층을 먼저 조회한 뒤 작업한다.
- 보존할 항목과 바꿀 위치를 작업 전에 분리해서 적는다.

검증:
- 작업 전후 계층 구조를 비교한다.
""",
            },
        }

    def _apply_codex_skill_template(self) -> None:
        combo = getattr(self, "codex_skill_template_combo", None)
        if combo is None:
            return
        key = combo.currentData()
        template = self._codex_skill_templates().get(key)
        if not template:
            return

        self.codex_skill_order_input.setText(self._codex_next_skill_order())
        self.codex_skill_name_input.setText(template.get("name", "새 코덱스 스킬"))
        self.codex_skill_trigger_input.setText(template.get("trigger", ""))
        self.codex_skill_body_editor.setPlainText(template.get("body", ""))
        self._update_codex_skill_call_preview()
        try:
            self.connection_status_label.setText(
                f"스킬 양식 적용: {template.get('name', '')}"
            )
        except Exception:
            pass

    def _codex_skill_section(self, text: str, heading: str) -> str:
        pattern = rf"(?ms)^##\s+{re.escape(heading)}\s*$\n(.*?)(?=^##\s+|\Z)"
        match = re.search(pattern, text or "")
        return match.group(1).strip() if match else ""

    def _codex_skill_metadata_from_text(
        self, text: str, fallback_name: str = ""
    ) -> Dict[str, str]:
        lines = (text or "").splitlines()
        name = fallback_name
        for line in lines:
            if line.strip().startswith("# "):
                name = line.strip()[2:].strip()
                break

        order_no = self._codex_skill_section(text, "주문번호")
        if not order_no:
            match = re.search(r"(?im)^\s*(?:주문번호|Order)\s*[:：]\s*(.+?)\s*$", text or "")
            order_no = match.group(1).strip() if match else ""

        trigger = self._codex_skill_section(text, "Trigger")
        body = self._codex_skill_section(text, "Instructions")
        if not body:
            body = text or ""

        return {
            "name": name or "새 코덱스 스킬",
            "order": order_no,
            "trigger": trigger,
            "body": body,
        }

    def _codex_skill_metadata_from_file(
        self, path: str, fallback_name: str = ""
    ) -> Dict[str, str]:
        stat = os.stat(path)
        signature = (stat.st_mtime_ns, stat.st_size)
        cache = getattr(self, "_codex_skill_metadata_cache", {})
        cached = cache.get(path)
        if cached and cached[0] == signature:
            return dict(cached[1])

        with open(path, "r", encoding="utf-8") as f:
            meta = self._codex_skill_metadata_from_text(f.read(), fallback_name)

        cache[path] = (signature, dict(meta))
        if len(cache) > 256:
            for old_path in list(cache)[: len(cache) - 256]:
                cache.pop(old_path, None)
        self._codex_skill_metadata_cache = cache
        return dict(meta)

    def _codex_next_skill_order(self) -> str:
        used: Set[int] = set()
        skills_dir = self._codex_skills_dir()
        try:
            records = getattr(self, "_codex_skill_records", [])
            records_mtime = getattr(self, "_codex_skill_records_dir_mtime", None)
            if records and records_mtime == os.path.getmtime(skills_dir):
                for record in records:
                    match = re.search(r"(\d+)", record.get("order", ""))
                    if match:
                        used.add(int(match.group(1)))
                n = 1
                while n in used:
                    n += 1
                return f"SK-{n:03d}"
        except Exception:
            used.clear()

        try:
            for filename in os.listdir(skills_dir):
                if not filename.lower().endswith(".md"):
                    continue
                if filename in ("README.md", "skill-order-index.md", "skill-audit.md"):
                    continue
                path = os.path.join(skills_dir, filename)
                meta = self._codex_skill_metadata_from_file(path, filename[:-3])
                match = re.search(r"(\d+)", meta.get("order", ""))
                if match:
                    used.add(int(match.group(1)))
        except Exception:
            pass

        n = 1
        while n in used:
            n += 1
        return f"SK-{n:03d}"

    def _write_codex_skill_order_index(self, skills: List[Dict[str, str]]) -> None:
        path = self._codex_skill_order_index_path()
        rows = []
        for skill in sorted(skills, key=lambda s: (s.get("order") or "ZZZ", s.get("name") or "")):
            order_no = skill.get("order") or "미지정"
            name = skill.get("name") or "이름 없음"
            filename = skill.get("filename") or ""
            trigger = (skill.get("trigger") or "").replace("\n", " ").strip()
            rows.append(f"| {order_no} | {name} | `{filename}` | {trigger} |")

        text = "\n".join(
            [
                "# 사용자 스킬 주문번호표",
                "",
                "사용자가 주문번호로 스킬을 지시하면 이 표에서 해당 Markdown 파일을 찾아 따른다. OneNote 조작 방식은 `docs/codex/instructions`의 코덱스 전용 지침에서 관리한다.",
                "",
                "| 주문번호 | 스킬 이름 | 파일 | 호출 조건 |",
                "| --- | --- | --- | --- |",
                *rows,
                "",
            ]
        )
        self._write_text_file_atomic(path, text)

    def _codex_skill_search_key(self, *parts: Any) -> str:
        text = " ".join("" if p is None else str(p) for p in parts)
        return re.sub(r"\s+", "", unicodedata.normalize("NFKC", text)).casefold()

    def _populate_codex_skill_list(self, records: List[Dict[str, str]]) -> None:
        skill_list = getattr(self, "codex_skill_list", None)
        if skill_list is None:
            return
        current_path = self._selected_codex_skill_path()
        skill_list.blockSignals(True)
        skill_list.clear()
        for skill in records:
            order_no = skill.get("order", "")
            name = skill.get("name", "")
            label = f"[{order_no}] {name}" if order_no else name
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, skill.get("path", ""))
            skill_list.addItem(item)
        if current_path:
            for i in range(skill_list.count()):
                item = skill_list.item(i)
                if item.data(Qt.ItemDataRole.UserRole) == current_path:
                    skill_list.setCurrentRow(i)
                    break
        skill_list.blockSignals(False)

    def _schedule_filter_codex_skill_list(self, *args) -> None:
        timer = getattr(self, "_codex_skill_filter_timer", None)
        if timer is None:
            timer = QTimer(self)
            timer.setSingleShot(True)
            timer.setInterval(120)
            timer.timeout.connect(self._filter_codex_skill_list)
            self._codex_skill_filter_timer = timer
        timer.start()

    def _filter_codex_skill_list(self) -> None:
        records = getattr(self, "_codex_skill_records", [])
        search_input = getattr(self, "codex_skill_search_input", None)
        query = search_input.text().strip() if search_input is not None else ""
        key = self._codex_skill_search_key(query)
        if not key:
            self._populate_codex_skill_list(records)
            return
        filtered = [
            skill
            for skill in records
            if key
            in self._codex_skill_search_key(
                skill.get("order", ""),
                skill.get("name", ""),
                skill.get("filename", ""),
                skill.get("trigger", ""),
            )
        ]
        self._populate_codex_skill_list(filtered)

    def _selected_codex_skill_path(self) -> str:
        skill_list = getattr(self, "codex_skill_list", None)
        item = skill_list.currentItem() if skill_list is not None else None
        if item is None:
            return ""
        return item.data(Qt.ItemDataRole.UserRole) or ""

    def _current_codex_skill_markdown(self) -> str:
        return self._codex_skill_markdown(
            self.codex_skill_name_input.text(),
            self.codex_skill_trigger_input.text(),
            self.codex_skill_body_editor.toPlainText(),
            self.codex_skill_order_input.text(),
        )

    def _codex_skill_call_prompt_text(self) -> str:
        def _line_text(attr: str) -> str:
            widget = getattr(self, attr, None)
            if widget is None:
                return ""
            try:
                return widget.text().strip()
            except Exception:
                return ""

        def _plain_text(attr: str) -> str:
            widget = getattr(self, attr, None)
            if widget is None:
                return ""
            try:
                return widget.toPlainText().strip()
            except Exception:
                return ""

        order_no = _line_text("codex_skill_order_input") or self._codex_next_skill_order()
        name = _line_text("codex_skill_name_input") or "선택 스킬"
        trigger = _line_text("codex_skill_trigger_input")
        target = ""
        try:
            target = self._codex_target_from_fields().get("path", "")
        except Exception:
            target = ""
        request_title = _line_text("codex_request_title_input")
        request_body = _plain_text("codex_request_body_editor")

        return f"""아래 사용자 스킬을 적용해서 OneNote 작업을 수행해라.

스킬:
- 주문번호: {order_no}
- 이름: {name}
- 적용 조건: {trigger or "현재 사용자가 요청한 OneNote 작업"}

작업 위치:
{target or "현재 앱에서 선택한 OneNote 작업 위치를 먼저 확인한다."}

사용자 요청:
{request_title or name}

추가 내용:
{request_body or "- 현재 선택된 스킬의 Instructions를 우선 적용한다."}

처리 기준:
- 사용자 스킬 파일에서는 `## Instructions`만 작업에 맞게 적용한다.
- OneNote 내부 조작 방식은 코덱스 전용 지침에서 필요한 때 직접 확인한다.
"""

    def _update_codex_skill_call_preview(self) -> None:
        preview = getattr(self, "codex_skill_call_preview", None)
        if preview is None:
            self._update_codex_work_order_preview()
            self._update_codex_status_summary()
            return
        self._set_plain_text_if_changed(preview, self._codex_skill_call_prompt_text())
        self._update_codex_work_order_preview()
        self._update_codex_status_summary()

    def _copy_codex_skill_call_prompt_to_clipboard(self) -> None:
        text = self._codex_skill_call_prompt_text()
        QApplication.clipboard().setText(text)
        try:
            self.connection_status_label.setText("스킬 적용 요청을 한국어로 복사했습니다.")
        except Exception:
            pass

    def _open_selected_codex_skill_file(self) -> None:
        path = self._selected_codex_skill_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 스킬을 선택하세요.")
            return
        try:
            os.startfile(path)
        except Exception as e:
            QMessageBox.warning(self, "스킬 파일 열기 실패", str(e))

    def _open_codex_skills_folder(self) -> None:
        try:
            os.makedirs(self._codex_skills_dir(), exist_ok=True)
            os.startfile(self._codex_skills_dir())
        except Exception as e:
            QMessageBox.warning(self, "스킬 폴더 열기 실패", str(e))

    def _open_codex_requests_folder(self) -> None:
        try:
            os.makedirs(self._codex_requests_dir(), exist_ok=True)
            os.startfile(self._codex_requests_dir())
        except Exception as e:
            QMessageBox.warning(self, "주문서 폴더 열기 실패", str(e))

    def _codex_work_order_text(self) -> str:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        order_no = ""
        skill_name = ""
        skill_path = self._selected_codex_skill_path()
        try:
            order_no = self.codex_skill_order_input.text().strip()
            skill_name = self.codex_skill_name_input.text().strip()
        except Exception:
            pass

        try:
            target = self._codex_target_from_fields()
        except Exception:
            target = {}

        try:
            request_text = self._codex_request_text()
        except Exception:
            request_text = ""

        return f"""# 코덱스 작업 주문서

생성 시각: {timestamp}

## 사용자 스킬

- 주문번호: {order_no or "미선택"}
- 스킬명: {skill_name or "미선택"}
- 파일: `{skill_path or "미선택"}`

## 작업 위치

- 위치 이름: {target.get("name", "")}
- 작업 경로: {target.get("path", "")}
- 전자필기장: {target.get("notebook", "")}
- 섹션 그룹: {target.get("section_group", "")}
- 섹션: {target.get("section", "")}

## 요청문

```text
{request_text}
```
"""

    def _update_codex_work_order_preview(self) -> None:
        preview = getattr(self, "codex_work_order_preview", None)
        if preview is None:
            return
        self._set_plain_text_if_changed(preview, self._codex_work_order_text())

    def _copy_codex_work_order_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_work_order_text())
        try:
            self.connection_status_label.setText("코덱스 작업 주문서를 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _save_codex_work_order(self) -> None:
        text = self._codex_work_order_text()
        order_no = self.codex_skill_order_input.text().strip() or "NO-SKILL"
        name = self.codex_skill_name_input.text().strip() or "codex-work-order"
        stamp = time.strftime("%Y%m%d-%H%M%S")
        filename = f"{stamp}-{order_no}-{self._codex_skill_slug(name)}.md"
        try:
            path = os.path.join(self._codex_requests_dir(), filename)
            self._write_text_file_atomic(path, text)
            try:
                self.connection_status_label.setText(f"코덱스 작업 주문서 저장 완료: {path}")
            except Exception:
                pass
            self._refresh_codex_work_order_list(path)
            QMessageBox.information(self, "작업 주문서 저장 완료", path)
        except Exception as e:
            QMessageBox.warning(self, "작업 주문서 저장 실패", str(e))

    def _codex_context_pack_text(self) -> str:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")

        try:
            page_reader_summary = self._codex_page_reader_result_summary_text()
        except Exception:
            page_reader_summary = ""

        try:
            skill_call = self._codex_skill_call_prompt_text()
        except Exception as e:
            skill_call = f"스킬 적용 요청 생성 실패: {e}"

        try:
            request_text = self._codex_request_text()
        except Exception:
            request_text = ""

        try:
            work_order_text = self._codex_work_order_text()
        except Exception:
            work_order_text = ""

        try:
            checklist_text = self._codex_execution_checklist_text()
        except Exception:
            checklist_text = ""

        try:
            compact_prompt = self._codex_compact_prompt_text()
        except Exception:
            compact_prompt = ""

        try:
            review_prompt = self._codex_review_prompt_text()
        except Exception:
            review_prompt = ""

        try:
            breakdown_prompt = self._codex_task_breakdown_prompt_text()
        except Exception:
            breakdown_prompt = ""

        try:
            completion_report = self._codex_completion_report_template_text()
        except Exception:
            completion_report = ""

        try:
            skill_recommendations = self._codex_skill_recommendation_report_text()
        except Exception:
            skill_recommendations = ""

        try:
            status_text = self._codex_status_summary_text()
        except Exception:
            status_text = ""

        try:
            current_skill = self._current_codex_skill_markdown()
        except Exception:
            current_skill = ""

        skill_index = ""
        try:
            self._refresh_codex_skill_list()
            with open(self._codex_skill_order_index_path(), "r", encoding="utf-8") as f:
                skill_index = f.read().strip()
        except Exception:
            skill_index = ""

        return f"""# 코덱스 작업 자료 묶음
생성 시각: {timestamp}

이 문서는 다음 코덱스 작업에 필요한 사용자 요청, 작업 위치, 사용자 스킬 자료를 묶은 자료입니다.

## 바로 실행할 지시

아래 사용자 요청과 작업 주문서를 기준으로 OneNote 작업을 수행해라.

## 내부 처리 기준

OneNote 조작 방식과 검증 기준은 코덱스 전용 지침에서 필요한 때 직접 확인한다.
사용자에게 붙여 넣는 요청에는 내부 구현 절차를 포함하지 않는다.

## 짧은 작업 요청

```text
{compact_prompt}
```

## 검토 요청

```text
{review_prompt}
```

## 작업 단계 정리 요청

```text
{breakdown_prompt}
```

## 완료 보고 양식

```markdown
{completion_report}
```

## 스킬 추천 리포트

```markdown
{skill_recommendations}
```

## 스킬 적용 요청

{skill_call}

## 페이지 읽기 결과 요약

```markdown
{page_reader_summary}
```

## 현재 상태

```text
{status_text}
```

## 요청문

```text
{request_text}
```

## 작업 주문서

````markdown
{work_order_text}
````

## 실행 체크리스트

```text
{checklist_text}
```

## 현재 스킬 초안

````markdown
{current_skill}
````

## 스킬 주문번호표

````markdown
{skill_index}
````
"""

    def _update_codex_context_pack_preview(self) -> None:
        preview = getattr(self, "codex_context_pack_preview", None)
        if preview is None:
            return
        self._set_plain_text_if_changed(preview, self._codex_context_pack_text())

    def _copy_codex_context_pack_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_context_pack_text())
        try:
            self.connection_status_label.setText(
                "코덱스 작업 자료 묶음을 한국어로 복사했습니다."
            )
        except Exception:
            pass

    def _save_codex_context_pack(self) -> None:
        text = self._codex_context_pack_text()
        stamp = time.strftime("%Y%m%d-%H%M%S")
        filename = f"{stamp}-codex-work-materials.md"
        try:
            path = os.path.join(self._codex_requests_dir(), filename)
            self._write_text_file_atomic(path, text)
            try:
                self.connection_status_label.setText(
                    f"코덱스 작업 자료 묶음 저장 완료: {path}"
                )
            except Exception:
                pass
            self._refresh_codex_work_order_list(path)
            QMessageBox.information(self, "작업 자료 묶음 저장 완료", path)
        except Exception as e:
            QMessageBox.warning(self, "작업 자료 묶음 저장 실패", str(e))

    def _codex_markdown_file_count(self, folder: str) -> int:
        try:
            if not os.path.isdir(folder):
                return 0
            folder_mtime = os.path.getmtime(folder)
            cache = getattr(self, "_codex_markdown_file_count_cache", {})
            cached = cache.get(folder)
            if cached and cached[0] == folder_mtime:
                return cached[1]

            count = sum(
                1
                for filename in os.listdir(folder)
                if filename.lower().endswith(".md")
                and filename not in ("README.md", "skill-order-index.md", "skill-audit.md")
            )
            cache[folder] = (folder_mtime, count)
            self._codex_markdown_file_count_cache = cache
            return count
        except Exception:
            return 0

    def _codex_status_summary_text(self) -> str:
        skill_count = self._codex_markdown_file_count(self._codex_skills_dir())
        request_count = self._codex_markdown_file_count(self._codex_requests_dir())
        target_count = len(self._load_codex_targets())

        action = ""
        title = ""
        request_target = ""
        try:
            action = self.codex_request_action_combo.currentText().strip()
            title = self.codex_request_title_input.text().strip()
            request_target = self.codex_request_target_input.text().strip()
        except Exception:
            pass

        skill_order = ""
        skill_name = ""
        try:
            skill_order = self.codex_skill_order_input.text().strip()
            skill_name = self.codex_skill_name_input.text().strip()
        except Exception:
            pass

        target = {}
        try:
            target = self._codex_target_from_fields()
        except Exception:
            target = {}

        draft_state = "있음" if os.path.exists(self._codex_request_draft_path()) else "없음"
        self._codex_last_status_snapshot = {
            "skill_count": skill_count,
            "request_count": request_count,
            "target_count": target_count,
            "draft_state": draft_state,
            "target": target,
            "action": action,
            "title": title,
            "request_target": request_target,
            "skill_order": skill_order,
            "skill_name": skill_name,
        }

        return f"""코덱스 현재 상태

- 스킬 파일: {skill_count}개
- 저장된 작업 주문서/자료 묶음: {request_count}개
- 저장된 작업 위치: {target_count}개
- 요청 초안: {draft_state}

현재 요청
- 작업: {action or "미지정"}
- 제목/이름: {title or "미지정"}
- 요청 대상: {request_target or target.get("path", "") or "미지정"}

선택 스킬
- 주문번호: {skill_order or "미지정"}
- 스킬명: {skill_name or "미지정"}

작업 위치
- 위치 이름: {target.get("name", "") or "미지정"}
- 전자필기장: {target.get("notebook", "") or "미지정"}
- 섹션 그룹: {target.get("section_group", "") or "미지정"}
- 섹션: {target.get("section", "") or "미지정"}
"""

    def _update_codex_status_summary(self) -> None:
        preview = getattr(self, "codex_status_summary_preview", None)
        summary_text = self._codex_status_summary_text()
        if preview is not None:
            self._set_plain_text_if_changed(preview, summary_text)

        try:
            snapshot = getattr(self, "_codex_last_status_snapshot", {})
            skill_count = snapshot.get("skill_count", 0)
            request_count = snapshot.get("request_count", 0)
            target_count = snapshot.get("target_count", 0)
            draft_state = (
                "저장됨" if snapshot.get("draft_state") == "있음" else "대기"
            )
            target = snapshot.get("target", {})
            action = snapshot.get("action", "")
            title = snapshot.get("title", "")
            skill_order = snapshot.get("skill_order", "")
            skill_name = snapshot.get("skill_name", "")

            hero_values = {
                "codex_metric_target_value": f"{target_count}개",
                "codex_metric_draft_value": draft_state,
                "codex_metric_skill_value": f"{skill_count}개",
                "codex_metric_order_value": f"{request_count}개",
                "codex_hero_target_value": (
                    target.get("path")
                    or target.get("name")
                    or "작업 위치 미지정"
                ),
                "codex_hero_request_value": title or action or "요청 대기 중",
                "codex_copy_target_value": (
                    target.get("path")
                    or target.get("name")
                    or "작업 위치 미지정"
                ),
                "codex_copy_skill_value": (
                    " / ".join(
                        part for part in (skill_order, skill_name) if part
                    )
                    or "선택 스킬 미지정"
                ),
                "codex_copy_request_value": title or action or "요청 대기 중",
            }
            for attr, value in hero_values.items():
                widget = getattr(self, attr, None)
                if widget is not None:
                    self._set_label_text_if_changed(widget, value)
        except Exception:
            pass

    def _copy_codex_status_summary_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_status_summary_text())
        try:
            self.connection_status_label.setText("코덱스 현재 상태 요약을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _codex_current_target_copy_text(self) -> str:
        try:
            target = self._codex_target_from_fields()
        except Exception:
            target = {}
        path = target.get("path", "") or target.get("name", "")
        parts = [
            f"작업 위치: {path or '미지정'}",
            f"전자필기장: {target.get('notebook', '') or '미지정'}",
            f"섹션 그룹: {target.get('section_group', '') or '미지정'}",
            f"섹션: {target.get('section', '') or '미지정'}",
        ]
        return "\n".join(parts)

    def _copy_codex_current_target_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_current_target_copy_text())
        try:
            self.connection_status_label.setText("현재 작업 위치를 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _codex_current_skill_copy_text(self) -> str:
        order_no = ""
        skill_name = ""
        try:
            order_no = self.codex_skill_order_input.text().strip()
            skill_name = self.codex_skill_name_input.text().strip()
        except Exception:
            pass
        path = self._selected_codex_skill_path()
        return "\n".join(
            [
                f"주문번호: {order_no or '미지정'}",
                f"스킬명: {skill_name or '미지정'}",
                f"파일: {path or '미선택'}",
            ]
        )

    def _copy_codex_current_skill_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_current_skill_copy_text())
        try:
            self.connection_status_label.setText("현재 선택 스킬을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _codex_current_request_copy_text(self) -> str:
        try:
            return self._codex_visible_request_text()
        except Exception:
            return ""

    def _copy_codex_current_request_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_current_request_copy_text())
        try:
            self.connection_status_label.setText("현재 요청 요약을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _selected_codex_work_order_path(self) -> str:
        order_list = getattr(self, "codex_work_order_list", None)
        item = order_list.currentItem() if order_list is not None else None
        if item is None:
            return ""
        return item.data(Qt.ItemDataRole.UserRole) or ""

    def _selected_codex_work_order_text(self) -> str:
        path = self._selected_codex_work_order_path()
        if not path:
            return ""
        with open(path, "r", encoding="utf-8") as f:
            return f.read()

    def _schedule_refresh_codex_work_order_list(self, *args) -> None:
        timer = getattr(self, "_codex_work_order_list_timer", None)
        if timer is None:
            timer = QTimer(self)
            timer.setSingleShot(True)
            timer.setInterval(160)
            timer.timeout.connect(self._refresh_codex_work_order_list)
            self._codex_work_order_list_timer = timer
        timer.start()

    def _codex_work_order_search_text(self, path: str, filename: str) -> str:
        try:
            mtime = os.path.getmtime(path)
        except Exception:
            return filename

        cache = getattr(self, "_codex_work_order_search_text_cache", {})
        cached = cache.get(path)
        if cached and cached[0] == mtime:
            return cached[1]

        text = filename
        try:
            with open(path, "r", encoding="utf-8") as f:
                text += "\n" + f.read(8000)
        except Exception:
            pass

        cache[path] = (mtime, text)
        if len(cache) > 256:
            for old_path in list(cache)[: len(cache) - 256]:
                cache.pop(old_path, None)
        self._codex_work_order_search_text_cache = cache
        return text

    def _refresh_codex_work_order_list(self, selected_path: str = "") -> None:
        order_list = getattr(self, "codex_work_order_list", None)
        if order_list is None:
            return

        current_path = selected_path or self._selected_codex_work_order_path()
        query_input = getattr(self, "codex_work_order_search_input", None)
        query = query_input.text().strip() if query_input is not None else ""
        query_key = self._codex_skill_search_key(query)
        order_list.blockSignals(True)
        order_list.clear()
        try:
            os.makedirs(self._codex_requests_dir(), exist_ok=True)
            records = []
            for filename in os.listdir(self._codex_requests_dir()):
                if not filename.lower().endswith(".md"):
                    continue
                path = os.path.join(self._codex_requests_dir(), filename)
                try:
                    mtime = os.path.getmtime(path)
                except Exception:
                    mtime = 0.0
                if query_key:
                    haystack = self._codex_work_order_search_text(path, filename)
                    if query_key not in self._codex_skill_search_key(haystack):
                        continue
                records.append((mtime, filename, path))
            records.sort(key=lambda row: (-row[0], row[1]))
            for mtime, filename, path in records:
                stamp = time.strftime("%Y-%m-%d %H:%M", time.localtime(mtime)) if mtime else ""
                label = f"{stamp}  {filename}" if stamp else filename
                item = QListWidgetItem(label)
                item.setData(Qt.ItemDataRole.UserRole, path)
                order_list.addItem(item)
        except Exception as e:
            try:
                self.connection_status_label.setText(f"작업 주문서 기록 로드 실패: {e}")
            except Exception:
                pass
        finally:
            order_list.blockSignals(False)

        if current_path:
            for i in range(order_list.count()):
                item = order_list.item(i)
                if item.data(Qt.ItemDataRole.UserRole) == current_path:
                    order_list.setCurrentRow(i)
                    self._on_codex_work_order_selected(item)
                    return
        if order_list.count() > 0:
            order_list.setCurrentRow(0)
            self._on_codex_work_order_selected(order_list.item(0))
        else:
            preview = getattr(self, "codex_work_order_history_preview", None)
            if preview is not None:
                preview.setPlainText("저장된 작업 주문서가 없습니다.")

    def _on_codex_work_order_selected(
        self, item: Optional[QListWidgetItem] = None
    ) -> None:
        preview = getattr(self, "codex_work_order_history_preview", None)
        if preview is None:
            return
        if item is None:
            order_list = getattr(self, "codex_work_order_list", None)
            item = order_list.currentItem() if order_list is not None else None
        path = item.data(Qt.ItemDataRole.UserRole) if item is not None else ""
        if not path:
            preview.setPlainText("선택된 작업 주문서가 없습니다.")
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                preview.setPlainText(f.read())
        except Exception as e:
            preview.setPlainText(f"작업 주문서를 읽지 못했습니다.\n\n{e}")

    def _copy_selected_codex_work_order_to_clipboard(self) -> None:
        path = self._selected_codex_work_order_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 작업 주문서를 선택하세요.")
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                QApplication.clipboard().setText(f.read())
            try:
                self.connection_status_label.setText("선택 작업 주문서를 클립보드에 복사했습니다.")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "작업 주문서 복사 실패", str(e))

    def _open_selected_codex_work_order_file(self) -> None:
        path = self._selected_codex_work_order_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 작업 주문서를 선택하세요.")
            return
        try:
            os.startfile(path)
        except Exception as e:
            QMessageBox.warning(self, "작업 주문서 열기 실패", str(e))

    def _delete_selected_codex_work_order(self) -> None:
        path = self._selected_codex_work_order_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 삭제할 작업 주문서를 선택하세요.")
            return
        answer = QMessageBox.question(
            self,
            "작업 주문서 삭제",
            f"선택한 작업 주문서를 삭제합니다.\n\n{path}\n\n계속하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            os.remove(path)
            self._refresh_codex_work_order_list()
            try:
                self.connection_status_label.setText("작업 주문서를 삭제했습니다.")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "작업 주문서 삭제 실패", str(e))

    def _codex_extract_fenced_section(self, text: str, heading: str) -> str:
        pattern = rf"(?ms)^##\s+{re.escape(heading)}\s*$\n+```[^\n]*\n(.*?)\n```\s*"
        match = re.search(pattern, text or "")
        if match:
            return match.group(1).strip()
        return self._codex_skill_section(text, heading)

    def _load_selected_codex_work_order_into_request(self) -> None:
        try:
            text = self._selected_codex_work_order_text()
        except Exception as e:
            QMessageBox.warning(self, "주문서 불러오기 실패", str(e))
            return
        if not text:
            QMessageBox.information(self, "안내", "먼저 작업 주문서를 선택하세요.")
            return

        request_text = self._codex_extract_fenced_section(text, "요청문")
        target_path = ""
        skill_order = ""
        title = ""

        target_match = re.search(r"(?m)^-\s*작업 경로:\s*(.+?)\s*$", text)
        if target_match:
            target_path = target_match.group(1).strip()
        order_match = re.search(r"(?m)^-\s*주문번호:\s*(.+?)\s*$", text)
        if order_match:
            skill_order = order_match.group(1).strip()
        title_match = re.search(r"(?m)^제목/이름:\s*\n(.+?)\s*(?:\n\n|\Z)", request_text)
        if title_match:
            title = title_match.group(1).strip()

        target_input = getattr(self, "codex_request_target_input", None)
        if target_input is not None and target_path:
            target_input.setText(target_path)

        title_input = getattr(self, "codex_request_title_input", None)
        if title_input is not None:
            title_input.setText(title or os.path.basename(self._selected_codex_work_order_path()))

        body_editor = getattr(self, "codex_request_body_editor", None)
        if body_editor is not None:
            body_editor.setPlainText(request_text or text)

        if skill_order:
            lookup = getattr(self, "codex_skill_order_lookup_input", None)
            if lookup is not None:
                lookup.setText(skill_order)
                self._select_codex_skill_by_order_input()

        self._update_codex_codegen_previews()
        try:
            self._scroll_codex_to_widget("codex_request_group_widget")
        except Exception:
            pass
        try:
            self.connection_status_label.setText("선택 주문서를 현재 요청으로 불러왔습니다.")
        except Exception:
            pass

    def _copy_selected_codex_work_order_followup_prompt(self) -> None:
        try:
            text = self._selected_codex_work_order_text()
        except Exception as e:
            QMessageBox.warning(self, "후속 요청 생성 실패", str(e))
            return
        if not text:
            QMessageBox.information(self, "안내", "먼저 작업 주문서를 선택하세요.")
            return

        prompt = f"""아래 이전 OneNote 작업 주문서를 이어서 처리해라.

먼저 기존 주문서의 대상 위치, 스킬, 요청문을 요약하고, 이어서 필요한 다음 행동만 제안해라.
실행이 필요하면 코덱스 전용 OneNote 조작 지침을 따르고, 완료 후 변경 항목과 검증 결과만 간단히 보고해라.

## 이전 작업 주문서

````markdown
{text}
````
"""
        QApplication.clipboard().setText(prompt)
        try:
            self.connection_status_label.setText("선택 주문서 후속 요청을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _copy_selected_codex_skill_path_to_clipboard(self) -> None:
        path = self._selected_codex_skill_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 스킬을 선택하세요.")
            return
        QApplication.clipboard().setText(path)
        try:
            self.connection_status_label.setText("선택 스킬 파일 경로를 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _select_codex_skill_by_order_input(self) -> None:
        lookup = getattr(self, "codex_skill_order_lookup_input", None)
        order_no = lookup.text().strip() if lookup is not None else ""
        if not order_no:
            order_no = self.codex_skill_order_input.text().strip()
        if not order_no:
            return

        target_key = self._codex_skill_search_key(order_no)
        records = getattr(self, "_codex_skill_records", [])
        match = None
        for record in records:
            if self._codex_skill_search_key(record.get("order", "")) == target_key:
                match = record
                break
        if match is None:
            QMessageBox.information(self, "스킬 없음", f"주문번호를 찾지 못했습니다: {order_no}")
            return

        self.codex_skill_search_input.blockSignals(True)
        self.codex_skill_search_input.setText("")
        self.codex_skill_search_input.blockSignals(False)
        self._populate_codex_skill_list(records)

        path = match.get("path", "")
        for i in range(self.codex_skill_list.count()):
            item = self.codex_skill_list.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == path:
                self.codex_skill_list.setCurrentRow(i)
                self._load_selected_codex_skill(item)
                return

    def _recommended_codex_skill_order(self) -> str:
        ranked = self._rank_codex_skill_records()
        if ranked:
            order_no = ranked[0].get("record", {}).get("order", "")
            if order_no:
                return order_no
        return "SK-001"

    def _select_recommended_codex_skill(self) -> None:
        order_no = self._recommended_codex_skill_order()
        if not getattr(self, "_codex_skill_records", None):
            self._refresh_codex_skill_list()
        target_key = self._codex_skill_search_key(order_no)
        records = getattr(self, "_codex_skill_records", [])
        if not any(
            self._codex_skill_search_key(record.get("order", "")) == target_key
            for record in records
        ):
            QMessageBox.information(
                self,
                "스킬 없음",
                f"현재 작업에 맞는 기본 스킬을 찾지 못했습니다: {order_no}",
            )
            return
        lookup = getattr(self, "codex_skill_order_lookup_input", None)
        if lookup is not None:
            lookup.setText(order_no)
        self._select_codex_skill_by_order_input()
        try:
            self.connection_status_label.setText(f"현재 작업에 맞는 스킬을 선택했습니다: {order_no}")
        except Exception:
            pass

    def _duplicate_selected_codex_skill(self) -> None:
        path = self._selected_codex_skill_path()
        if path and os.path.exists(path):
            try:
                meta = self._codex_skill_metadata_from_file(path, os.path.basename(path)[:-3])
            except Exception as e:
                QMessageBox.warning(self, "스킬 복제 실패", str(e))
                return
        else:
            meta = {
                "name": self.codex_skill_name_input.text().strip() or "새 코덱스 스킬",
                "trigger": self.codex_skill_trigger_input.text().strip(),
                "body": self.codex_skill_body_editor.toPlainText().strip(),
            }

        new_order = self._codex_next_skill_order()
        base_name = meta.get("name", "새 코덱스 스킬")
        new_name = f"{base_name} 복사본"
        new_body = meta.get("body", "")
        new_trigger = meta.get("trigger", "")
        skills_dir = self._codex_skills_dir()
        try:
            os.makedirs(skills_dir, exist_ok=True)
            path = os.path.join(skills_dir, self._codex_skill_slug(new_name) + ".md")
            if os.path.exists(path):
                path = os.path.join(skills_dir, self._codex_skill_slug(f"{new_name}-{new_order}") + ".md")
            self._write_text_file_atomic(
                path,
                self._codex_skill_markdown(new_name, new_trigger, new_body, new_order),
            )
            self._refresh_codex_skill_list()
            self.codex_skill_order_lookup_input.setText(new_order)
            self._select_codex_skill_by_order_input()
            try:
                self.connection_status_label.setText(f"스킬 복제 완료: {new_order}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "스킬 복제 실패", str(e))

    def _delete_selected_codex_skill(self) -> None:
        path = self._selected_codex_skill_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 삭제할 스킬을 선택하세요.")
            return
        filename = os.path.basename(path)
        if filename in ("README.md", "skill-order-index.md"):
            QMessageBox.warning(self, "삭제 불가", "기본 안내 파일은 삭제할 수 없습니다.")
            return
        answer = QMessageBox.question(
            self,
            "스킬 삭제",
            f"선택한 스킬 파일을 삭제합니다.\n\n{path}\n\n계속하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            os.remove(path)
            self._refresh_codex_skill_list()
            self._new_codex_skill_draft()
            try:
                self.connection_status_label.setText(f"스킬 삭제 완료: {filename}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "스킬 삭제 실패", str(e))

    def _new_codex_skill_from_clipboard(self) -> None:
        text = QApplication.clipboard().text().strip()
        if not text:
            QMessageBox.information(self, "안내", "클립보드에 스킬로 만들 내용이 없습니다.")
            return
        meta = self._codex_skill_metadata_from_text(text, "클립보드 스킬")
        order_no = meta.get("order") or self._codex_next_skill_order()
        name = meta.get("name") or "클립보드 스킬"
        if name.startswith("[") and "]" in name:
            name = name.split("]", 1)[1].strip() or "클립보드 스킬"
        self.codex_skill_order_input.setText(order_no)
        self.codex_skill_name_input.setText(name)
        self.codex_skill_trigger_input.setText(meta.get("trigger", "클립보드에서 만든 스킬"))
        self.codex_skill_body_editor.setPlainText(meta.get("body", text))
        self._update_codex_skill_call_preview()
        try:
            self.connection_status_label.setText("클립보드 내용으로 스킬 초안을 만들었습니다.")
        except Exception:
            pass

    def _codex_skill_markdown(
        self, name: str, trigger: str, body: str, order_no: str = ""
    ) -> str:
        name = (name or "새 코덱스 스킬").strip()
        trigger = (trigger or "").strip()
        body = (body or "").strip()
        order_no = (order_no or self._codex_next_skill_order()).strip()
        return f"""# {name}

## 주문번호

{order_no}

## Trigger

{trigger or "- 사용자가 이 스킬을 호출할 자연어 조건을 적는다."}

## Instructions

{body or "여기에 스킬 실행 절차, 입력 형식, 출력 기준, 검증 기준을 적는다."}
"""

    def _new_codex_skill_draft(self) -> None:
        self.codex_skill_order_input.setText(self._codex_next_skill_order())
        self.codex_skill_name_input.setText("새 코덱스 스킬")
        self.codex_skill_trigger_input.setText("이 스킬을 쓸 상황을 적는다")
        self.codex_skill_body_editor.setPlainText(
            """목표:
- 이 스킬이 해결할 작업을 명확히 적는다.

형식:
- 제목:
- 입력:
- 처리 방식:
- 출력:

검증:
- 완료 후 확인할 기준을 적는다.
"""
        )
        self._update_codex_skill_call_preview()

    def _create_default_codex_skill_set(self) -> None:
        skills_dir = self._codex_skills_dir()
        try:
            os.makedirs(skills_dir, exist_ok=True)
            existing_names: Set[str] = set()
            for filename in os.listdir(skills_dir):
                if not filename.lower().endswith(".md"):
                    continue
                if filename in ("README.md", "skill-order-index.md", "skill-audit.md"):
                    continue
                path = os.path.join(skills_dir, filename)
                try:
                    meta = self._codex_skill_metadata_from_file(path, filename[:-3])
                    existing_names.add(self._codex_skill_search_key(meta.get("name", "")))
                except Exception:
                    existing_names.add(self._codex_skill_search_key(filename[:-3]))

            created: List[str] = []
            for template in self._codex_skill_templates().values():
                name = template.get("name", "새 코덱스 스킬")
                if self._codex_skill_search_key(name) in existing_names:
                    continue
                order_no = self._codex_next_skill_order()
                path = os.path.join(skills_dir, self._codex_skill_slug(name) + ".md")
                if os.path.exists(path):
                    path = os.path.join(
                        skills_dir, self._codex_skill_slug(f"{name}-{order_no}") + ".md"
                    )
                self._write_text_file_atomic(
                    path,
                    self._codex_skill_markdown(
                        name,
                        template.get("trigger", ""),
                        template.get("body", ""),
                        order_no,
                    ),
                )
                created.append(f"{order_no} {name}")
                existing_names.add(self._codex_skill_search_key(name))

            self._refresh_codex_skill_list()
            self._update_codex_status_summary()
            if created:
                QMessageBox.information(
                    self,
                    "기본 스킬 세트 생성 완료",
                    "생성한 스킬:\n\n" + "\n".join(created),
                )
            else:
                QMessageBox.information(
                    self,
                    "기본 스킬 세트",
                    "이미 기본 스킬 양식이 모두 준비되어 있습니다.",
                )
        except Exception as e:
            QMessageBox.warning(self, "기본 스킬 세트 생성 실패", str(e))

    def _refresh_codex_skill_list(self) -> None:
        skill_list = getattr(self, "codex_skill_list", None)
        if skill_list is None:
            return
        skills_dir = self._codex_skills_dir()
        try:
            skill_index_rows = self._codex_skill_records_from_files()
            self._codex_skill_records = skill_index_rows
            try:
                self._codex_skill_records_dir_mtime = os.path.getmtime(skills_dir)
            except Exception:
                self._codex_skill_records_dir_mtime = None
            self._filter_codex_skill_list()
            self._populate_codex_skill_package_user_skill_choices()
            self._write_codex_skill_order_index(skill_index_rows)
        except Exception as e:
            try:
                self.connection_status_label.setText(f"코덱스 스킬 목록 로드 실패: {e}")
            except Exception:
                pass

    def _copy_codex_skill_order_index_to_clipboard(self) -> None:
        try:
            self._refresh_codex_skill_list()
            with open(self._codex_skill_order_index_path(), "r", encoding="utf-8") as f:
                text = f.read()
            QApplication.clipboard().setText(text)
            try:
                self.connection_status_label.setText("스킬 주문번호표를 클립보드에 복사했습니다.")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "주문번호표 복사 실패", str(e))

    def _save_codex_skill_draft(self) -> None:
        order_no = self.codex_skill_order_input.text().strip()
        name = self.codex_skill_name_input.text().strip()
        trigger = self.codex_skill_trigger_input.text().strip()
        body = self.codex_skill_body_editor.toPlainText().strip()
        if not name:
            name = "새 코덱스 스킬"
        if not order_no:
            order_no = self._codex_next_skill_order()
            self.codex_skill_order_input.setText(order_no)
        skills_dir = self._codex_skills_dir()
        try:
            path = os.path.join(skills_dir, self._codex_skill_slug(name) + ".md")
            self._write_text_file_atomic(
                path,
                self._codex_skill_markdown(name, trigger, body, order_no),
            )
            self._refresh_codex_skill_list()
            self._update_codex_skill_call_preview()
            try:
                self.connection_status_label.setText(
                    f"코덱스 스킬 저장 완료: {order_no} / {path}"
                )
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "스킬 저장 실패", str(e))

    def _load_selected_codex_skill(self, item: Optional[QListWidgetItem] = None) -> None:
        if item is not None and not isinstance(item, QListWidgetItem):
            item = None
        if item is None:
            skill_list = getattr(self, "codex_skill_list", None)
            item = skill_list.currentItem() if skill_list is not None else None
        if item is None:
            return
        path = item.data(Qt.ItemDataRole.UserRole)
        if not path:
            return
        try:
            meta = self._codex_skill_metadata_from_file(path, item.text())
            self.codex_skill_order_input.setText(meta.get("order", ""))
            self.codex_skill_name_input.setText(meta.get("name", ""))
            self.codex_skill_trigger_input.setText(meta.get("trigger", ""))
            self.codex_skill_body_editor.setPlainText(meta.get("body", ""))
            self._update_codex_skill_call_preview()
            try:
                self.connection_status_label.setText(f"코덱스 스킬 불러옴: {path}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "스킬 불러오기 실패", str(e))

    def _copy_codex_skill_prompt_to_clipboard(self) -> None:
        text = self._current_codex_skill_markdown()
        QApplication.clipboard().setText(text)
        try:
            self.connection_status_label.setText("코덱스 스킬 초안을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _codex_targets_path(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "onenote-targets.json")

    def _default_codex_targets(self) -> List[Dict[str, str]]:
        return [
            {
                "name": "임시 메모 - 미정리",
                "path": "생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리",
                "notebook": "생산성도구-임시 메모",
                "section_group": "A 미정리-생성 메모",
                "section": "미정리",
                "section_group_id": "{2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}",
                "section_id": "{175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}",
            }
        ]

    def _load_codex_targets(self) -> List[Dict[str, str]]:
        path = self._codex_targets_path()
        try:
            if os.path.exists(path):
                file_mtime = os.path.getmtime(path)
                cached = getattr(self, "_codex_targets_cache", None)
                if cached and cached[0] == file_mtime:
                    return [dict(t) for t in cached[1]]

                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                targets = data.get("targets") if isinstance(data, dict) else data
                if isinstance(targets, list) and targets:
                    normalized = [dict(t) for t in targets if isinstance(t, dict)]
                    self._codex_targets_cache = (file_mtime, normalized)
                    return [dict(t) for t in normalized]
        except Exception as e:
            try:
                self.connection_status_label.setText(f"코덱스 작업 위치 로드 실패: {e}")
            except Exception:
                pass
        defaults = self._default_codex_targets()
        self._codex_targets_cache = (None, defaults)
        return [dict(t) for t in defaults]

    def _write_codex_targets(self, targets: List[Dict[str, str]]) -> None:
        path = self._codex_targets_path()
        payload = {
            "version": 1,
            "targets": targets,
        }
        self._write_json_file_atomic(path, payload)
        try:
            self._codex_targets_cache = (
                os.path.getmtime(path),
                [dict(t) for t in targets if isinstance(t, dict)],
            )
        except Exception:
            self._codex_targets_cache = None

    def _selected_codex_target_profile(self) -> Dict[str, str]:
        combo = getattr(self, "codex_target_combo", None)
        if combo is not None:
            data = combo.currentData()
            if isinstance(data, dict):
                return data
        try:
            profile = self._codex_target_from_fields()
            if profile.get("path") or profile.get("notebook"):
                return profile
        except Exception:
            pass
        return self._default_codex_targets()[0]

    def _codex_target_from_fields(self) -> Dict[str, str]:
        def _text(attr: str) -> str:
            widget = getattr(self, attr, None)
            if widget is None:
                return ""
            try:
                return widget.text().strip()
            except Exception:
                return ""

        return {
            "name": _text("codex_target_name_input") or "새 대상",
            "path": _text("codex_target_path_input"),
            "notebook": _text("codex_target_notebook_input"),
            "section_group": _text("codex_target_group_input"),
            "section": _text("codex_target_section_input"),
            "section_group_id": _text("codex_target_group_id_input"),
            "section_id": _text("codex_target_section_id_input"),
        }

    def _populate_codex_target_fields(self, profile: Dict[str, str]) -> None:
        mapping = {
            "codex_target_name_input": profile.get("name", ""),
            "codex_target_path_input": profile.get("path", ""),
            "codex_target_notebook_input": profile.get("notebook", ""),
            "codex_target_group_input": profile.get("section_group", ""),
            "codex_target_section_input": profile.get("section", ""),
            "codex_target_group_id_input": profile.get("section_group_id", ""),
            "codex_target_section_id_input": profile.get("section_id", ""),
        }
        for attr, value in mapping.items():
            widget = getattr(self, attr, None)
            if widget is not None:
                if widget.text() == value:
                    continue
                widget.blockSignals(True)
                widget.setText(value)
                widget.blockSignals(False)

    def _refresh_codex_target_combo(self, selected_name: Optional[str] = None) -> None:
        combo = getattr(self, "codex_target_combo", None)
        if combo is None:
            return
        current_name = selected_name or combo.currentText()
        combo.blockSignals(True)
        combo.clear()
        for profile in self._load_codex_targets():
            combo.addItem(profile.get("name", "대상"), profile)
        idx = combo.findText(current_name)
        if idx >= 0:
            combo.setCurrentIndex(idx)
        combo.blockSignals(False)
        self._on_codex_target_selected()

    def _on_codex_target_selected(self) -> None:
        profile = self._selected_codex_target_profile()
        self._populate_codex_target_fields(profile)
        self._apply_codex_target_to_request()

    def _apply_codex_target_to_request(self) -> None:
        target_input = getattr(self, "codex_request_target_input", None)
        if target_input is None:
            return
        profile = self._codex_target_from_fields()
        target_text = (
            profile.get("path")
            or "생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리"
        )
        if target_input.text() != target_text:
            target_input.setText(target_text)
        self._schedule_codex_codegen_previews()

    def _codex_target_profile_from_fav_item(
        self, item: Optional[QTreeWidgetItem]
    ) -> Dict[str, str]:
        if item is None:
            return {}

        try:
            node_type = item.data(0, ROLE_TYPE)
        except Exception:
            return {}

        if node_type not in ("section", "notebook"):
            return {}

        def _clean(value: Any) -> str:
            if value is None:
                return ""
            return str(value).strip()

        try:
            payload = item.data(0, ROLE_DATA) or {}
        except Exception:
            payload = {}
        if not isinstance(payload, dict):
            payload = {}

        target = payload.get("target") or {}
        if not isinstance(target, dict):
            target = {}

        display_name = _clean(item.text(0))
        notebook = _clean(
            target.get("notebook")
            or target.get("notebook_text")
            or target.get("notebookName")
        )
        section_group = _clean(
            target.get("section_group")
            or target.get("section_group_text")
            or target.get("sectionGroup")
            or target.get("sectionGroupName")
        )
        section = _clean(
            target.get("section")
            or target.get("section_text")
            or target.get("sectionName")
        )

        ancestors: List[tuple[str, str]] = []
        parent = item.parent()
        while parent is not None:
            try:
                parent_type = _clean(parent.data(0, ROLE_TYPE))
                parent_name = _clean(parent.text(0))
            except Exception:
                parent_type = ""
                parent_name = ""
            if parent_name:
                ancestors.append((parent_type, parent_name))
            parent = parent.parent()
        ancestors.reverse()

        if node_type == "notebook":
            notebook = notebook or display_name
            section_group = ""
            section = ""
        else:
            section = section or display_name
            if not notebook:
                for ancestor_type, ancestor_name in ancestors:
                    if ancestor_type == "notebook":
                        notebook = ancestor_name
                        break
            if not section_group:
                group_names = [
                    ancestor_name
                    for ancestor_type, ancestor_name in ancestors
                    if ancestor_type == "group"
                ]
                if group_names:
                    section_group = group_names[-1]

        path = _clean(target.get("path") or target.get("hierarchy_path"))
        if not path:
            path_parts: List[str] = []
            for candidate in (notebook, section_group, section):
                if candidate and candidate not in path_parts:
                    path_parts.append(candidate)
            if not path_parts:
                path_parts = [name for _, name in ancestors]
                if display_name:
                    path_parts.append(display_name)
            path = " > ".join(path_parts)

        if node_type == "notebook":
            name = notebook or display_name or path or "전자필기장"
        else:
            name_parts = [part for part in (notebook, section) if part]
            name = " - ".join(name_parts) or display_name or path or "섹션"

        section_group_id = _clean(
            target.get("section_group_id")
            or target.get("sectionGroupId")
            or target.get("section_group_id_text")
        )
        section_id = _clean(target.get("section_id") or target.get("sectionId"))
        if node_type == "section" and not section_id:
            section_id = _clean(target.get("id"))

        return {
            "name": name,
            "path": path,
            "notebook": notebook,
            "section_group": section_group,
            "section": section,
            "section_group_id": section_group_id,
            "section_id": section_id,
        }

    def _sync_codex_target_from_fav_item(
        self,
        item: Optional[QTreeWidgetItem],
        *,
        switch_to_codex: bool = False,
    ) -> bool:
        profile = self._codex_target_profile_from_fav_item(item)
        if not profile:
            return False

        self._populate_codex_target_fields(profile)
        self._apply_codex_target_to_request()
        try:
            self._update_codex_status_summary()
        except Exception:
            pass

        if switch_to_codex:
            self._scroll_codex_to_widget("codex_target_group_widget")

        try:
            self.connection_status_label.setText(
                f"코덱스 작업 위치 반영: {profile.get('path') or profile.get('name')}"
            )
        except Exception:
            pass
        return True

    def _sync_codex_target_from_current_fav_item(self) -> None:
        try:
            item = self.fav_tree.currentItem()
        except Exception:
            item = None
        self._sync_codex_target_from_fav_item(item)

    def _save_codex_target_profile(self) -> None:
        profile = self._codex_target_from_fields()
        targets = self._load_codex_targets()
        replaced = False
        for i, old in enumerate(targets):
            if old.get("name") == profile.get("name"):
                targets[i] = profile
                replaced = True
                break
        if not replaced:
            targets.append(profile)
        try:
            self._write_codex_targets(targets)
            self._refresh_codex_target_combo(profile.get("name"))
            try:
                self.connection_status_label.setText(
                    f"코덱스 작업 위치 저장 완료: {profile.get('name')}"
                )
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "작업 위치 저장 실패", str(e))

    def _new_codex_target_profile(self) -> None:
        combo = getattr(self, "codex_target_combo", None)
        if combo is not None:
            combo.blockSignals(True)
            combo.setCurrentIndex(-1)
            combo.blockSignals(False)
        self._populate_codex_target_fields(
            {
                "name": "새 대상",
                "path": "",
                "notebook": "",
                "section_group": "",
                "section": "",
                "section_group_id": "",
                "section_id": "",
            }
        )
        self._apply_codex_target_to_request()

    def _build_codex_target_group(self) -> QWidget:
        group = QWidget()
        group.setObjectName("CodexCard")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(8)

        header = QLabel("📍 OneNote 작업 위치 설정")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        desc = QLabel("왼쪽 패널에서 선택한 전자필기장/섹션이 작업 위치로 들어옵니다.")
        desc.setObjectName("CodexHeroSubtitle")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        lookup_row = QHBoxLayout()
        lookup_row.setSpacing(6)
        self.codex_location_lookup_toggle = QToolButton()
        self.codex_location_lookup_toggle.setText("OneNote 조회 OFF")
        self.codex_location_lookup_toggle.setMinimumWidth(0)
        self.codex_location_lookup_toggle.setSizePolicy(
            QSizePolicy.Policy.Ignored,
            QSizePolicy.Policy.Fixed,
        )
        self.codex_location_lookup_toggle.setCheckable(True)
        self.codex_location_lookup_toggle.toggled.connect(
            self._set_codex_location_lookup_enabled
        )
        lookup_row.addWidget(self.codex_location_lookup_toggle, stretch=1)

        self.codex_location_lookup_refresh_btn = QToolButton()
        self.codex_location_lookup_refresh_btn.setText("조회")
        self.codex_location_lookup_refresh_btn.setMinimumWidth(0)
        self.codex_location_lookup_refresh_btn.setSizePolicy(
            QSizePolicy.Policy.Ignored,
            QSizePolicy.Policy.Fixed,
        )
        self.codex_location_lookup_refresh_btn.clicked.connect(
            self._refresh_codex_location_lookup
        )
        lookup_row.addWidget(self.codex_location_lookup_refresh_btn, stretch=1)
        layout.addLayout(lookup_row)

        lookup_form = QGridLayout()
        lookup_form.setHorizontalSpacing(6)
        lookup_form.setVerticalSpacing(6)

        def add_lookup_combo(row: int, label: str, attr: str, handler) -> WheelSafeComboBox:
            lbl = QLabel(label)
            lbl.setMinimumWidth(54)
            combo = WheelSafeComboBox()
            self._configure_codex_lookup_combo(combo)
            combo.currentIndexChanged.connect(handler)
            setattr(self, attr, combo)
            lookup_form.addWidget(lbl, row, 0)
            lookup_form.addWidget(combo, row, 1)
            return combo

        add_lookup_combo(
            0,
            "필기장",
            "codex_location_notebook_combo",
            self._on_codex_location_notebook_selected,
        )
        add_lookup_combo(
            1,
            "섹션그룹",
            "codex_location_group_combo",
            self._on_codex_location_group_selected,
        )
        add_lookup_combo(
            2,
            "섹션",
            "codex_location_section_combo",
            self._on_codex_location_section_selected,
        )
        layout.addLayout(lookup_form)
        self._set_codex_location_lookup_enabled(False)

        form_layout = QGridLayout()
        form_layout.setSpacing(6)

        def add_line(label_text: str, attr: str, placeholder: str = "", hidden: bool = False) -> QLineEdit:
            lbl = QLabel(label_text)
            lbl.setMinimumWidth(64)
            field = QLineEdit()
            field.setPlaceholderText(placeholder)
            setattr(self, attr, field)

            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0,0,0,0)
            row_layout.setSpacing(4)
            row_layout.addWidget(lbl)
            row_layout.addWidget(field, stretch=1)

            setattr(self, f"{attr}_row", row_widget)
            form_layout.addWidget(row_widget, form_layout.rowCount(), 0)
            row_widget.setVisible(not hidden)
            return field

        self.codex_target_name_input = add_line("위치 이름", "codex_target_name_input", "예: 업무일지 전용")
        self.codex_target_path_input = add_line("작업 경로", "codex_target_path_input", "전자필기장 > 섹션그룹 > 섹션")
        self.codex_target_notebook_input = add_line("전자필기장", "codex_target_notebook_input")
        self.codex_target_group_input = add_line("섹션 그룹", "codex_target_group_input")
        self.codex_target_section_input = add_line("섹션", "codex_target_section_input")
        self.codex_target_group_id_input = add_line("그룹 ID", "codex_target_group_id_input", hidden=True)
        self.codex_target_section_id_input = add_line("섹션 ID", "codex_target_section_id_input", hidden=True)
        layout.addLayout(form_layout)

        actions = QHBoxLayout()
        actions.setSpacing(6)

        apply_btn = QToolButton()
        apply_btn.setText("🎯 요청에 넣기")
        apply_btn.setProperty("variant", "primary")
        apply_btn.clicked.connect(self._apply_codex_target_to_request)

        save_btn = QToolButton()
        save_btn.setText("💾 위치 저장")
        save_btn.clicked.connect(self._save_codex_target_profile)

        actions.addWidget(apply_btn, stretch=2)
        actions.addWidget(save_btn, stretch=1)
        layout.addLayout(actions)

        util_buttons = QGridLayout()
        util_buttons.setHorizontalSpacing(6)
        util_buttons.setVerticalSpacing(6)
        target_tools = [
            ("📋 위치 복사", self._copy_codex_target_profile_json_to_clipboard),
            ("🗂️ 저장위치 복사", self._copy_codex_all_targets_json_to_clipboard),
            ("🧭 위치조회 요청", self._copy_codex_onenote_inventory_script_to_clipboard),
            ("📄 페이지목록 요청", self._copy_codex_page_reader_script_to_clipboard),
        ]
        for index, (text, cb) in enumerate(target_tools):
            btn = QToolButton()
            btn.setText(text)
            btn.setMinimumHeight(28)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            btn.clicked.connect(cb)
            util_buttons.addWidget(btn, index // 2, index % 2)
        layout.addLayout(util_buttons)

        for attr in (
            "codex_target_name_input",
            "codex_target_path_input",
            "codex_target_notebook_input",
            "codex_target_group_input",
            "codex_target_section_input",
            "codex_target_group_id_input",
            "codex_target_section_id_input",
        ):
            widget = getattr(self, attr, None)
            if widget is not None:
                widget.textChanged.connect(self._schedule_codex_codegen_previews)
        self.codex_target_path_input.textChanged.connect(self._apply_codex_target_to_request)
        self._refresh_codex_target_combo()
        self._update_codex_codegen_previews()

        return group

    def _codex_visible_request_text(self) -> str:
        action = self.codex_request_action_combo.currentText()
        target = self.codex_request_target_input.text().strip()
        title = self.codex_request_title_input.text().strip()
        body = self.codex_request_body_editor.toPlainText().strip()

        return f"""작업:
{action}

대상 경로:
{target or "생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리"}

제목/이름:
{title or "코덱스 작업"}

본문/내용:
{body or "- 필요한 내용을 여기에 작성한다."}
"""

    def _codex_request_text(self) -> str:
        return self._codex_visible_request_text()

    def _update_codex_request_preview(self) -> None:
        preview = getattr(self, "codex_request_preview", None)
        if preview is None:
            return
        self._set_plain_text_if_changed(preview, self._codex_visible_request_text())

    def _copy_codex_request_to_clipboard(self) -> None:
        text = self._codex_request_text()
        QApplication.clipboard().setText(text)
        try:
            self.connection_status_label.setText("코덱스 OneNote 요청문을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _codex_request_draft_payload(self) -> Dict[str, Any]:
        def line(attr: str) -> str:
            widget = getattr(self, attr, None)
            if widget is None:
                return ""
            try:
                return widget.text().strip()
            except Exception:
                return ""

        def plain(attr: str) -> str:
            widget = getattr(self, attr, None)
            if widget is None:
                return ""
            try:
                return widget.toPlainText()
            except Exception:
                return ""

        action = ""
        action_combo = getattr(self, "codex_request_action_combo", None)
        if action_combo is not None:
            try:
                action = action_combo.currentText().strip()
            except Exception:
                action = ""

        preset_key = ""
        preset_combo = getattr(self, "codex_request_preset_combo", None)
        if preset_combo is not None:
            try:
                preset_key = str(preset_combo.currentData() or "")
            except Exception:
                preset_key = ""

        return {
            "version": 1,
            "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "preset": preset_key,
            "request": {
                "action": action,
                "target": line("codex_request_target_input"),
                "title": line("codex_request_title_input"),
                "body": plain("codex_request_body_editor"),
            },
            "target_profile": self._codex_target_from_fields(),
        }

    def _save_codex_request_draft(self) -> None:
        try:
            path = self._codex_request_draft_path()
            self._write_json_file_atomic(path, self._codex_request_draft_payload())
            try:
                self.connection_status_label.setText(f"코덱스 요청 초안 저장 완료: {path}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "요청 초안 저장 실패", str(e))

    def _load_codex_request_draft(self) -> None:
        path = self._codex_request_draft_path()
        if not os.path.exists(path):
            QMessageBox.information(self, "요청 초안 없음", "저장된 코덱스 요청 초안이 없습니다.")
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            request = payload.get("request", {}) if isinstance(payload, dict) else {}
            target_profile = payload.get("target_profile", {}) if isinstance(payload, dict) else {}
            if not isinstance(request, dict):
                request = {}
            if not isinstance(target_profile, dict):
                target_profile = {}

            action_combo = getattr(self, "codex_request_action_combo", None)
            if action_combo is not None:
                idx = action_combo.findText(str(request.get("action", "")))
                if idx >= 0:
                    action_combo.setCurrentIndex(idx)

            target_input = getattr(self, "codex_request_target_input", None)
            if target_input is not None:
                target_input.setText(str(request.get("target", "")))

            title_input = getattr(self, "codex_request_title_input", None)
            if title_input is not None:
                title_input.setText(str(request.get("title", "")))

            body_editor = getattr(self, "codex_request_body_editor", None)
            if body_editor is not None:
                body_editor.setPlainText(str(request.get("body", "")))

            if target_profile:
                self._populate_codex_target_fields(target_profile)
            self._update_codex_codegen_previews()
            try:
                self.connection_status_label.setText(f"코덱스 요청 초안 불러오기 완료: {path}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "요청 초안 불러오기 실패", str(e))

    def _codex_execution_checklist_text(self) -> str:
        payload = self._codex_request_draft_payload()
        request = payload.get("request", {})
        target = payload.get("target_profile", {})
        return f"""# OneNote 작업 실행 순서
생성 시각: {time.strftime("%Y-%m-%d %H:%M:%S")}

## 현재 작업

- 작업: {request.get("action", "") or "미지정"}
- 제목/이름: {request.get("title", "") or "미지정"}
- 대상 경로: {request.get("target", "") or target.get("path", "") or "미지정"}

## 실행 전

- 대상 전자필기장/섹션/섹션 그룹이 맞는지 확인한다.
- 삭제/이동 작업은 작업 전 구조를 기록한다.
- 내부 처리 방식은 코덱스 전용 지침을 따른다.

## 실행 중

- 실제 OneNote 조작은 코덱스 전용 지침의 작업별 절차를 따른다.
- 사용자 스킬은 형식/내용 처리 기준으로만 적용한다.
- 예외가 나면 OneNote 상태와 대상 캐시를 다시 조회한다.

## 실행 후 검증

- 코덱스 전용 지침의 검증 기준으로 결과를 확인한다.
- 생성/수정한 제목과 본문 일부가 실제 반영됐는지 확인한다.
- 검증 결과와 다음 확인 항목을 사용자에게 짧게 보고한다.
"""

    def _copy_codex_execution_checklist_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_execution_checklist_text())
        try:
            self.connection_status_label.setText("OneNote 작업 실행 순서를 한국어로 복사했습니다.")
        except Exception:
            pass

    def _append_clipboard_to_codex_request_body(self) -> None:
        body_editor = getattr(self, "codex_request_body_editor", None)
        if body_editor is None:
            return
        clip = QApplication.clipboard().text().strip()
        if not clip:
            QMessageBox.information(self, "클립보드 비어 있음", "요청 본문에 추가할 클립보드 텍스트가 없습니다.")
            return

        current = body_editor.toPlainText().rstrip()
        stamp = time.strftime("%Y-%m-%d %H:%M:%S")
        addition = f"클립보드 자료 ({stamp}):\n{clip}"
        body_editor.setPlainText((current + "\n\n" + addition).strip())
        self._update_codex_codegen_previews()
        try:
            self.connection_status_label.setText("클립보드 텍스트를 코덱스 요청 본문에 추가했습니다.")
        except Exception:
            pass

    def _replace_codex_request_body_from_clipboard(self) -> None:
        body_editor = getattr(self, "codex_request_body_editor", None)
        if body_editor is None:
            return
        clip = QApplication.clipboard().text().strip()
        if not clip:
            QMessageBox.information(self, "클립보드 비어 있음", "요청 본문으로 교체할 클립보드 텍스트가 없습니다.")
            return
        body_editor.setPlainText(clip)
        self._update_codex_codegen_previews()
        try:
            self.connection_status_label.setText("클립보드 텍스트로 코덱스 요청 본문을 교체했습니다.")
        except Exception:
            pass

    def _codex_compact_prompt_text(self) -> str:
        try:
            target = self._codex_target_from_fields()
        except Exception:
            target = {}

        try:
            request = self._codex_request_draft_payload().get("request", {})
        except Exception:
            request = {}

        order_no = ""
        skill_name = ""
        try:
            order_no = self.codex_skill_order_input.text().strip()
            skill_name = self.codex_skill_name_input.text().strip()
        except Exception:
            pass

        return f"""아래 OneNote 작업을 바로 수행해라.

스킬:
- 주문번호: {order_no or "미지정"}
- 스킬명: {skill_name or "미지정"}

작업:
- 유형: {request.get("action", "") or "미지정"}
- 제목/이름: {request.get("title", "") or "미지정"}
- 대상 경로: {request.get("target", "") or target.get("path", "") or "미지정"}

본문:
{request.get("body", "") or "- 사용자가 제공한 내용을 OneNote에 정리한다."}

처리 기준:
- 내부 처리 방식은 코덱스 전용 지침을 따른다.
- 사용자 스킬은 글쓰기 형식과 내용 정리에만 적용한다.

완료 후 변경한 항목과 검증 결과만 간단히 보고해라.
"""

    def _copy_codex_compact_prompt_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_compact_prompt_text())
        try:
            self.connection_status_label.setText("짧은 작업 요청을 한국어로 복사했습니다.")
        except Exception:
            pass

    def _codex_review_prompt_text(self) -> str:
        return f"""아래 OneNote 작업 계획을 검토해라. 실행하지 말고 위험 요소, 빠진 확인, 잘못된 대상 위치 가능성만 점검해라.

## 현재 상태

{self._codex_status_summary_text()}

## 실행 체크리스트

{self._codex_execution_checklist_text()}

## 요청문

{self._codex_request_text()}

검토 결과는 다음 형식으로 짧게 작성해라.
- 실행 전 막아야 할 문제:
- 대상 위치 확인 필요:
- 검증 방법:
"""

    def _copy_codex_review_prompt_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_review_prompt_text())
        try:
            self.connection_status_label.setText("검토 요청을 한국어로 복사했습니다.")
        except Exception:
            pass

    def _codex_task_breakdown_prompt_text(self) -> str:
        return f"""아래 OneNote 작업을 실행 가능한 하위 작업으로 분해해라. 아직 실행하지 말고 순서, 필요한 조회, 검증 기준만 작성해라.

## 현재 요청

{self._codex_request_text()}

출력 형식:
- 목표:
- 선행 조회:
- 작업 순서:
- 검증:
- 실패 시 복구:
"""

    def _copy_codex_task_breakdown_prompt_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_task_breakdown_prompt_text())
        try:
            self.connection_status_label.setText("작업 단계 정리 요청을 한국어로 복사했습니다.")
        except Exception:
            pass

    def _codex_completion_report_template_text(self) -> str:
        try:
            request = self._codex_request_draft_payload().get("request", {})
        except Exception:
            request = {}
        return f"""# OneNote 작업 완료 보고 양식

## 요청

- 작업: {request.get("action", "") or "미지정"}
- 제목/이름: {request.get("title", "") or "미지정"}
- 대상: {request.get("target", "") or "미지정"}

## 수행한 작업

-

## 변경한 OneNote 항목

- 전자필기장:
- 섹션 그룹:
- 섹션:
- 페이지:

## 검증 결과

- 확인 방식: 코덱스 전용 지침 기준
- 확인한 값:
- 결과:

## 남은 확인 사항

-
"""

    def _copy_codex_completion_report_template_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_completion_report_template_text())
        try:
            self.connection_status_label.setText("OneNote 작업 완료 보고 양식을 한국어로 복사했습니다.")
        except Exception:
            pass

    def _codex_current_request_text_blob(self) -> str:
        parts: List[str] = []
        try:
            parts.append(self.codex_request_action_combo.currentText())
            parts.append(self.codex_request_target_input.text())
            parts.append(self.codex_request_title_input.text())
            parts.append(self.codex_request_body_editor.toPlainText())
        except Exception:
            pass
        return "\n".join(p for p in parts if p)

    def _codex_request_tokens(self) -> Set[str]:
        text = unicodedata.normalize("NFKC", self._codex_current_request_text_blob()).casefold()
        return {
            token
            for token in re.findall(r"[0-9A-Za-z가-힣_]{2,}", text)
            if token not in {"one", "note", "onenote", "com", "api"}
        }

    def _rank_codex_skill_records(self) -> List[Dict[str, Any]]:
        if not getattr(self, "_codex_skill_records", None):
            self._refresh_codex_skill_list()

        request_key = self._codex_skill_search_key(self._codex_current_request_text_blob())
        request_tokens = self._codex_request_tokens()
        ranked: List[Dict[str, Any]] = []

        for record in getattr(self, "_codex_skill_records", []):
            text = " ".join(
                [
                    record.get("order", ""),
                    record.get("name", ""),
                    record.get("filename", ""),
                    record.get("trigger", ""),
                ]
            )
            try:
                path = record.get("path", "")
                if path and os.path.exists(path):
                    text += "\n" + self._codex_skill_metadata_from_file(
                        path, record.get("name", "")
                    ).get("body", "")
            except Exception:
                pass

            skill_key = self._codex_skill_search_key(text)
            skill_tokens = {
                token.casefold()
                for token in re.findall(r"[0-9A-Za-z가-힣_]{2,}", unicodedata.normalize("NFKC", text))
            }
            hits = sorted(request_tokens & skill_tokens)
            score = len(hits) * 10
            if record.get("name") and self._codex_skill_search_key(record.get("name", "")) in request_key:
                score += 25
            if record.get("trigger") and self._codex_skill_search_key(record.get("trigger", "")) in request_key:
                score += 20
            if record.get("order") and self._codex_skill_search_key(record.get("order", "")) in request_key:
                score += 50
            if score > 0:
                ranked.append(
                    {
                        "score": score,
                        "record": record,
                        "hits": hits[:8],
                    }
                )

        ranked.sort(
            key=lambda item: (
                -int(item.get("score", 0)),
                item.get("record", {}).get("order", "") or "ZZZ",
                item.get("record", {}).get("name", ""),
            )
        )
        return ranked

    def _codex_skill_recommendation_report_text(self) -> str:
        ranked = self._rank_codex_skill_records()
        request = self._codex_current_request_text_blob().strip()
        rows = []
        for item in ranked[:10]:
            record = item.get("record", {})
            hits = ", ".join(item.get("hits", [])) or "-"
            rows.append(
                f"| {item.get('score', 0)} | {record.get('order', '') or '-'} | "
                f"{record.get('name', '') or '-'} | {hits} | `{record.get('filename', '')}` |"
            )

        if not rows:
            rows.append("| 0 | - | 추천 없음 | 현재 요청과 겹치는 스킬 키워드가 없습니다. | - |")

        return "\n".join(
            [
                "# 코덱스 스킬 추천 리포트",
                "",
                "## 현재 요청 요약",
                "",
                request or "요청 내용 없음",
                "",
                "## 추천 스킬",
                "",
                "| 점수 | 주문번호 | 스킬 | 근거 키워드 | 파일 |",
                "| ---: | --- | --- | --- | --- |",
                *rows,
                "",
                "추천 점수는 현재 요청의 단어가 스킬 이름, 호출 조건, 본문과 겹치는 정도로 계산한다.",
                "",
            ]
        )

    def _copy_codex_skill_recommendation_report_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_skill_recommendation_report_text())
        try:
            self.connection_status_label.setText("코덱스 스킬 추천 리포트를 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _select_best_codex_skill_recommendation(self) -> None:
        ranked = self._rank_codex_skill_records()
        if not ranked:
            QMessageBox.information(self, "추천 스킬 없음", "현재 요청에 맞는 추천 스킬을 찾지 못했습니다.")
            return
        record = ranked[0].get("record", {})
        order_no = record.get("order", "")
        if not order_no:
            QMessageBox.information(self, "추천 스킬 선택 실패", "추천된 스킬에 주문번호가 없습니다.")
            return
        lookup = getattr(self, "codex_skill_order_lookup_input", None)
        if lookup is not None:
            lookup.setText(order_no)
        self._select_codex_skill_by_order_input()
        try:
            self.connection_status_label.setText(
                f"추천 스킬 선택: {order_no} / {record.get('name', '')}"
            )
        except Exception:
            pass

    def _codex_request_presets(self) -> Dict[str, Dict[str, str]]:
        today = time.strftime("%Y-%m-%d")
        return {
            "quick_note": {
                "name": "빠른 메모",
                "action": "페이지 추가",
                "title": f"{today} 빠른 메모",
                "body": """핵심:
-

세부 내용:
-

다음 행동:
-
""",
            },
            "meeting": {
                "name": "회의록",
                "action": "나의 기본 글쓰기 형식으로 페이지 작성",
                "title": f"{today} 회의록",
                "body": """회의명:
참석자:

논의 요약:
-

결정 사항:
-

액션 아이템:
- 담당자 / 마감 / 확인 방법:

보류/리스크:
-
""",
            },
            "daily_log": {
                "name": "업무일지",
                "action": "나의 기본 글쓰기 형식으로 페이지 작성",
                "title": f"{today} 업무일지",
                "body": """오늘 한 일:
-

결정한 것:
-

막힌 것:
-

다음 행동:
-
""",
            },
            "project_plan": {
                "name": "작업 계획",
                "action": "페이지 추가",
                "title": "작업 계획",
                "body": """목표:
-

범위:
- 포함:
- 제외:

실행 순서:
1.
2.
3.

검증:
-

리스크:
-
""",
            },
            "cleanup": {
                "name": "OneNote 정리",
                "action": "나의 기본 글쓰기 형식으로 페이지 작성",
                "title": "OneNote 정리 계획",
                "body": """현재 위치:
-

바꿀 위치:
-

정리할 항목:
-

보존할 항목:
-

검증 방법:
- 작업 전후 구조 확인
""",
            },
            "weekly_review": {
                "name": "주간 회고",
                "action": "페이지 추가",
                "title": f"{today} 주간 회고",
                "body": """이번 주 완료:
-

이번 주 미완료:
-

배운 점:
-

다음 주 우선순위:
1.
2.
3.
""",
            },
        }

    def _apply_codex_request_preset(self) -> None:
        combo = getattr(self, "codex_request_preset_combo", None)
        if combo is None:
            return
        preset = self._codex_request_presets().get(combo.currentData())
        if not preset:
            return

        action_combo = getattr(self, "codex_request_action_combo", None)
        if action_combo is not None:
            idx = action_combo.findText(preset.get("action", ""))
            if idx >= 0:
                action_combo.setCurrentIndex(idx)

        title_input = getattr(self, "codex_request_title_input", None)
        if title_input is not None:
            title_input.setText(preset.get("title", ""))

        body_editor = getattr(self, "codex_request_body_editor", None)
        if body_editor is not None:
            body_editor.setPlainText(preset.get("body", ""))

        self._update_codex_codegen_previews()
        try:
            self.connection_status_label.setText(
                f"요청 양식 적용: {preset.get('name', '')}"
            )
        except Exception:
            pass

    def _new_codex_skill_from_current_request(self) -> None:
        required = (
            "codex_skill_order_input",
            "codex_skill_name_input",
            "codex_skill_trigger_input",
            "codex_skill_body_editor",
        )
        if any(getattr(self, attr, None) is None for attr in required):
            return

        action = ""
        title = ""
        target = ""
        body = ""
        try:
            action = self.codex_request_action_combo.currentText().strip()
            title = self.codex_request_title_input.text().strip()
            target = self.codex_request_target_input.text().strip()
            body = self.codex_request_body_editor.toPlainText().strip()
        except Exception:
            pass

        skill_name = f"{title or action or 'OneNote 작업'} 반복 스킬"
        trigger = (
            f"{action or 'OneNote 작업'}을(를) {target or '현재 작업 위치'}에서 "
            "반복해서 수행할 때"
        )
        skill_body = f"""목표:
- 아래 요청 유형을 같은 기준으로 반복 수행한다.

반복 요청:
- 작업: {action or '미지정'}
- 대상: {target or '미지정'}
- 제목/이름: {title or '미지정'}

본문 기준:
{body or '- 사용자가 제공한 본문을 유지하되 OneNote에 읽기 쉬운 구조로 작성한다.'}

처리 방식:
- 작업 위치를 먼저 확인한다.
- 사용자 요청의 형식과 출력 기준을 먼저 정한다.
- 실제 OneNote 조작 방식은 코덱스 전용 지침을 따른다.

출력:
- OneNote에 반영한 항목
- 검증 결과
- 사용자가 다음에 확인할 항목
"""

        self.codex_skill_order_input.setText(self._codex_next_skill_order())
        self.codex_skill_name_input.setText(skill_name)
        self.codex_skill_trigger_input.setText(trigger)
        self.codex_skill_body_editor.setPlainText(skill_body)
        self._update_codex_skill_call_preview()
        try:
            self._scroll_codex_to_widget("codex_skill_editor_widget")
        except Exception:
            pass
        try:
            self.connection_status_label.setText("현재 요청으로 스킬 초안을 만들었습니다.")
        except Exception:
            pass

    def _build_codex_request_group(self) -> QWidget:
        group = QWidget()
        group.setObjectName("CodexCard")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(8)

        header = QLabel("✍️ 작업 내용 작성")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        form = QGridLayout()
        form.setVerticalSpacing(6)
        form.setHorizontalSpacing(6)
        form.setColumnMinimumWidth(0, 48)
        form.setColumnStretch(1, 1)

        self.codex_request_preset_combo = WheelSafeComboBox()
        self._configure_codex_lookup_combo(self.codex_request_preset_combo)
        for key, preset in self._codex_request_presets().items():
            self.codex_request_preset_combo.addItem(preset.get("name", key), key)
        form.addWidget(QLabel("양식"), 0, 0)
        preset_row = QHBoxLayout()
        preset_row.setSpacing(6)
        preset_row.addWidget(self.codex_request_preset_combo, stretch=1)
        preset_apply_btn = QToolButton()
        preset_apply_btn.setText("적용")
        preset_apply_btn.clicked.connect(self._apply_codex_request_preset)
        preset_row.addWidget(preset_apply_btn)
        form.addLayout(preset_row, 0, 1)

        self.codex_request_action_combo = WheelSafeComboBox()
        self._configure_codex_lookup_combo(self.codex_request_action_combo)
        self.codex_request_action_combo.addItems(
            ["페이지 추가", "새 섹션 생성", "섹션 그룹 생성", "기본 형식 작성"]
        )
        self.codex_request_action_combo.currentIndexChanged.connect(self._schedule_codex_codegen_previews)
        form.addWidget(QLabel("작업"), 1, 0)
        form.addWidget(self.codex_request_action_combo, 1, 1)

        self.codex_request_target_input = QLineEdit()
        self.codex_request_target_input.setPlaceholderText("전자필기장 > 섹션그룹 > 섹션")
        self.codex_request_target_input.textChanged.connect(self._schedule_codex_codegen_previews)
        form.addWidget(QLabel("경로"), 2, 0)
        form.addWidget(self.codex_request_target_input, 2, 1)

        self.codex_request_title_input = QLineEdit()
        self.codex_request_title_input.setPlaceholderText("항목의 이름을 입력하세요...")
        self.codex_request_title_input.textChanged.connect(self._schedule_codex_codegen_previews)
        form.addWidget(QLabel("제목"), 3, 0)
        form.addWidget(self.codex_request_title_input, 3, 1)

        layout.addLayout(form)

        self.codex_request_body_editor = QTextEdit()
        self.codex_request_body_editor.setPlaceholderText("구체적인 작업 지시 내용을 입력하세요. (Shift+Enter로 줄바꿈)")
        self.codex_request_body_editor.setMinimumHeight(130)
        self.codex_request_body_editor.textChanged.connect(self._schedule_codex_codegen_previews)
        layout.addWidget(self.codex_request_body_editor)

        actions = QHBoxLayout()
        actions.setSpacing(8)

        copy_btn = QToolButton()
        copy_btn.setText("🚀 작업요청 복사")
        copy_btn.setProperty("variant", "primary")
        copy_btn.setMinimumHeight(38)
        copy_btn.setMinimumWidth(0)
        copy_btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        copy_btn.clicked.connect(self._copy_codex_request_to_clipboard)

        draft_btn = QToolButton()
        draft_btn.setText("💾 초안")
        draft_btn.setMinimumHeight(38)
        draft_btn.setFixedWidth(66)
        draft_btn.clicked.connect(self._save_codex_request_draft)

        load_btn = QToolButton()
        load_btn.setText("📂 불러오기")
        load_btn.setMinimumHeight(38)
        load_btn.setFixedWidth(86)
        load_btn.clicked.connect(self._load_codex_request_draft)

        actions.addWidget(copy_btn)
        actions.addWidget(draft_btn)
        actions.addWidget(load_btn)
        layout.addLayout(actions)

        tool_grid = QGridLayout()
        tool_grid.setHorizontalSpacing(6)
        tool_grid.setVerticalSpacing(6)
        tool_specs = [
            ("📎 붙인내용 추가", self._append_clipboard_to_codex_request_body),
            ("♻️ 붙인내용 교체", self._replace_codex_request_body_from_clipboard),
            ("✔️ 실행순서 복사", self._copy_codex_execution_checklist_to_clipboard),
            ("🧩 짧은요청 복사", self._copy_codex_compact_prompt_to_clipboard),
            ("🛡️ 검토요청 복사", self._copy_codex_review_prompt_to_clipboard),
            ("🪜 단계정리 복사", self._copy_codex_task_breakdown_prompt_to_clipboard),
            ("📣 보고양식 복사", self._copy_codex_completion_report_template_to_clipboard),
            ("🎯 맞는스킬 선택", self._select_recommended_codex_skill),
            ("🧠 스킬로 만들기", self._new_codex_skill_from_current_request),
        ]
        for index, (text, cb) in enumerate(tool_specs):
            btn = QToolButton()
            btn.setText(text)
            btn.setMinimumHeight(28)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            btn.clicked.connect(cb)
            tool_grid.addWidget(btn, index // 2, index % 2)
        layout.addLayout(tool_grid)

        layout.addWidget(QLabel("생성된 요청문 미리보기"))
        self.codex_request_preview = QTextEdit()
        self.codex_request_preview.setReadOnly(True)
        self.codex_request_preview.setMinimumHeight(130)
        layout.addWidget(self.codex_request_preview)
        self._update_codex_request_preview()

        return group



    def _build_codex_template_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(8)

        header = QLabel("📄 OneNote 작업 양식")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        row = QHBoxLayout()
        self.codex_template_combo = WheelSafeComboBox()
        self.codex_template_combo.addItem("페이지 추가", "add_page")
        self.codex_template_combo.addItem("새 섹션 생성", "add_section")
        self.codex_template_combo.addItem("새 섹션 그룹 생성", "add_section_group")
        self.codex_template_combo.addItem("새 전자필기장 생성", "add_notebook")
        self.codex_template_combo.currentIndexChanged.connect(self._update_codex_template_preview)
        row.addWidget(self.codex_template_combo, stretch=1)

        copy_button = QToolButton()
        copy_button.setText("📋 양식 복사")
        copy_button.clicked.connect(self._copy_codex_template_to_clipboard)
        row.addWidget(copy_button)
        layout.addLayout(row)

        self.codex_template_preview = QTextEdit()
        self.codex_template_preview.setReadOnly(True)
        self.codex_template_preview.setMinimumHeight(140)
        layout.addWidget(self.codex_template_preview)
        self._update_codex_template_preview()

        return card


    def _build_codex_skill_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 8, 10, 10)
        layout.setSpacing(8)

        header = QLabel("스킬 관리 도구")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        desc = QLabel("주문번호표, 추천, 기본 세트처럼 가끔 쓰는 관리 작업만 모았습니다.")
        desc.setObjectName("CodexPageSubtitle")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        tools = QGridLayout()
        tools.setHorizontalSpacing(6)
        tools.setVerticalSpacing(6)
        specs = [
            ("주문표 복사", self._copy_codex_skill_order_index_to_clipboard),
            ("추천 리포트", self._copy_codex_skill_recommendation_report_to_clipboard),
            ("기본 세트", self._create_default_codex_skill_set),
            ("스킬 폴더", self._open_codex_skills_folder),
        ]
        for index, (text, cb) in enumerate(specs):
            btn = QToolButton()
            btn.setText(text)
            btn.setMinimumHeight(28)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            btn.clicked.connect(cb)
            tools.addWidget(btn, index // 3, index % 3)
        layout.addLayout(tools)
        return card


    def _build_codex_skill_editor_group(self) -> QWidget:
        root = QWidget()
        root.setObjectName("CodexSkillEditorRoot")
        root.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        main_layout = QVBoxLayout(root)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(6)

        list_panel = QWidget()
        list_panel.setObjectName("CodexCard")
        list_panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        list_layout = QVBoxLayout(list_panel)
        list_layout.setContentsMargins(8, 8, 8, 8)
        list_layout.setSpacing(4)

        list_header_row = QHBoxLayout()
        list_header_row.setSpacing(6)
        list_header = QLabel("스킬 목록")
        list_header.setObjectName("CodexCardTitle")
        list_header_row.addWidget(list_header)
        list_header_row.addStretch(1)
        list_layout.addLayout(list_header_row)

        search_row = QHBoxLayout()
        search_row.setSpacing(6)
        self.codex_skill_search_input = QLineEdit()
        self.codex_skill_search_input.setPlaceholderText("스킬 검색")
        self.codex_skill_search_input.textChanged.connect(self._schedule_filter_codex_skill_list)
        search_row.addWidget(self.codex_skill_search_input, stretch=2)

        self.codex_skill_order_lookup_input = QLineEdit()
        self.codex_skill_order_lookup_input.setPlaceholderText("주문번호 바로가기")
        self.codex_skill_order_lookup_input.returnPressed.connect(self._select_codex_skill_by_order_input)
        search_row.addWidget(self.codex_skill_order_lookup_input, stretch=1)
        list_layout.addLayout(search_row)

        self.codex_skill_list = QListWidget()
        self.codex_skill_list.itemDoubleClicked.connect(self._load_selected_codex_skill)
        self.codex_skill_list.setMinimumHeight(72)
        self.codex_skill_list.setMaximumHeight(118)
        self.codex_skill_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        list_layout.addWidget(self.codex_skill_list)

        list_actions = QHBoxLayout()
        list_actions.setSpacing(6)
        for index, (text, cb) in enumerate(
            [
                ("갱신", self._refresh_codex_skill_list),
                ("폴더", self._open_codex_skills_folder),
                ("주문표", self._copy_codex_skill_order_index_to_clipboard),
                ("추천", self._select_best_codex_skill_recommendation),
            ]
        ):
            btn = QToolButton()
            btn.setText(text)
            btn.setMinimumHeight(26)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            btn.clicked.connect(cb)
            list_actions.addWidget(btn, stretch=1)
        list_layout.addLayout(list_actions)
        main_layout.addWidget(list_panel)

        editor_area = QWidget()
        editor_area.setObjectName("CodexCard")
        editor_area.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        editor_layout = QVBoxLayout(editor_area)
        editor_layout.setContentsMargins(8, 8, 8, 8)
        editor_layout.setSpacing(6)

        header_layout = QHBoxLayout()
        header_title = QLabel("스킬 편집")
        header_title.setObjectName("CodexCardTitle")
        header_layout.addWidget(header_title)
        header_layout.addStretch()
        editor_layout.addLayout(header_layout)

        template_row = QHBoxLayout()
        template_row.setSpacing(6)
        self.codex_skill_template_combo = WheelSafeComboBox()
        for key, template in self._codex_skill_templates().items():
            self.codex_skill_template_combo.addItem(template.get("name", key), key)
        template_row.addWidget(QLabel("양식"))
        template_row.addWidget(self.codex_skill_template_combo, stretch=1)

        apply_btn = QToolButton()
        apply_btn.setText("적용")
        apply_btn.setProperty("variant", "secondary")
        apply_btn.clicked.connect(self._apply_codex_skill_template)
        template_row.addWidget(apply_btn)
        editor_layout.addLayout(template_row)

        form_layout = QGridLayout()
        form_layout.setHorizontalSpacing(6)
        form_layout.setVerticalSpacing(6)

        def add_field(row: int, col: int, label: str, field: QLineEdit, placeholder: str = "", stretch: int = 1):
            lbl = QLabel(label)
            lbl.setObjectName("CodexFieldLabel")
            field.setPlaceholderText(placeholder)
            form_layout.addWidget(lbl, row, col)
            form_layout.addWidget(field, row, col + 1)
            form_layout.setColumnStretch(col + 1, stretch)

        self.codex_skill_order_input = QLineEdit()
        self.codex_skill_order_input.textChanged.connect(self._schedule_codex_skill_call_preview)
        self.codex_skill_order_input.setMaximumWidth(92)
        add_field(0, 0, "번호", self.codex_skill_order_input, "SK-001", 0)

        self.codex_skill_name_input = QLineEdit()
        self.codex_skill_name_input.textChanged.connect(self._schedule_codex_skill_call_preview)
        add_field(0, 2, "이름", self.codex_skill_name_input, "스킬명", 2)

        self.codex_skill_trigger_input = QLineEdit()
        self.codex_skill_trigger_input.textChanged.connect(self._schedule_codex_skill_call_preview)
        trigger_label = QLabel("조건")
        trigger_label.setObjectName("CodexFieldLabel")
        self.codex_skill_trigger_input.setPlaceholderText("이 스킬이 사용될 상황")
        form_layout.addWidget(trigger_label, 1, 0)
        form_layout.addWidget(self.codex_skill_trigger_input, 1, 1, 1, 3)

        editor_layout.addLayout(form_layout)

        body_panel = QWidget()
        body_layout = QVBoxLayout(body_panel)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(6)
        body_layout.addWidget(QLabel("스킬 내용"))
        self.codex_skill_body_editor = QTextEdit()
        self.codex_skill_body_editor.setPlaceholderText("사용자 요청에 적용할 처리 기준을 작성하세요.")
        self.codex_skill_body_editor.textChanged.connect(self._schedule_codex_skill_call_preview)
        self.codex_skill_body_editor.setMinimumHeight(220)
        self.codex_skill_body_editor.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self.codex_skill_body_editor.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        body_layout.addWidget(self.codex_skill_body_editor, stretch=1)

        preview_toggle = QToolButton()
        preview_toggle.setText("스킬 요청 미리보기 펼치기")
        preview_toggle.setCheckable(True)
        preview_toggle.setMinimumHeight(28)
        body_layout.addWidget(preview_toggle)

        self.codex_skill_call_preview = QTextEdit()
        self.codex_skill_call_preview.setReadOnly(True)
        self.codex_skill_call_preview.setMinimumHeight(130)
        self.codex_skill_call_preview.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self.codex_skill_call_preview.setVisible(False)
        body_layout.addWidget(self.codex_skill_call_preview)

        def toggle_preview(checked: bool) -> None:
            self.codex_skill_call_preview.setVisible(checked)
            preview_toggle.setText(
                "스킬 요청 미리보기 접기" if checked else "스킬 요청 미리보기 펼치기"
            )

        preview_toggle.toggled.connect(toggle_preview)
        editor_layout.addWidget(body_panel, stretch=1)

        actions = QGridLayout()
        actions.setHorizontalSpacing(6)
        actions.setVerticalSpacing(6)

        save_btn = QToolButton()
        save_btn.setText("저장")
        save_btn.setProperty("variant", "primary")
        save_btn.clicked.connect(self._save_codex_skill_draft)

        new_btn = QToolButton()
        new_btn.setText("새로")
        new_btn.clicked.connect(self._new_codex_skill_draft)

        load_btn = QToolButton()
        load_btn.setText("불러")
        load_btn.clicked.connect(self._load_selected_codex_skill)

        copy_prompt_btn = QToolButton()
        copy_prompt_btn.setText("초안 복사")
        copy_prompt_btn.clicked.connect(self._copy_codex_skill_prompt_to_clipboard)

        delete_btn = QToolButton()
        delete_btn.setText("삭제")
        delete_btn.clicked.connect(self._delete_selected_codex_skill)

        more_btn = QToolButton()
        more_btn.setText("더보기")
        more_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        more_menu = QMenu(more_btn)

        def add_more_action(text: str, cb: Callable[[], None]) -> None:
            action = QAction(text, self)
            action.triggered.connect(lambda checked=False, callback=cb: callback())
            more_menu.addAction(action)

        for text, cb in [
            ("복제", self._duplicate_selected_codex_skill),
            ("클립보드에서 스킬 만들기", self._new_codex_skill_from_clipboard),
            ("현재 요청으로 스킬 만들기", self._new_codex_skill_from_current_request),
            ("선택 스킬 열기", self._open_selected_codex_skill_file),
            ("선택 경로 복사", self._copy_selected_codex_skill_path_to_clipboard),
            ("스킬 요청 복사", self._copy_codex_skill_call_prompt_to_clipboard),
            ("추천 리포트 복사", self._copy_codex_skill_recommendation_report_to_clipboard),
            ("기본 스킬 세트 생성", self._create_default_codex_skill_set),
        ]:
            add_more_action(text, cb)
        more_btn.setMenu(more_menu)

        for index, btn in enumerate(
            (save_btn, new_btn, load_btn, copy_prompt_btn, delete_btn, more_btn)
        ):
            btn.setMinimumHeight(30)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            actions.addWidget(btn, index // 3, index % 3)
        editor_layout.addLayout(actions)

        main_layout.addWidget(editor_area, stretch=1)

        self._refresh_codex_skill_list()
        self._new_codex_skill_draft()
        self._update_codex_skill_call_preview()

        return root


    def _build_codex_internal_instructions_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(8)

        header = QLabel("코덱스 전용 지침")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        help_label = QLabel(
            "OneNote COM 사용 방식, ID 우선순위, 검증 기준처럼 사용자가 매번 볼 필요 없는 실행 전제를 관리합니다."
        )
        help_label.setObjectName("CodexPageSubtitle")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        self._ensure_codex_internal_instructions_file()
        self.codex_internal_instructions_editor = QTextEdit()
        self.codex_internal_instructions_editor.setPlainText(
            self._codex_internal_instructions_text()
        )
        self.codex_internal_instructions_editor.setMinimumHeight(300)
        layout.addWidget(self.codex_internal_instructions_editor)

        actions = QHBoxLayout()
        actions.setSpacing(8)

        save_btn = QToolButton()
        save_btn.setText("저장")
        save_btn.setProperty("variant", "primary")
        save_btn.clicked.connect(self._save_codex_internal_instructions)

        reload_btn = QToolButton()
        reload_btn.setText("다시 불러오기")
        reload_btn.clicked.connect(self._reload_codex_internal_instructions)

        copy_btn = QToolButton()
        copy_btn.setText("복사")
        copy_btn.clicked.connect(self._copy_codex_internal_instructions_to_clipboard)

        open_btn = QToolButton()
        open_btn.setText("폴더 열기")
        open_btn.clicked.connect(self._open_codex_instructions_folder)

        for btn in (save_btn, reload_btn, copy_btn, open_btn):
            btn.setMinimumHeight(30)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            actions.addWidget(btn, stretch=1)

        layout.addLayout(actions)
        return card


    def _build_codex_status_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 8, 10, 10)
        layout.setSpacing(8)

        header = QLabel("📋 현재 선택값 복사")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        def _copy_row(label_text: str, value_attr: str, button_text: str, callback) -> None:
            row = QHBoxLayout()
            row.setSpacing(4)

            label = QLabel(label_text)
            label.setMinimumWidth(58)
            row.addWidget(label)

            value = QLabel("미지정")
            value.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            value.setObjectName("CodexHeroMetaValue")
            value.setWordWrap(True)
            setattr(self, value_attr, value)
            row.addWidget(value, stretch=1)

            btn = QToolButton()
            btn.setText(button_text)
            btn.setMinimumHeight(28)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            btn.clicked.connect(callback)
            row.addWidget(btn)
            layout.addLayout(row)

        _copy_row(
            "작업 위치",
            "codex_copy_target_value",
            "위치 복사",
            self._copy_codex_current_target_to_clipboard,
        )
        _copy_row(
            "선택 스킬",
            "codex_copy_skill_value",
            "스킬 복사",
            self._copy_codex_current_skill_to_clipboard,
        )
        _copy_row(
            "현재 요청",
            "codex_copy_request_value",
            "요청 복사",
            self._copy_codex_current_request_to_clipboard,
        )

        return card


    def _build_codex_quick_tools_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 8, 10, 10)
        layout.setSpacing(8)

        header = QLabel("⚡ 빠른 실행 도구")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        tools_layout = QGridLayout()
        tools_layout.setSpacing(6)

        tool_specs = [
            ("📝 주문서 복사", self._copy_codex_work_order_to_clipboard, "primary"),
            ("🚀 스킬요청 복사", self._copy_codex_skill_call_prompt_to_clipboard, "secondary"),
            ("📦 자료묶음 복사", self._copy_codex_context_pack_to_clipboard, "secondary"),
            ("📄 페이지목록 요청", self._copy_codex_page_reader_script_to_clipboard, ""),
            ("🛠️ 위치조회 요청", self._copy_codex_onenote_inventory_script_to_clipboard, ""),
        ]

        for i, (text, cb, variant) in enumerate(tool_specs):
            btn = QToolButton()
            btn.setText(text)
            btn.setMinimumHeight(28)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            if variant: btn.setProperty("variant", variant)
            btn.clicked.connect(cb)
            tools_layout.addWidget(btn, i // 2, i % 2)

        layout.addLayout(tools_layout)

        return card


    def _build_codex_context_pack_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 8, 10, 10)
        layout.setSpacing(8)

        header = QLabel("📦 작업 자료 묶음")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        self.codex_context_pack_preview = QTextEdit()
        self.codex_context_pack_preview.setReadOnly(True)
        self.codex_context_pack_preview.setMinimumHeight(96)
        self.codex_context_pack_preview.setPlaceholderText("작업 자료 묶음 미리보기가 여기에 표시됩니다.")
        layout.addWidget(self.codex_context_pack_preview)

        actions = QHBoxLayout()
        actions.setSpacing(6)
        copy_btn = QToolButton()
        copy_btn.setText("📋 자료묶음 복사")
        copy_btn.setProperty("variant", "secondary")
        copy_btn.clicked.connect(self._copy_codex_context_pack_to_clipboard)

        save_btn = QToolButton()
        save_btn.setText("💾 자료묶음 저장")
        save_btn.clicked.connect(self._save_codex_context_pack)

        refresh_btn = QToolButton()
        refresh_btn.setText("🔄 미리보기 갱신")
        refresh_btn.clicked.connect(self._update_codex_context_pack_preview)

        for btn in (copy_btn, save_btn, refresh_btn):
            btn.setMinimumWidth(0)
            btn.setMinimumHeight(28)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            actions.addWidget(btn, stretch=1)
        layout.addLayout(actions)

        return card


    def _codex_skill_package_default_codex_skills(self) -> List[str]:
        return [
            "페이지 추가",
            "전자필기장 추가",
            "전자필기장 삭제",
            "섹션 추가",
            "섹션 그룹 추가",
            "페이지 읽기",
            "위치 조회",
        ]

    def _codex_skill_package_default_instructions(self) -> List[str]:
        return [
            "OneNote COM API 우선",
            "대상 ID 우선",
            "작업별 안전 실행 순서",
            "완료 후 자동 검증",
            "실패 시 단계와 원인 보고",
        ]

    def _codex_text_lines_from_widget(self, widget: Optional[QTextEdit]) -> List[str]:
        if widget is None:
            return []
        return [
            line.strip().lstrip("-").strip()
            for line in widget.toPlainText().splitlines()
            if line.strip().lstrip("-").strip()
        ]

    def _set_codex_text_lines(self, widget: Optional[QTextEdit], lines: List[str]) -> None:
        if widget is None:
            return
        widget.setPlainText("\n".join(lines or []))

    def _codex_skill_records_from_files(self) -> List[Dict[str, str]]:
        skills_dir = self._codex_skills_dir()
        os.makedirs(skills_dir, exist_ok=True)
        rows: List[Dict[str, str]] = []
        for filename in sorted(os.listdir(skills_dir)):
            if not filename.lower().endswith(".md"):
                continue
            if filename in ("README.md", "skill-order-index.md", "skill-audit.md"):
                continue
            path = os.path.join(skills_dir, filename)
            order_no = ""
            name = filename[:-3]
            trigger = ""
            try:
                meta = self._codex_skill_metadata_from_file(path, name)
                order_no = meta.get("order", "")
                name = meta.get("name", name)
                trigger = meta.get("trigger", "")
            except Exception:
                pass
            rows.append(
                {
                    "order": order_no,
                    "name": name,
                    "filename": filename,
                    "trigger": trigger,
                    "path": path,
                }
            )
        rows.sort(key=lambda s: (s.get("order") or "ZZZ", s.get("name") or ""))
        return rows

    def _codex_checked_list_values(self, list_widget: Optional[QListWidget]) -> List[str]:
        values: List[str] = []
        if list_widget is None:
            return values
        for i in range(list_widget.count()):
            item = list_widget.item(i)
            if item.checkState() != Qt.CheckState.Checked:
                continue
            value = str(item.data(Qt.ItemDataRole.UserRole) or item.text()).strip()
            if value and value not in values:
                values.append(value)
        return values

    def _make_codex_checkable_item(
        self, label: str, value: str, checked: bool = False
    ) -> QListWidgetItem:
        item = QListWidgetItem(label)
        item.setData(Qt.ItemDataRole.UserRole, value)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        return item

    def _codex_selected_package_user_skills(self) -> List[str]:
        return self._codex_checked_list_values(
            getattr(self, "codex_skill_package_user_skill_list", None)
        )

    def _populate_codex_skill_package_user_skill_choices(
        self, checked_values: Optional[List[str]] = None
    ) -> None:
        skill_list = getattr(self, "codex_skill_package_user_skill_list", None)
        if skill_list is None:
            return

        checked: Set[str] = set()
        if checked_values is None:
            checked.update(self._codex_selected_package_user_skills())
        else:
            checked.update(
                str(value).strip() for value in checked_values if str(value).strip()
            )

        skill_list.blockSignals(True)
        try:
            skill_list.clear()
            seen: Set[str] = set()
            for skill in self._codex_skill_records_from_files():
                order_no = str(skill.get("order") or "").strip()
                name = str(skill.get("name") or "").strip()
                filename = str(skill.get("filename") or "").strip()
                value = order_no or name
                if not value:
                    continue
                label = f"[{order_no}] {name}" if order_no else name
                aliases = {
                    value,
                    order_no,
                    name,
                    filename[:-3] if filename.endswith(".md") else filename,
                }
                checked_aliases = {alias for alias in aliases if alias}
                item = self._make_codex_checkable_item(
                    label,
                    value,
                    bool(checked.intersection(checked_aliases)),
                )
                skill_list.addItem(item)
                seen.update(checked_aliases)

            for value in sorted(checked - seen):
                item = self._make_codex_checkable_item(value, value, True)
                skill_list.addItem(item)
        except Exception:
            pass
        finally:
            skill_list.blockSignals(False)

    def _set_codex_package_user_skills(self, skills: List[str]) -> None:
        self._populate_codex_skill_package_user_skill_choices(
            [str(item) for item in skills or [] if str(item).strip()]
        )

    def _codex_selected_package_codex_skills(self) -> List[str]:
        selected: List[str] = []
        skill_list = getattr(self, "codex_skill_package_codex_skill_list", None)
        if skill_list is not None:
            selected.extend(self._codex_checked_list_values(skill_list))
        for text in self._codex_text_lines_from_widget(
            getattr(self, "codex_skill_package_extra_skills_editor", None)
        ):
            if text not in selected:
                selected.append(text)
        return selected

    def _set_codex_package_codex_skills(self, skills: List[str]) -> None:
        known = set(self._codex_skill_package_default_codex_skills())
        selected = set(skills or [])
        skill_list = getattr(self, "codex_skill_package_codex_skill_list", None)
        if skill_list is not None:
            skill_list.blockSignals(True)
            for i in range(skill_list.count()):
                item = skill_list.item(i)
                value = str(item.data(Qt.ItemDataRole.UserRole) or item.text()).strip()
                item.setCheckState(
                    Qt.CheckState.Checked
                    if value in selected or item.text() in selected
                    else Qt.CheckState.Unchecked
                )
            skill_list.blockSignals(False)
        extras = [skill for skill in skills or [] if skill not in known]
        self._set_codex_text_lines(
            getattr(self, "codex_skill_package_extra_skills_editor", None),
            extras,
        )

    def _codex_skill_package_templates(self) -> Dict[str, Dict[str, Any]]:
        default_instructions = self._codex_skill_package_default_instructions()
        return {
            "quick_note": {
                "version": 1,
                "name": "기본 메모 패키지",
                "description": "사용자 글쓰기 형태에 맞춰 OneNote 페이지를 빠르게 추가하는 기본 패키지입니다.",
                "user_skills": ["SK-001"],
                "codex_skills": ["페이지 추가"],
                "instructions": default_instructions,
            },
            "work_log": {
                "version": 1,
                "name": "업무 기록 패키지",
                "description": "업무 메모를 정리하고 필요한 경우 기존 페이지를 읽어 맥락을 이어가는 패키지입니다.",
                "user_skills": ["SK-001"],
                "codex_skills": ["페이지 추가", "페이지 읽기"],
                "instructions": default_instructions
                + ["결과는 업무 기록 형식으로 정리", "다음 행동은 체크리스트로 분리"],
            },
            "notebook_admin": {
                "version": 1,
                "name": "전자필기장 관리 패키지",
                "description": "전자필기장 추가, 삭제, 위치 확인 같은 관리 작업을 안전하게 실행하는 패키지입니다.",
                "user_skills": [],
                "codex_skills": ["전자필기장 추가", "전자필기장 삭제", "위치 조회"],
                "instructions": default_instructions + ["삭제 작업은 실행 전 대상 ID를 재확인"],
            },
            "meeting_note": {
                "version": 1,
                "name": "회의 정리 패키지",
                "description": "회의 내용을 OneNote 페이지로 만들고 결정 사항과 후속 작업을 분리하는 패키지입니다.",
                "user_skills": ["SK-001"],
                "codex_skills": ["페이지 추가", "페이지 읽기"],
                "instructions": default_instructions
                + ["결정 사항과 할 일을 분리", "담당자와 기한이 있으면 본문에 유지"],
            },
        }

    def _apply_codex_skill_package_template(self) -> None:
        combo = getattr(self, "codex_skill_package_template_combo", None)
        key = str(combo.currentData() or "") if combo is not None else ""
        package = self._codex_skill_package_templates().get(key)
        if not package:
            return
        self._set_codex_skill_package_editor(dict(package))
        try:
            self.connection_status_label.setText(
                f"스킬 패키지 템플릿 적용: {package.get('name', '')}"
            )
        except Exception:
            pass

    def _default_codex_skill_package(self) -> Dict[str, Any]:
        return {
            "version": 1,
            "name": "기본 원노트 하네스",
            "description": "OneNote에 새 페이지를 만들고 사용자 스킬 형식에 맞춰 기록할 때 쓰는 기본 패키지입니다.",
            "user_skills": ["SK-001"],
            "codex_skills": ["페이지 추가"],
            "instructions": self._codex_skill_package_default_instructions(),
        }

    def _current_codex_skill_package(self) -> Dict[str, Any]:
        name_input = getattr(self, "codex_skill_package_name_input", None)
        desc_editor = getattr(self, "codex_skill_package_desc_editor", None)
        name = name_input.text().strip() if name_input is not None else ""
        desc = desc_editor.toPlainText().strip() if desc_editor is not None else ""
        return {
            "version": 1,
            "name": name or "새 스킬 패키지",
            "description": desc,
            "user_skills": self._codex_selected_package_user_skills(),
            "codex_skills": self._codex_selected_package_codex_skills(),
            "instructions": self._codex_text_lines_from_widget(
                getattr(self, "codex_skill_package_instructions_editor", None)
            ),
            "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        }

    def _set_codex_skill_package_editor(self, package: Dict[str, Any]) -> None:
        package = package or self._default_codex_skill_package()
        name_input = getattr(self, "codex_skill_package_name_input", None)
        if name_input is not None:
            name_input.setText(str(package.get("name") or "새 스킬 패키지"))
        desc_editor = getattr(self, "codex_skill_package_desc_editor", None)
        if desc_editor is not None:
            desc_editor.setPlainText(str(package.get("description") or ""))
        self._set_codex_package_user_skills(
            [str(item) for item in package.get("user_skills", [])]
        )
        self._set_codex_package_codex_skills(
            [str(item) for item in package.get("codex_skills", [])]
        )
        self._set_codex_text_lines(
            getattr(self, "codex_skill_package_instructions_editor", None),
            [str(item) for item in package.get("instructions", [])],
        )
        self._update_codex_skill_package_preview()

    def _codex_skill_package_prompt_text(
        self, package: Optional[Dict[str, Any]] = None
    ) -> str:
        package = package or self._current_codex_skill_package()
        name = str(package.get("name") or "새 스킬 패키지")
        description = str(package.get("description") or "").strip()
        user_skills = [str(item) for item in package.get("user_skills", []) if str(item).strip()]
        codex_skills = [str(item) for item in package.get("codex_skills", []) if str(item).strip()]
        instructions = [str(item) for item in package.get("instructions", []) if str(item).strip()]

        def bullet(items: List[str], fallback: str) -> str:
            return "\n".join(f"- {item}" for item in items) if items else f"- {fallback}"

        return f"""# 스킬 패키지: {name}

설명:
{description or "-"}

## 사용자 스킬

{bullet(user_skills, "사용자 스킬 미지정")}

## 코덱스 스킬

{bullet(codex_skills, "코덱스 스킬 미지정")}

## 코덱스 지침

{bullet(instructions, "코덱스 지침 미지정")}

## 사용 방식

이 스킬 패키지를 적용해서 현재 OneNote 작업 요청을 처리한다.
사용자 스킬은 결과물의 글쓰기 형태와 에이전트 역할에만 적용한다.
코덱스 스킬과 코덱스 지침은 OneNote 작업 실행 방식과 검증 기준으로 적용한다.
"""

    def _schedule_codex_skill_package_preview(self, *args) -> None:
        timer = getattr(self, "_codex_skill_package_preview_timer", None)
        if timer is None:
            timer = QTimer(self)
            timer.setSingleShot(True)
            timer.setInterval(120)
            timer.timeout.connect(self._update_codex_skill_package_preview)
            self._codex_skill_package_preview_timer = timer
        timer.start()

    def _update_codex_skill_package_preview(self) -> None:
        preview = getattr(self, "codex_skill_package_preview", None)
        if preview is None:
            return
        self._set_plain_text_if_changed(
            preview,
            self._codex_skill_package_prompt_text(),
        )

    def _selected_codex_skill_package_path(self) -> str:
        package_list = getattr(self, "codex_skill_package_list", None)
        item = package_list.currentItem() if package_list is not None else None
        return item.data(Qt.ItemDataRole.UserRole) if item is not None else ""

    def _refresh_codex_skill_package_list(self, selected_name: str = "") -> None:
        package_list = getattr(self, "codex_skill_package_list", None)
        if package_list is None:
            return
        current_name = selected_name or ""
        try:
            current_item = package_list.currentItem()
            if not current_name and current_item is not None:
                current_name = current_item.text()
        except Exception:
            pass

        package_list.blockSignals(True)
        package_list.clear()
        packages_dir = self._codex_skill_packages_dir()
        try:
            os.makedirs(packages_dir, exist_ok=True)
            rows: List[Dict[str, str]] = []
            for filename in sorted(os.listdir(packages_dir)):
                if not filename.lower().endswith(".json"):
                    continue
                path = os.path.join(packages_dir, filename)
                name = filename[:-5]
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    if isinstance(data, dict):
                        name = str(data.get("name") or name)
                except Exception:
                    pass
                rows.append({"name": name, "path": path})
            rows.sort(key=lambda row: _name_sort_key(row.get("name", "")))
            for row in rows:
                item = QListWidgetItem(row["name"])
                item.setData(Qt.ItemDataRole.UserRole, row["path"])
                package_list.addItem(item)
                if current_name and row["name"] == current_name:
                    package_list.setCurrentItem(item)
        finally:
            package_list.blockSignals(False)

    def _new_codex_skill_package(self) -> None:
        self._set_codex_skill_package_editor(self._default_codex_skill_package())
        try:
            self.connection_status_label.setText("새 스킬 패키지 초안을 만들었습니다.")
        except Exception:
            pass

    def _save_codex_skill_package(self) -> None:
        package = self._current_codex_skill_package()
        name = str(package.get("name") or "새 스킬 패키지").strip()
        packages_dir = self._codex_skill_packages_dir()
        path = os.path.join(packages_dir, self._codex_skill_slug(name) + ".json")
        try:
            self._write_json_file_atomic(path, package)
            self._refresh_codex_skill_package_list(name)
            try:
                self.connection_status_label.setText(f"스킬 패키지 저장 완료: {path}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "스킬 패키지 저장 실패", str(e))

    def _load_selected_codex_skill_package(
        self, item: Optional[QListWidgetItem] = None
    ) -> None:
        if item is not None and not isinstance(item, QListWidgetItem):
            item = None
        if item is None:
            package_list = getattr(self, "codex_skill_package_list", None)
            item = package_list.currentItem() if package_list is not None else None
        if item is None:
            QMessageBox.information(self, "안내", "먼저 스킬 패키지를 선택하세요.")
            return
        path = item.data(Qt.ItemDataRole.UserRole)
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                package = json.load(f)
            if not isinstance(package, dict):
                raise ValueError("스킬 패키지 JSON 형식이 올바르지 않습니다.")
            self._set_codex_skill_package_editor(package)
            try:
                self.connection_status_label.setText(f"스킬 패키지 불러옴: {path}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "스킬 패키지 불러오기 실패", str(e))

    def _delete_selected_codex_skill_package(self) -> None:
        path = self._selected_codex_skill_package_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 삭제할 스킬 패키지를 선택하세요.")
            return
        answer = QMessageBox.question(
            self,
            "스킬 패키지 삭제",
            f"선택한 스킬 패키지를 삭제합니다.\n\n{path}\n\n계속하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            os.remove(path)
            self._refresh_codex_skill_package_list()
            self._new_codex_skill_package()
            try:
                self.connection_status_label.setText("스킬 패키지를 삭제했습니다.")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "스킬 패키지 삭제 실패", str(e))

    def _copy_codex_skill_package_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_skill_package_prompt_text())
        try:
            self.connection_status_label.setText("스킬 패키지 호출문을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _open_codex_skill_packages_folder(self) -> None:
        try:
            os.makedirs(self._codex_skill_packages_dir(), exist_ok=True)
            os.startfile(self._codex_skill_packages_dir())
        except Exception as e:
            QMessageBox.warning(self, "스킬 패키지 폴더 열기 실패", str(e))


    def _build_codex_skill_package_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 8, 10, 10)
        layout.setSpacing(8)

        header = QLabel("스킬 패키지")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        desc = QLabel(
            "사용자 스킬, 코덱스 스킬, 실행 지침을 하나의 템플릿처럼 묶어 저장합니다. "
            "각 항목은 여러 개를 넣을 수 있습니다."
        )
        desc.setObjectName("CodexPageSubtitle")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        self.codex_skill_package_list = QListWidget()
        self.codex_skill_package_list.setMinimumHeight(72)
        self.codex_skill_package_list.setMaximumHeight(120)
        self.codex_skill_package_list.itemDoubleClicked.connect(
            self._load_selected_codex_skill_package
        )
        layout.addWidget(self.codex_skill_package_list)

        template_row = QHBoxLayout()
        template_row.setSpacing(6)
        template_row.addWidget(QLabel("패키지 템플릿"))
        self.codex_skill_package_template_combo = WheelSafeComboBox()
        self.codex_skill_package_template_combo.setMinimumContentsLength(12)
        for key, template in self._codex_skill_package_templates().items():
            self.codex_skill_package_template_combo.addItem(
                str(template.get("name") or key),
                key,
            )
        template_row.addWidget(self.codex_skill_package_template_combo, stretch=1)
        apply_template_btn = QToolButton()
        apply_template_btn.setText("템플릿 적용")
        apply_template_btn.setMinimumHeight(30)
        apply_template_btn.clicked.connect(self._apply_codex_skill_package_template)
        template_row.addWidget(apply_template_btn)
        layout.addLayout(template_row)

        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(8)
        form.setColumnStretch(1, 1)

        self.codex_skill_package_name_input = QLineEdit()
        self.codex_skill_package_name_input.setPlaceholderText("예: 기본 메모 작성 패키지")
        self.codex_skill_package_name_input.textChanged.connect(
            self._schedule_codex_skill_package_preview
        )
        form.addWidget(QLabel("이름"), 0, 0)
        form.addWidget(self.codex_skill_package_name_input, 0, 1)

        self.codex_skill_package_desc_editor = QTextEdit()
        self.codex_skill_package_desc_editor.setMinimumHeight(54)
        self.codex_skill_package_desc_editor.setPlaceholderText(
            "이 패키지를 언제 쓰는지 적어두세요."
        )
        self.codex_skill_package_desc_editor.textChanged.connect(
            self._schedule_codex_skill_package_preview
        )
        form.addWidget(QLabel("설명"), 1, 0)
        form.addWidget(self.codex_skill_package_desc_editor, 1, 1)

        self.codex_skill_package_user_skill_list = QListWidget()
        self.codex_skill_package_user_skill_list.setSelectionMode(
            QAbstractItemView.SelectionMode.NoSelection
        )
        self.codex_skill_package_user_skill_list.setMinimumHeight(110)
        self.codex_skill_package_user_skill_list.itemChanged.connect(
            self._schedule_codex_skill_package_preview
        )
        form.addWidget(QLabel("사용자 스킬"), 2, 0)
        form.addWidget(self.codex_skill_package_user_skill_list, 2, 1)

        self.codex_skill_package_codex_skill_list = QListWidget()
        self.codex_skill_package_codex_skill_list.setSelectionMode(
            QAbstractItemView.SelectionMode.NoSelection
        )
        self.codex_skill_package_codex_skill_list.setMinimumHeight(92)
        for name in self._codex_skill_package_default_codex_skills():
            self.codex_skill_package_codex_skill_list.addItem(
                self._make_codex_checkable_item(name, name)
            )
        self.codex_skill_package_codex_skill_list.itemChanged.connect(
            self._schedule_codex_skill_package_preview
        )
        form.addWidget(QLabel("코덱스 스킬"), 3, 0)
        form.addWidget(self.codex_skill_package_codex_skill_list, 3, 1)

        self.codex_skill_package_extra_skills_editor = QTextEdit()
        self.codex_skill_package_extra_skills_editor.setMinimumHeight(54)
        self.codex_skill_package_extra_skills_editor.setPlaceholderText(
            "목록에 없는 코덱스 스킬을 한 줄에 하나씩 추가"
        )
        self.codex_skill_package_extra_skills_editor.textChanged.connect(
            self._schedule_codex_skill_package_preview
        )
        form.addWidget(QLabel("추가 스킬"), 4, 0)
        form.addWidget(self.codex_skill_package_extra_skills_editor, 4, 1)

        self.codex_skill_package_instructions_editor = QTextEdit()
        self.codex_skill_package_instructions_editor.setMinimumHeight(96)
        self.codex_skill_package_instructions_editor.setPlaceholderText(
            "코덱스 지침을 한 줄에 하나씩 입력"
        )
        self.codex_skill_package_instructions_editor.textChanged.connect(
            self._schedule_codex_skill_package_preview
        )
        form.addWidget(QLabel("코덱스 지침"), 5, 0)
        form.addWidget(self.codex_skill_package_instructions_editor, 5, 1)

        layout.addLayout(form)

        self.codex_skill_package_preview = QTextEdit()
        self.codex_skill_package_preview.setReadOnly(True)
        self.codex_skill_package_preview.setMinimumHeight(150)
        self.codex_skill_package_preview.setPlaceholderText(
            "스킬 패키지 미리보기가 여기에 표시됩니다."
        )
        layout.addWidget(QLabel("패키지 호출문 미리보기"))
        layout.addWidget(self.codex_skill_package_preview)

        actions = QGridLayout()
        actions.setHorizontalSpacing(6)
        actions.setVerticalSpacing(6)
        specs = [
            ("새 패키지", self._new_codex_skill_package),
            ("저장", self._save_codex_skill_package),
            ("불러오기", self._load_selected_codex_skill_package),
            ("삭제", self._delete_selected_codex_skill_package),
            ("복사", self._copy_codex_skill_package_to_clipboard),
            ("폴더", self._open_codex_skill_packages_folder),
        ]
        for index, (text, cb) in enumerate(specs):
            btn = QToolButton()
            btn.setText(text)
            btn.setMinimumHeight(30)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            if text in ("저장", "복사"):
                btn.setProperty("variant", "primary" if text == "복사" else "secondary")
            btn.clicked.connect(cb)
            actions.addWidget(btn, index // 3, index % 3)
        layout.addLayout(actions)

        self._refresh_codex_skill_package_list()
        if self.codex_skill_package_list.count() > 0:
            self.codex_skill_package_list.setCurrentRow(0)
            self._load_selected_codex_skill_package()
        else:
            self._new_codex_skill_package()
        return card


    def _build_codex_work_order_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 8, 10, 10)
        layout.setSpacing(8)

        header = QLabel("📝 작업 주문서")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        self.codex_work_order_preview = QTextEdit()
        self.codex_work_order_preview.setReadOnly(True)
        self.codex_work_order_preview.setMinimumHeight(96)
        self.codex_work_order_preview.setPlaceholderText("생성된 주문서 미리보기가 여기에 표시됩니다.")
        layout.addWidget(self.codex_work_order_preview)

        actions = QHBoxLayout()
        actions.setSpacing(6)
        copy_btn = QToolButton()
        copy_btn.setText("🚀 주문서 복사")
        copy_btn.setProperty("variant", "primary")
        copy_btn.clicked.connect(self._copy_codex_work_order_to_clipboard)

        save_btn = QToolButton()
        save_btn.setText("💾 저장")
        save_btn.clicked.connect(self._save_codex_work_order)

        for btn in (copy_btn, save_btn):
            btn.setMinimumWidth(0)
            btn.setMinimumHeight(28)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        actions.addWidget(copy_btn, stretch=2)
        actions.addWidget(save_btn, stretch=1)
        layout.addLayout(actions)

        self._update_codex_work_order_preview()
        return card


    def _build_codex_work_order_history_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        header_row = QHBoxLayout()
        header_row.setSpacing(6)

        header = QLabel("주문서 기록")
        header.setObjectName("CodexCardTitle")
        header_row.addWidget(header)
        header_row.addStretch(1)

        refresh_btn = QToolButton()
        refresh_btn.setText("갱신")
        refresh_btn.clicked.connect(self._refresh_codex_work_order_list)
        header_row.addWidget(refresh_btn)

        folder_btn = QToolButton()
        folder_btn.setText("폴더")
        folder_btn.clicked.connect(self._open_codex_requests_folder)
        header_row.addWidget(folder_btn)
        layout.addLayout(header_row)

        search_row = QHBoxLayout()
        search_row.setSpacing(6)
        self.codex_work_order_search_input = QLineEdit()
        self.codex_work_order_search_input.setPlaceholderText("기록 검색")
        self.codex_work_order_search_input.textChanged.connect(self._schedule_refresh_codex_work_order_list)
        search_row.addWidget(self.codex_work_order_search_input, stretch=1)
        layout.addLayout(search_row)

        self.codex_work_order_list = QListWidget()
        self.codex_work_order_list.setMinimumHeight(110)
        self.codex_work_order_list.setMaximumHeight(170)
        self.codex_work_order_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.codex_work_order_list.currentItemChanged.connect(lambda current, prev=None: self._on_codex_work_order_selected(current))
        layout.addWidget(self.codex_work_order_list)

        preview_label = QLabel("선택 주문서 내용")
        preview_label.setObjectName("CodexFieldLabel")
        layout.addWidget(preview_label)

        self.codex_work_order_history_preview = QTextEdit()
        self.codex_work_order_history_preview.setReadOnly(True)
        self.codex_work_order_history_preview.setMinimumHeight(190)
        self.codex_work_order_history_preview.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        layout.addWidget(self.codex_work_order_history_preview, stretch=1)

        actions = QGridLayout()
        actions.setHorizontalSpacing(6)
        actions.setVerticalSpacing(6)

        copy_btn = QToolButton()
        copy_btn.setText("복사")
        copy_btn.setProperty("variant", "secondary")
        copy_btn.clicked.connect(self._copy_selected_codex_work_order_to_clipboard)

        load_btn = QToolButton()
        load_btn.setText("요청 불러오기")
        load_btn.clicked.connect(self._load_selected_codex_work_order_into_request)

        followup_btn = QToolButton()
        followup_btn.setText("후속 요청")
        followup_btn.clicked.connect(self._copy_selected_codex_work_order_followup_prompt)

        open_btn = QToolButton()
        open_btn.setText("열기")
        open_btn.clicked.connect(self._open_selected_codex_work_order_file)

        delete_btn = QToolButton()
        delete_btn.setText("삭제")
        delete_btn.clicked.connect(self._delete_selected_codex_work_order)

        for index, btn in enumerate((copy_btn, load_btn, followup_btn, open_btn, delete_btn)):
            btn.setMinimumHeight(28)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            actions.addWidget(btn, index // 3, index % 3)
        layout.addLayout(actions)

        self._refresh_codex_work_order_list()
        return card


    def _scroll_codex_to_widget(self, attr_name: str) -> None:
        widget = getattr(self, attr_name, None)
        if widget is None:
            return

        page_mapping = {
            "codex_status_summary_widget": ("codex_remocon_stacked_widget", "_codex_remocon_nav_buttons", 1, 0),
            "codex_quick_tools_widget": ("codex_remocon_stacked_widget", "_codex_remocon_nav_buttons", 1, 0),
            "codex_work_order_widget": ("codex_remocon_stacked_widget", "_codex_remocon_nav_buttons", 1, 0),
            "codex_request_group_widget": ("codex_remocon_stacked_widget", "_codex_remocon_nav_buttons", 1, 1),
            "codex_target_group_widget": ("codex_remocon_stacked_widget", "_codex_remocon_nav_buttons", 1, 1),
            "codex_work_order_history_widget": ("codex_remocon_stacked_widget", "_codex_remocon_nav_buttons", 1, 2),
            "codex_skill_package_widget": ("codex_harness_stacked_widget", "_codex_harness_nav_buttons", 2, 0),
            "codex_context_pack_widget": ("codex_harness_stacked_widget", "_codex_harness_nav_buttons", 2, 0),
            "codex_skill_editor_widget": ("codex_harness_stacked_widget", "_codex_harness_nav_buttons", 2, 1),
            "codex_skill_guide_widget": ("codex_harness_stacked_widget", "_codex_harness_nav_buttons", 2, 1),
            "codex_template_group_widget": ("codex_harness_stacked_widget", "_codex_harness_nav_buttons", 2, 2),
            "codex_internal_instructions_widget": ("codex_harness_stacked_widget", "_codex_harness_nav_buttons", 2, 3),
        }
        stack_attr, buttons_attr, workspace_index, idx = page_mapping.get(
            attr_name, ("", "", -1, -1)
        )
        stacked = getattr(self, stack_attr, None)
        if stacked is None:
            return

        workspace_tabs = getattr(self, "remocon_workspace_tabs", None)
        if workspace_tabs is not None and workspace_index >= 0:
            try:
                workspace_tabs.setCurrentIndex(workspace_index)
            except Exception:
                pass

        if idx >= 0:
            stacked.setCurrentIndex(idx)
            buttons = getattr(self, buttons_attr, [])
            for i, b in enumerate(buttons):
                b.setChecked(i == idx)

            try:
                scroll_area = stacked.currentWidget()
                content = scroll_area.widget()
                if content:
                    target_y = widget.mapTo(content, widget.rect().topLeft()).y()
                    scroll_area.verticalScrollBar().setValue(max(0, target_y - 16))
            except Exception:
                pass

    def _workspace_mode_from_tab_index(self, index: int) -> str:
        return "codex" if index in (1, 2) else "remocon"

    def _current_workspace_splitter_mode(self) -> str:
        tabs = getattr(self, "remocon_workspace_tabs", None)
        if tabs is None:
            return getattr(self, "_active_workspace_splitter_mode", "remocon")
        try:
            return self._workspace_mode_from_tab_index(tabs.currentIndex())
        except Exception:
            return getattr(self, "_active_workspace_splitter_mode", "remocon")

    def _splitter_total_size(self, splitter: Optional[QSplitter], fallback: int) -> int:
        if splitter is None:
            return fallback
        try:
            total = sum(max(0, int(size)) for size in splitter.sizes())
            if total > 0:
                return total
        except Exception:
            pass
        try:
            width = int(splitter.width())
            if width > 0:
                return width
        except Exception:
            pass
        return fallback

    def _capture_workspace_splitter_profile(self, mode: Optional[str] = None) -> None:
        mode = mode or self._current_workspace_splitter_mode()
        if mode not in ("remocon", "codex"):
            return
        profiles = getattr(self, "_workspace_splitter_profiles", {})
        try:
            profiles[mode] = {
                "main": list(self.main_splitter.sizes()),
                "left": list(self.left_splitter.sizes()),
                "codex": list(self.codex_splitter.sizes())
                if getattr(self, "codex_splitter", None) is not None
                else [],
            }
            self._workspace_splitter_profiles = profiles
        except Exception:
            pass

    def _restore_workspace_splitter_profile(self, mode: str) -> bool:
        profile = getattr(self, "_workspace_splitter_profiles", {}).get(mode)
        if not isinstance(profile, dict):
            return False
        try:
            main_sizes = profile.get("main") or []
            left_sizes = profile.get("left") or []
            codex_sizes = profile.get("codex") or []
            if len(main_sizes) >= 2 and sum(main_sizes) > 0:
                self.main_splitter.setSizes(main_sizes)
            if len(left_sizes) >= 2 and sum(left_sizes) > 0:
                self.left_splitter.setSizes(left_sizes)
            if (
                len(codex_sizes) >= 2
                and sum(codex_sizes) > 0
                and getattr(self, "codex_splitter", None) is not None
            ):
                self.codex_splitter.setSizes(codex_sizes)
            return bool(main_sizes or left_sizes or codex_sizes)
        except Exception:
            return False

    def _apply_workspace_splitter_preset(
        self,
        mode: Optional[str] = None,
        *,
        show_status: bool = True,
    ) -> None:
        mode = mode or self._current_workspace_splitter_mode()
        main_splitter = getattr(self, "main_splitter", None)
        left_splitter = getattr(self, "left_splitter", None)
        if main_splitter is None or left_splitter is None:
            return

        fallback_width = max(960, int(self.width()) - 24)
        total = self._splitter_total_size(main_splitter, fallback_width)
        if total <= 0:
            total = fallback_width

        if mode == "codex":
            left_width = min(max(int(total * 0.32), 320), 500)
            if total - left_width < 620:
                left_width = max(280, total - 620)
            status_name = "코덱스"
        elif mode == "balanced":
            left_width = max(300, total // 2)
            status_name = "균등"
        else:
            left_width = min(max(int(total * 0.42), 390), 610)
            if total - left_width < 420:
                left_width = max(300, total - 420)
            status_name = "위치정렬"

        right_width = max(260, total - left_width)
        main_splitter.setSizes([left_width, right_width])

        if mode == "codex":
            first_width = min(max(int(left_width * 0.42), 145), 215)
        else:
            first_width = min(max(int(left_width * 0.38), 155), 245)
        second_width = max(180, left_width - first_width)
        left_splitter.setSizes([first_width, second_width])

        codex_splitter = getattr(self, "codex_splitter", None)
        if codex_splitter is not None:
            codex_total = self._splitter_total_size(
                codex_splitter,
                max(620, right_width),
            )
            nav_width = min(max(int(codex_total * 0.20), 176), 240)
            codex_splitter.setSizes([nav_width, max(360, codex_total - nav_width)])

        capture_mode = (
            mode
            if mode in ("remocon", "codex")
            else self._current_workspace_splitter_mode()
        )
        self._capture_workspace_splitter_profile(capture_mode)
        if show_status:
            try:
                self.connection_status_label.setText(f"패널 폭 적용: {status_name}")
            except Exception:
                pass

    def _select_workspace_splitter_preset(self, mode: str) -> None:
        tabs = getattr(self, "remocon_workspace_tabs", None)
        target_index = 1 if mode == "codex" else 0
        if tabs is not None and mode in ("remocon", "codex"):
            try:
                tabs.setCurrentIndex(target_index)
            except Exception:
                pass
        self._apply_workspace_splitter_preset(mode, show_status=True)

    def _save_workspace_splitter_layout_now(self) -> None:
        self._capture_workspace_splitter_profile()
        self._save_window_state()
        try:
            self.connection_status_label.setText("현재 패널 폭을 저장했습니다.")
        except Exception:
            pass

    def _on_remocon_workspace_tab_changed(self, index: int) -> None:
        next_mode = self._workspace_mode_from_tab_index(index)
        self._active_workspace_splitter_mode = next_mode

    def _build_workspace_splitter_toolbar(self) -> QWidget:
        bar = QWidget()
        bar.setObjectName("WorkspaceSplitterToolbar")
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        label = QLabel("패널 폭")
        label.setStyleSheet("font-weight: bold; color: #D8DEE9;")
        layout.addWidget(label)

        fit_btn = QToolButton()
        fit_btn.setText("현재 탭 맞춤")
        fit_btn.clicked.connect(
            lambda: self._apply_workspace_splitter_preset(
                self._current_workspace_splitter_mode(),
                show_status=True,
            )
        )

        remocon_btn = QToolButton()
        remocon_btn.setText("위치정렬 폭")
        remocon_btn.clicked.connect(
            lambda: self._select_workspace_splitter_preset("remocon")
        )

        codex_btn = QToolButton()
        codex_btn.setText("코덱스 폭")
        codex_btn.clicked.connect(
            lambda: self._select_workspace_splitter_preset("codex")
        )

        balanced_btn = QToolButton()
        balanced_btn.setText("균등")
        balanced_btn.clicked.connect(
            lambda: self._apply_workspace_splitter_preset(
                "balanced",
                show_status=True,
            )
        )

        save_btn = QToolButton()
        save_btn.setText("현재 저장")
        save_btn.clicked.connect(self._save_workspace_splitter_layout_now)

        for btn in (fit_btn, remocon_btn, codex_btn, balanced_btn, save_btn):
            btn.setMinimumHeight(28)
            layout.addWidget(btn)

        layout.addStretch(1)
        return bar

    def _open_environment_settings_dialog(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("환경설정")
        dialog.resize(560, 160)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        panel_group = QGroupBox("패널 폭")
        panel_layout = QVBoxLayout(panel_group)
        panel_layout.setContentsMargins(12, 14, 12, 12)
        panel_layout.setSpacing(10)

        desc = QLabel("1, 2, 3번째 패널 폭을 현재 작업에 맞게 조정합니다.")
        desc.setWordWrap(True)
        panel_layout.addWidget(desc)
        panel_layout.addWidget(self._build_workspace_splitter_toolbar())
        layout.addWidget(panel_group)

        button_row = QHBoxLayout()
        button_row.addStretch(1)
        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(dialog.accept)
        button_row.addWidget(close_btn)
        layout.addLayout(button_row)

        dialog.exec()

    def _build_codex_tab(self, section: str) -> QWidget:
        root = QWidget()
        root.setObjectName("CodexRoot")
        root.setStyleSheet(
            """
            QWidget#CodexRoot {
                background-color: #111316;
                color: #E2E2E6;
                font-family: 'Segoe UI', 'Malgun Gothic';
            }
            QWidget#CodexSidePanel {
                background-color: #0C0E11;
            }
            QWidget#CodexTopBar {
                background-color: #0C0E11;
            }
            QScrollArea#CodexTopNavScroll {
                border: none;
                background-color: #0C0E11;
            }
            QLabel#CodexBrand {
                color: #E2E2E6;
                font-size: 14pt;
                font-weight: 800;
                padding: 0px;
            }
            QLabel#CodexSidebarCaption {
                color: #8D937F;
                font-size: 8pt;
                font-weight: bold;
            }
            QLabel#CodexStatusPill {
                background-color: #1E2023;
                color: #C1F56D;
                border-radius: 6px;
                padding: 3px 5px;
                font-size: 8pt;
                font-weight: bold;
            }
            QToolButton#CodexTopNavButton {
                background-color: #1E2023;
                color: #E2E2E6;
                border: 1px solid #2F3338;
                border-bottom: 2px solid transparent;
                border-radius: 4px;
                padding: 3px 3px;
                font-size: 8pt;
                font-weight: bold;
            }
            QToolButton#CodexTopNavButton:hover {
                background-color: #282A2D;
                color: #E2E2E6;
            }
            QToolButton#CodexTopNavButton:checked {
                background-color: #A6D854;
                color: #223600;
                border-color: #A6D854;
                border-bottom: 2px solid #C1F56D;
            }
            QToolButton#CodexSideNavButton {
                background-color: transparent;
                color: #C3C9B3;
                border: none;
                border-left: 3px solid transparent;
                border-radius: 0px;
                padding: 8px 8px 8px 12px;
                text-align: left;
                font-size: 10pt;
                font-weight: bold;
            }
            QToolButton#CodexSideNavButton:hover {
                background-color: #1A1C1F;
                color: #E2E2E6;
            }
            QToolButton#CodexSideNavButton:checked {
                background-color: #1A1C1F;
                color: #C1F56D;
                border-left: 3px solid #A6D854;
            }
            QStackedWidget#CodexStackedWidget,
            QWidget#CodexPageContent {
                background-color: #111316;
            }
            QScrollArea#CodexPageScroll {
                border: none;
                background-color: #111316;
            }
            QLabel#CodexPageEyebrow {
                color: #66D9CC;
                font-size: 8pt;
                font-weight: bold;
            }
            QLabel#CodexPageTitle {
                color: #E2E2E6;
                font-size: 17pt;
                font-weight: 800;
            }
            QLabel#CodexPageSubtitle {
                color: #8D937F;
                font-size: 9pt;
            }
            QWidget#CodexHeroBand {
                background-color: #1A1C1F;
                border-radius: 6px;
            }
            QLabel#CodexHeroTitle {
                color: #E2E2E6;
                font-size: 14pt;
                font-weight: 800;
            }
            QLabel#CodexHeroSubtitle,
            QLabel#CodexHeroMetaLabel {
                color: #8D937F;
                font-size: 9pt;
            }
            QLabel#CodexHeroMetaValue {
                color: #C3C9B3;
                font-size: 9pt;
                font-weight: bold;
            }
            QWidget#CodexMetricTile {
                background-color: #0C0E11;
                border-radius: 6px;
            }
            QLabel#CodexMetricLabel {
                color: #8D937F;
                font-size: 8pt;
                font-weight: bold;
            }
            QLabel#CodexMetricValue {
                color: #C1F56D;
                font-size: 14pt;
                font-weight: 800;
            }
            QLabel#CodexClusterLabel {
                color: #66D9CC;
                font-size: 8pt;
                font-weight: bold;
                padding-top: 4px;
            }
            QGroupBox {
                background-color: #1A1C1F;
                border: none;
                border-left: 4px solid #A6D854;
                border-radius: 6px;
                margin-top: 18px;
                padding: 16px 10px 10px 12px;
                font-weight: bold;
                color: #C1F56D;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 12px;
                padding: 0px 4px;
                color: #C1F56D;
                background-color: transparent;
            }
            QLabel {
                color: #C3C9B3;
            }
            QToolButton {
                background-color: #282A2D;
                color: #E2E2E6;
                border: none;
                border-radius: 6px;
                padding: 4px 6px;
                font-weight: bold;
            }
            QToolButton:hover {
                background-color: #333538;
            }
            QToolButton:pressed {
                background-color: #1E2023;
            }
            QToolButton[variant="primary"] {
                background-color: #A6D854;
                color: #223600;
            }
            QToolButton[variant="primary"]:hover {
                background-color: #C1F56D;
            }
            QLineEdit, QTextEdit, QListWidget, QComboBox {
                background-color: #0C0E11;
                color: #E2E2E6;
                border: none;
                border-radius: 6px;
                padding: 4px;
                selection-background-color: #1EA296;
                selection-color: #0C0E11;
            }
            QLineEdit:focus, QTextEdit:focus, QListWidget:focus, QComboBox:focus {
                border: 1px solid rgba(166, 216, 84, 102);
            }
            QListWidget::item {
                padding: 5px 6px;
                margin: 1px 0px;
            }
            QListWidget::item:selected {
                background-color: #282A2D;
                color: #C1F56D;
            }
            QSplitter::handle {
                background-color: #111316;
            }
            QSplitter::handle:hover {
                background-color: #1E2023;
            }
            """
        )
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        top_bar = QWidget()
        top_bar.setObjectName("CodexTopBar")
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(2, 2, 2, 2)
        top_layout.setSpacing(1)

        top_nav_scroll = QScrollArea()
        top_nav_scroll.setObjectName("CodexTopNavScroll")
        top_nav_scroll.setWidgetResizable(True)
        top_nav_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        top_nav_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        top_nav_scroll.setFixedHeight(30)

        top_nav_content = QWidget()
        top_nav_layout = QHBoxLayout(top_nav_content)
        top_nav_layout.setContentsMargins(0, 0, 0, 0)
        top_nav_layout.setSpacing(1)
        top_nav_scroll.setWidget(top_nav_content)
        top_layout.addWidget(top_nav_scroll, stretch=1)

        root_layout.addWidget(top_bar)

        stacked_widget = QStackedWidget()
        stacked_widget.setObjectName("CodexStackedWidget")
        stack_attr = (
            "codex_harness_stacked_widget"
            if section == "harness"
            else "codex_remocon_stacked_widget"
        )
        buttons_attr = (
            "_codex_harness_nav_buttons"
            if section == "harness"
            else "_codex_remocon_nav_buttons"
        )
        setattr(self, stack_attr, stacked_widget)
        setattr(self, buttons_attr, [])

        def switch_page(idx: int, btn: QToolButton):
            stacked_widget.setCurrentIndex(idx)
            for b in getattr(self, buttons_attr, []):
                b.setChecked(False)
            btn.setChecked(True)

        def add_nav_item(text: str, index: int) -> QToolButton:
            btn = QToolButton()
            btn.setObjectName("CodexTopNavButton")
            btn.setText(text)
            btn.setCheckable(True)
            btn.setMinimumHeight(26)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            btn.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
            btn.clicked.connect(lambda checked, i=index, b=btn: switch_page(i, b))
            top_nav_layout.addWidget(btn, stretch=1)
            getattr(self, buttons_attr).append(btn)
            return btn

        def make_scroll_page(
            eyebrow: str,
            title: str,
            subtitle: str,
            *,
            show_header: bool = True,
        ):
            page = QScrollArea()
            page.setObjectName("CodexPageScroll")
            page.setWidgetResizable(True)
            page.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            page.setViewportMargins(0, 0, 0, 0)
            content = QWidget()
            content.setObjectName("CodexPageContent")
            content.setMinimumWidth(0)
            content.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
            layout = QVBoxLayout(content)
            if show_header:
                layout.setContentsMargins(8, 8, 8, 10)
                layout.setSpacing(8)
            else:
                layout.setContentsMargins(4, 4, 4, 6)
                layout.setSpacing(4)

            if show_header:
                eyebrow_label = QLabel(eyebrow)
                eyebrow_label.setObjectName("CodexPageEyebrow")
                layout.addWidget(eyebrow_label)

                title_label = QLabel(title)
                title_label.setObjectName("CodexPageTitle")
                layout.addWidget(title_label)

                subtitle_label = QLabel(subtitle)
                subtitle_label.setObjectName("CodexPageSubtitle")
                subtitle_label.setWordWrap(True)
                layout.addWidget(subtitle_label)

            page.setWidget(content)
            return page, layout

        def make_metric(label: str, value: str):
            tile = QWidget()
            tile.setObjectName("CodexMetricTile")
            tile_layout = QVBoxLayout(tile)
            tile_layout.setContentsMargins(7, 5, 7, 5)
            tile_layout.setSpacing(2)

            label_widget = QLabel(label)
            label_widget.setObjectName("CodexMetricLabel")
            value_widget = QLabel(value)
            value_widget.setObjectName("CodexMetricValue")
            value_widget.setMinimumHeight(20)

            tile_layout.addWidget(label_widget)
            tile_layout.addWidget(value_widget)
            return tile, value_widget

        pages = []

        if section == "harness":
            page_package, package_layout = make_scroll_page(
                "SKILL PACKAGE",
                "스킬 패키지",
                "사용자 스킬과 코덱스 실행 자료를 하나의 요청 묶음으로 관리합니다.",
            )
            self.codex_skill_package_widget = self._build_codex_skill_package_group()
            package_layout.addWidget(self.codex_skill_package_widget)
            package_layout.addStretch(1)
            pages.append((page_package, "스킬 패키지", "스킬 패키지"))

            page_user_skills, user_skills_layout = make_scroll_page(
                "USER SKILLS",
                "사용자 스킬",
                "글쓰기 형태와 에이전트 역할처럼 결과물의 형식과 역할을 관리합니다.",
                show_header=False,
            )
            self.codex_skill_editor_widget = self._build_codex_skill_editor_group()
            self.codex_skill_guide_widget = self._build_codex_skill_group()
            user_skills_layout.addWidget(self.codex_skill_editor_widget, stretch=1)
            user_skills_layout.addWidget(self.codex_skill_guide_widget)
            user_skills_layout.addStretch(1)
            pages.append((page_user_skills, "사용자 스킬", "사용자 스킬"))

            page_codex_skills, codex_skills_layout = make_scroll_page(
                "CODEX SKILLS",
                "코덱스 스킬",
                "페이지 추가, 전자필기장 추가 같은 OneNote 실행 템플릿을 관리합니다.",
            )
            self.codex_template_group_widget = self._build_codex_template_group()
            codex_skills_layout.addWidget(self.codex_template_group_widget)
            codex_skills_layout.addStretch(1)
            pages.append((page_codex_skills, "코덱스 스킬", "코덱스 스킬"))

            page_instructions, instructions_layout = make_scroll_page(
                "CODEX INSTRUCTIONS",
                "코덱스 지침",
                "OneNote COM API 우선, 대상 ID 우선, 안전 실행 순서, 자동 검증 기준을 관리합니다.",
            )
            self.codex_internal_instructions_widget = self._build_codex_internal_instructions_group()
            instructions_layout.addWidget(self.codex_internal_instructions_widget)
            instructions_layout.addStretch(1)
            pages.append((page_instructions, "코덱스 지침", "코덱스 지침"))
        else:
            page_dashboard, dashboard_layout = make_scroll_page(
                "COMMAND HOME",
                "명령홈",
                "Codex에 보낼 작업 위치, 요청, 주문서를 한 화면에서 정리합니다.",
            )

            hero = QWidget()
            hero.setObjectName("CodexHeroBand")
            hero_layout = QVBoxLayout(hero)
            hero_layout.setContentsMargins(10, 8, 10, 8)
            hero_layout.setSpacing(6)

            hero_top = QHBoxLayout()
            hero_title_area = QVBoxLayout()
            hero_title_area.setSpacing(4)
            hero_title = QLabel("현재 작업")
            hero_title.setObjectName("CodexHeroTitle")
            hero_subtitle = QLabel("작업 위치와 요청을 바로 복사합니다.")
            hero_subtitle.setObjectName("CodexHeroSubtitle")
            hero_subtitle.setWordWrap(True)
            hero_title_area.addWidget(hero_title)
            hero_title_area.addWidget(hero_subtitle)
            hero_top.addLayout(hero_title_area, stretch=1)

            self.codex_hero_request_value = QLabel("요청 대기 중")
            self.codex_hero_request_value.setObjectName("CodexStatusPill")
            self.codex_hero_request_value.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.codex_hero_request_value.setMinimumWidth(72)
            self.codex_hero_request_value.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            hero_top.addWidget(self.codex_hero_request_value)
            hero_layout.addLayout(hero_top)

            target_row = QHBoxLayout()
            target_row.setSpacing(4)
            target_label = QLabel("대상")
            target_label.setObjectName("CodexHeroMetaLabel")
            target_row.addWidget(target_label)
            self.codex_hero_target_value = QLabel("작업 위치 미선택")
            self.codex_hero_target_value.setObjectName("CodexHeroMetaValue")
            self.codex_hero_target_value.setWordWrap(True)
            target_row.addWidget(self.codex_hero_target_value, stretch=1)
            hero_layout.addLayout(target_row)

            metrics = QHBoxLayout()
            metrics.setSpacing(4)
            target_tile, self.codex_metric_target_value = make_metric("작업 위치", "0개")
            draft_tile, self.codex_metric_draft_value = make_metric("요청 초안", "대기")
            skill_tile, self.codex_metric_skill_value = make_metric("스킬 수", "0개")
            order_tile, self.codex_metric_order_value = make_metric("작업 주문서", "0개")
            for tile in (target_tile, draft_tile, skill_tile, order_tile):
                metrics.addWidget(tile)
            hero_layout.addLayout(metrics)
            dashboard_layout.addWidget(hero)

            self.codex_status_summary_widget = self._build_codex_status_group()
            dashboard_layout.addWidget(self.codex_status_summary_widget)
            self.codex_quick_tools_widget = self._build_codex_quick_tools_group()
            dashboard_layout.addWidget(self.codex_quick_tools_widget)
            self.codex_work_order_widget = self._build_codex_work_order_group()
            dashboard_layout.addWidget(self.codex_work_order_widget)
            dashboard_layout.addStretch(1)
            pages.append((page_dashboard, "명령홈", "명령홈"))

            page_request, request_layout = make_scroll_page(
                "COMPOSITION",
                "작업요청",
                "위에서 작업 위치를 정하고 아래에서 지시 내용을 작성하세요.",
            )

            self.codex_target_group_widget = self._build_codex_target_group()
            self.codex_request_group_widget = self._build_codex_request_group()

            request_layout.addWidget(self.codex_target_group_widget)
            request_layout.addWidget(self.codex_request_group_widget)
            request_layout.addStretch(1)
            self._apply_codex_target_to_request()
            self._update_codex_request_preview()
            pages.append((page_request, "작업요청", "작업 요청"))

            page_history, history_layout = make_scroll_page(
                "HISTORY",
                "기록",
                "저장된 작업 주문서 기록을 검색하고 다시 불러옵니다.",
                show_header=False,
            )

            self.codex_work_order_history_widget = self._build_codex_work_order_history_group()
            history_layout.addWidget(self.codex_work_order_history_widget, stretch=1)
            history_layout.addStretch(1)
            pages.append((page_history, "기록", "작업 기록"))

        nav_buttons = []
        for index, (page, text, tooltip) in enumerate(pages):
            stacked_widget.addWidget(page)
            btn = add_nav_item(text, index)
            btn.setToolTip(tooltip)
            nav_buttons.append(btn)

        self.codex_splitter = None
        root_layout.addWidget(stacked_widget, stretch=1)

        self._update_codex_work_order_preview()
        self._update_codex_context_pack_preview()
        self._update_codex_status_summary()
        if nav_buttons:
            switch_page(0, nav_buttons[0])

        return root

    def _show_onenote_harness_help(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("원노트 하네스 도움말")
        dialog.resize(780, 680)
        dialog.setStyleSheet(
            """
            QDialog {
                background-color: #111316;
                color: #E2E2E6;
                font-family: 'Malgun Gothic', 'Segoe UI';
            }
            QTextEdit {
                background-color: #0C0E11;
                border: 1px solid #2F3338;
                border-radius: 8px;
                padding: 0px;
            }
            QPushButton {
                background-color: #A6D854;
                color: #223600;
                border: none;
                border-radius: 6px;
                padding: 8px 18px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #C1F56D;
            }
            QPushButton:pressed {
                background-color: #95C743;
            }
            """
        )

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        help_view = QTextEdit(dialog)
        help_view.setReadOnly(True)
        help_view.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        help_view.setHtml(
            """
            <html>
            <head>
            <style>
                body {
                    margin: 0;
                    padding: 0;
                    background: #0C0E11;
                    color: #E2E2E6;
                    font-family: 'Malgun Gothic', 'Segoe UI', sans-serif;
                    font-size: 13px;
                    line-height: 1.55;
                }
                .wrap {
                    padding: 22px;
                }
                .hero {
                    background: #1A1C1F;
                    border: 1px solid #2F3338;
                    border-radius: 8px;
                    padding: 20px;
                    margin-bottom: 16px;
                }
                .eyebrow {
                    color: #66D9CC;
                    font-size: 11px;
                    font-weight: 800;
                    margin-bottom: 6px;
                }
                h1 {
                    color: #E2E2E6;
                    font-size: 26px;
                    margin: 0 0 8px 0;
                }
                .lead {
                    color: #C3C9B3;
                    margin: 0;
                }
                .section {
                    background: #15171A;
                    border-left: 4px solid #A6D854;
                    border-radius: 8px;
                    padding: 16px;
                    margin: 12px 0;
                }
                h2 {
                    color: #C1F56D;
                    font-size: 18px;
                    margin: 0 0 8px 0;
                }
                p {
                    margin: 0 0 10px 0;
                    color: #C3C9B3;
                }
                .tags {
                    margin-top: 8px;
                }
                .tag {
                    display: inline-block;
                    background: #282A2D;
                    color: #E2E2E6;
                    border: 1px solid #3B4046;
                    border-radius: 6px;
                    padding: 5px 8px;
                    margin: 3px 4px 3px 0;
                    font-weight: 700;
                }
                .tag-accent {
                    background: #223600;
                    color: #C1F56D;
                    border-color: #A6D854;
                }
                .flow {
                    margin: 0;
                    padding-left: 20px;
                    color: #C3C9B3;
                }
                .flow li {
                    margin: 7px 0;
                }
                .note {
                    background: #0C0E11;
                    border: 1px solid #2F3338;
                    border-radius: 8px;
                    padding: 12px;
                    color: #8D937F;
                    margin-top: 12px;
                }
                code {
                    color: #66D9CC;
                    background: #111316;
                    padding: 2px 5px;
                    border-radius: 4px;
                }
            </style>
            </head>
            <body>
            <div class="wrap">
                <div class="hero">
                    <div class="eyebrow">ONENOTE HARNESS</div>
                    <h1>원노트 하네스</h1>
                    <p class="lead">
                        사용자 스킬, 코덱스 스킬, 실행 지침을 하나의 패키지로 묶어
                        OneNote 작업 요청을 빠르게 만들고 안전하게 검증하는 작업 공간입니다.
                    </p>
                </div>

                <div class="section">
                    <h2>스킬 패키지</h2>
                    <p>
                        사용자 스킬들을 조합한 템플릿입니다. 자주 쓰는 글쓰기 방식,
                        에이전트 역할, 실행 작업을 하나의 구성으로 묶어 Codex 요청에 바로 붙일 수 있습니다.
                    </p>
                    <div class="tags">
                        <span class="tag tag-accent">템플릿</span>
                        <span class="tag">여러 사용자 스킬</span>
                        <span class="tag">여러 코덱스 스킬</span>
                        <span class="tag">지침 묶음</span>
                    </div>
                </div>

                <div class="section">
                    <h2>사용자 스킬</h2>
                    <p>결과물의 형태와 에이전트 역할을 정합니다.</p>
                    <div class="tags">
                        <span class="tag">글쓰기 형태</span>
                        <span class="tag">에이전트 역할</span>
                    </div>
                </div>

                <div class="section">
                    <h2>코덱스 스킬</h2>
                    <p>Codex가 실제로 수행할 OneNote 작업입니다.</p>
                    <div class="tags">
                        <span class="tag">페이지 추가</span>
                        <span class="tag">전자필기장 추가</span>
                        <span class="tag">전자필기장 삭제</span>
                    </div>
                </div>

                <div class="section">
                    <h2>코덱스 지침</h2>
                    <p>Codex가 OneNote 작업을 안전하게 실행하기 위한 내부 실행 기준입니다.</p>
                    <div class="tags">
                        <span class="tag tag-accent">OneNote COM API 우선</span>
                        <span class="tag">대상 ID 우선</span>
                        <span class="tag">작업별 안전 실행 순서</span>
                        <span class="tag">완료 후 자동 검증</span>
                        <span class="tag">실패 시 단계와 원인 보고</span>
                    </div>
                </div>

                <div class="section">
                    <h2>사용 흐름</h2>
                    <ol class="flow">
                        <li>사용할 사용자 스킬을 고릅니다.</li>
                        <li>실행할 코덱스 스킬을 고릅니다.</li>
                        <li>대상 위치, 제목, 본문 같은 요청 내용을 입력합니다.</li>
                        <li>Codex는 코덱스 지침을 기준으로 실행하고 검증합니다.</li>
                    </ol>
                    <div class="note">
                        저장된 스킬 패키지는 <code>docs/codex/skill-packages</code>에 JSON으로 보관됩니다.
                    </div>
                </div>
            </div>
            </body>
            </html>
            """
        )
        layout.addWidget(help_view)

        close_btn = QPushButton("닫기", dialog)
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignRight)

        dialog.exec()

    def init_ui(self, initial_status):
        self.setWindowTitle("OneNote 전자필기장 위치정렬")

        # --- 메뉴바 생성 ---
        menubar = self.menuBar()
        file_menu = menubar.addMenu("&파일")

        backup_action = QAction("백업하기...", self)
        backup_action.triggered.connect(self._backup_full_settings)
        file_menu.addAction(backup_action)

        restore_action = QAction("복원하기...", self)
        restore_action.triggered.connect(self._restore_full_settings)
        file_menu.addAction(restore_action)

        file_menu.addSeparator()

        shared_settings_action = QAction("공용 설정 JSON 위치 지정...", self)
        shared_settings_action.triggered.connect(self._choose_shared_settings_json)
        file_menu.addAction(shared_settings_action)

        show_settings_path_action = QAction("현재 설정 JSON 위치 보기", self)
        show_settings_path_action.triggered.connect(self._show_settings_json_path)
        file_menu.addAction(show_settings_path_action)

        open_settings_folder_action = QAction("설정 JSON 폴더 열기", self)
        open_settings_folder_action.triggered.connect(self._open_settings_json_folder)
        file_menu.addAction(open_settings_folder_action)

        clear_shared_settings_action = QAction("공용 설정 JSON 연결 해제", self)
        clear_shared_settings_action.triggered.connect(self._clear_shared_settings_json)
        file_menu.addAction(clear_shared_settings_action)

        settings_menu = menubar.addMenu("&환경설정")
        panel_width_action = QAction("패널 폭 조정...", self)
        panel_width_action.triggered.connect(self._open_environment_settings_dialog)
        settings_menu.addAction(panel_width_action)

        special_menu = menubar.addMenu("&특수 기능")
        self.open_all_notebooks_action = QAction(
            "실제 OneNote 전자필기장 모두 열기", self
        )
        self.open_all_notebooks_action.setStatusTip(
            "OneNote의 '전자 필기장 열기' 화면에서 아직 안 열린 전자필기장을 순서대로 엽니다."
        )
        self.open_all_notebooks_action.triggered.connect(
            self._open_all_notebooks_from_connected_onenote
        )
        self.open_all_notebooks_action.setEnabled(False)
        special_menu.addAction(self.open_all_notebooks_action)

        help_menu = menubar.addMenu("&도움말")
        onenote_harness_help_action = QAction("원노트 하네스 도움말", self)
        onenote_harness_help_action.triggered.connect(self._show_onenote_harness_help)
        help_menu.addAction(onenote_harness_help_action)

        # --- 스타일시트 정의 (생략) ---
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
        COLOR_STATUS_BAR = "#252525"

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
                padding: 4px 6px;
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
            QMenuBar {{
                background-color: {COLOR_GROUPBOX_BG};
                color: {COLOR_PRIMARY_TEXT};
            }}
            QMenuBar::item:selected {{
                background-color: {COLOR_SECONDARY_BUTTON_HOVER};
            }}
            QMenu {{
                background-color: {COLOR_GROUPBOX_BG};
                border: 1px solid {COLOR_SECONDARY_BUTTON};
            }}
            QMenu::item:selected {{
                background-color: {COLOR_LIST_SELECTED};
            }}
            #StatusBarLabel {{
                background-color: {COLOR_STATUS_BAR};
                color: {COLOR_PRIMARY_TEXT};
                padding: 5px 12px;
                font-size: 9pt;
                border-top: 1px solid #444444;
            }}
            QLineEdit {{
                background-color: {COLOR_LIST_BG};
                border: 1px solid {COLOR_SECONDARY_BUTTON};
                border-radius: 4px;
                padding: 4px 8px;
            }}
            QLineEdit:focus {{
                border: 1px solid {COLOR_LIST_SELECTED};
            }}
            QTextEdit {{
                background-color: {COLOR_LIST_BG};
                border: 1px solid {COLOR_GROUPBOX_BG};
                border-radius: 8px;
                color: {COLOR_PRIMARY_TEXT};
                padding: 10px;
            }}
            QTabWidget::pane {{
                border: 1px solid {COLOR_GROUPBOX_BG};
                border-radius: 6px;
            }}
            QTabBar::tab {{
                background-color: {COLOR_SECONDARY_BUTTON};
                color: {COLOR_PRIMARY_TEXT};
                padding: 7px 14px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                margin-right: 2px;
            }}
            QTabBar::tab:selected {{
                background-color: {COLOR_ACCENT};
                color: #111;
                font-weight: bold;
            }}
            QScrollArea {{
                border: none;
                background-color: transparent;
            }}
            QScrollArea > QWidget > QWidget {{
                background-color: {COLOR_BACKGROUND};
            }}
            #CodexSidePanel {{
                background-color: #242424;
                border-right: 1px solid #444444;
            }}
            #CodexSideSectionLabel {{
                color: {COLOR_SECONDARY_TEXT};
                font-size: 8pt;
                padding: 8px 8px 2px 8px;
            }}
            #CodexSideNavButton {{
                background-color: transparent;
                color: {COLOR_PRIMARY_TEXT};
                border-radius: 16px;
                padding: 7px 12px;
                text-align: left;
            }}
            #CodexSideNavButton:hover {{
                background-color: {COLOR_SECONDARY_BUTTON_HOVER};
            }}
            #CodexSideNavButton:pressed {{
                background-color: {COLOR_SECONDARY_BUTTON_PRESSED};
            }}
        """
        )

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(10)

        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)  # self로 저장
        self.main_splitter.setChildrenCollapsible(False)
        main_layout.addWidget(self.main_splitter, stretch=1)

        self.left_splitter = QSplitter(Qt.Orientation.Horizontal)  # self로 저장
        self.left_splitter.setChildrenCollapsible(False)

        # 1. 즐겨찾기 버퍼 관리 패널 (가장 왼쪽)
        buffer_panel = QWidget()
        buffer_layout = QVBoxLayout(buffer_panel)
        buffer_layout.setContentsMargins(0, 0, 0, 0)
        buffer_layout.setSpacing(8)

        buffer_group = QGroupBox("프로젝트/등록 영역")
        buffer_group_layout = QVBoxLayout(buffer_group)

        # 즐겨찾기 버퍼 상단 툴바: 추가, 이름변경
        buffer_toolbar_top_layout = QHBoxLayout()
        self.btn_add_buffer_group = QToolButton()
        self.btn_add_buffer_group.setText("그룹")
        self.btn_add_buffer_group.clicked.connect(self._add_buffer_group)

        self.btn_add_buffer = QToolButton()
        self.btn_add_buffer.setText("버퍼")
        self.btn_add_buffer.clicked.connect(self._add_buffer)

        self.btn_rename_buffer = QToolButton()
        self.btn_rename_buffer.setText("이름변경")
        self.btn_rename_buffer.clicked.connect(self._rename_buffer)

        self.btn_register_all_notebooks = QToolButton()
        self.btn_register_all_notebooks.setText("종합 새로고침")
        self.btn_register_all_notebooks.setToolTip(
            "현재 열린 OneNote 전자필기장 목록을 다시 읽고 미분류/분류 상태를 한 번에 갱신합니다."
        )
        self.btn_register_all_notebooks.clicked.connect(self._register_all_notebooks_from_current_onenote)
        self.btn_register_all_notebooks.setEnabled(False)  # 종합 버퍼에서만 활성화
        self.btn_register_all_notebooks.setVisible(False)  # 종합 버퍼에서만 표시

        buffer_toolbar_top_layout.addWidget(self.btn_add_buffer_group)
        buffer_toolbar_top_layout.addWidget(self.btn_add_buffer)
        buffer_toolbar_top_layout.addWidget(self.btn_rename_buffer)
        buffer_toolbar_top_layout.addStretch(1)
        buffer_group_layout.addLayout(buffer_toolbar_top_layout)

        # QListWidget -> BufferTree로 교체
        self.buffer_tree = BufferTree()
        self._tree_name_edit_delegate = TreeNameEditDelegate(self)
        self.buffer_tree.setItemDelegate(self._tree_name_edit_delegate)
        # PERF: 큰 트리에서 초기 렌더링 성능 개선
        try:
            self.buffer_tree.setUniformRowHeights(True)
            self.buffer_tree.setAnimated(False)
        except Exception:
            pass

        self.buffer_tree.itemClicked.connect(self._on_buffer_tree_item_clicked)
        self.buffer_tree.itemDoubleClicked.connect(self._on_buffer_tree_double_clicked)
        # ✅ 2패널(모듈/섹션)이 즉시 갱신되도록 하되, 부팅 중에는 무시
        self.buffer_tree.itemSelectionChanged.connect(
            lambda: getattr(self, "_boot_loading", False) or self._on_buffer_tree_selection_changed()
        )
        self.buffer_tree.structureChanged.connect(self._request_buffer_structure_save)
        self.buffer_tree.renameRequested.connect(self._rename_buffer)
        self.buffer_tree.deleteRequested.connect(self._delete_buffer)

        self.buffer_tree.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self.buffer_tree.customContextMenuRequested.connect(
            self._on_buffer_context_menu
        )

        buffer_group_layout.addWidget(self.buffer_tree)

        # 즐겨찾기 버퍼 하단 툴바: 삭제, 위로, 아래로
        buffer_toolbar_bottom_layout = QHBoxLayout()
        self.btn_delete_buffer = QToolButton()
        self.btn_delete_buffer.setText("삭제")
        self.btn_delete_buffer.setFixedWidth(52)
        self.btn_delete_buffer.clicked.connect(self._delete_buffer)

        self.btn_buffer_move_up = QToolButton()
        self.btn_buffer_move_up.setText("▲")
        self.btn_buffer_move_up.setToolTip("위로")
        self.btn_buffer_move_up.setFixedWidth(32)
        self.btn_buffer_move_up.clicked.connect(self._move_buffer_up)

        self.btn_buffer_move_down = QToolButton()
        self.btn_buffer_move_down.setText("▼")
        self.btn_buffer_move_down.setToolTip("아래로")
        self.btn_buffer_move_down.setFixedWidth(32)
        self.btn_buffer_move_down.clicked.connect(self._move_buffer_down)

        buffer_toolbar_bottom_layout.addWidget(self.btn_delete_buffer)
        buffer_toolbar_bottom_layout.addStretch(1)
        buffer_toolbar_bottom_layout.addWidget(self.btn_buffer_move_up)
        buffer_toolbar_bottom_layout.addWidget(self.btn_buffer_move_down)
        buffer_group_layout.addLayout(buffer_toolbar_bottom_layout)

        buffer_layout.addWidget(buffer_group)
        self.left_splitter.addWidget(buffer_panel)

        # 2. 즐겨찾기 관리 패널 (중앙)
        favorites_panel = QWidget()
        left_layout = QVBoxLayout(favorites_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        fav_group = QGroupBox("모듈영역")
        fav_layout = QVBoxLayout(fav_group)

        # 툴바 - 1행: 그룹추가, 현재 전자필기장 추가, 이름 바꾸기
        tb1_layout = QHBoxLayout()
        self.btn_add_group = QToolButton()
        self.btn_add_group.setText("그룹 추가")
        self.btn_add_group.clicked.connect(self._add_group)
        self.btn_add_section_current = QToolButton()
        self.btn_add_section_current.setText("현재 전자필기장 추가")
        self.btn_add_section_current.clicked.connect(self._add_section_from_current)
        self.btn_rename = QToolButton()
        self.btn_rename.setText("이름바꾸기")
        self.btn_rename.clicked.connect(self._rename_favorite_item)
        tb1_layout.addWidget(self.btn_add_section_current)
        tb1_layout.addWidget(self.btn_rename)
        tb1_layout.addStretch(1)

        # 툴바 - 2행: 그룹 추가, 그룹 펼치기/접기 드롭다운
        tb2_layout = QHBoxLayout()
        self.btn_group_expand_collapse = QToolButton()
        self.btn_group_expand_collapse.setText("그룹 펼치기/접기")
        self.btn_group_expand_collapse.setToolTip("그룹 펼치기 또는 그룹 접기를 선택합니다.")
        self.btn_group_expand_collapse.setPopupMode(
            QToolButton.ToolButtonPopupMode.InstantPopup
        )
        self.menu_group_expand_collapse = QMenu(self.btn_group_expand_collapse)
        self.action_expand_all_groups = QAction("그룹 펼치기", self)
        self.action_collapse_all_groups = QAction("그룹 접기", self)
        self.menu_group_expand_collapse.addAction(self.action_expand_all_groups)
        self.menu_group_expand_collapse.addAction(self.action_collapse_all_groups)
        self.btn_group_expand_collapse.setMenu(self.menu_group_expand_collapse)
        tb2_layout.addWidget(self.btn_add_group)
        tb2_layout.addStretch(1)
        tb2_layout.addWidget(self.btn_group_expand_collapse)

        tb3_layout = QHBoxLayout()
        tb3_layout.addStretch(1)
        tb3_layout.addWidget(self.btn_register_all_notebooks)
        tb3_layout.addStretch(1)

        fav_layout.addLayout(tb1_layout)
        fav_layout.addLayout(tb2_layout)
        fav_layout.addLayout(tb3_layout)

        self.fav_tree = FavoritesTree()
        self.fav_tree.setItemDelegate(self._tree_name_edit_delegate)
        # PERF: 큰 트리에서 초기 렌더링 성능 개선
        try:
            self.fav_tree.setUniformRowHeights(True)
            self.fav_tree.setAnimated(False)
        except Exception:
            pass

        # PERF: 아이콘 캐싱(standardIcon 반복 호출 비용 감소)
        try:
            self._icon_file = self.style().standardIcon(QApplication.style().StandardPixmap.SP_FileIcon)
            self._icon_dir = self.style().standardIcon(QApplication.style().StandardPixmap.SP_DirIcon)
            self._icon_agg = self.style().standardIcon(QApplication.style().StandardPixmap.SP_ComputerIcon)
        except Exception:
            self._icon_file = None
            self._icon_dir = None
            self._icon_agg = None

        self.action_expand_all_groups.triggered.connect(
            lambda: self._expand_fav_groups_always(reason="toolbar")
        )
        self.action_collapse_all_groups.triggered.connect(
            lambda: self._collapse_fav_groups_always(reason="toolbar")
        )
        self.fav_tree.itemDoubleClicked.connect(self._on_fav_item_double_clicked)
        self.fav_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.fav_tree.customContextMenuRequested.connect(self._on_fav_context_menu)
        self.fav_tree.structureChanged.connect(self._request_favorites_save)
        self.fav_tree.itemChanged.connect(self._request_favorites_save)
        # ✅ 키보드 매핑 (Del, F2, Ctrl+C/V) 직접 연결
        self.fav_tree.deleteRequested.connect(self._delete_favorite_item)
        self.fav_tree.renameRequested.connect(self._rename_favorite_item)
        self.fav_tree.copyRequested.connect(self._copy_favorite_item)
        self.fav_tree.pasteRequested.connect(self._paste_favorite_item)
        self.fav_tree.cutRequested.connect(self._cut_favorite_item)
        self.fav_tree.undoRequested.connect(self._undo_favorite_tree)
        self.fav_tree.redoRequested.connect(self._redo_favorite_tree)
        fav_layout.addWidget(self.fav_tree)

        # 즐겨찾기 하단 툴바: 삭제, 위로, 아래로 (삭제 버튼 재배치)
        move_buttons_layout = QHBoxLayout()

        # 삭제 버튼 (tb2에서 이동)
        self.btn_delete = QToolButton()
        self.btn_delete.setText("삭제")
        self.btn_delete.setFixedWidth(52)
        self.btn_delete.clicked.connect(self._delete_favorite_item)
        move_buttons_layout.addWidget(self.btn_delete)

        move_buttons_layout.addStretch(1)

        self.btn_move_up = QToolButton()
        self.btn_move_up.setText("▲")
        self.btn_move_up.setToolTip("위로")
        self.btn_move_up.setFixedWidth(32)
        self.btn_move_up.clicked.connect(self._move_item_up)
        self.btn_move_down = QToolButton()
        self.btn_move_down.setText("▼")
        self.btn_move_down.setToolTip("아래로")
        self.btn_move_down.setFixedWidth(32)
        self.btn_move_down.clicked.connect(self._move_item_down)
        move_buttons_layout.addWidget(self.btn_move_up)
        move_buttons_layout.addWidget(self.btn_move_down)
        fav_layout.addLayout(move_buttons_layout)

        self.fav_tree.itemSelectionChanged.connect(self._update_move_button_state)
        self.fav_tree.itemSelectionChanged.connect(
            self._sync_codex_target_from_current_fav_item
        )
        left_layout.addWidget(fav_group, stretch=1)

        self.left_splitter.addWidget(favorites_panel)
        self.main_splitter.addWidget(self.left_splitter)

        # 3. 오른쪽 패널: 위치정렬/코덱스 탭만 교체되고 1, 2패널은 고정된다.
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

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
        self.refresh_button.clicked.connect(self.refresh_onenote_list)
        list_header_layout.addWidget(self.refresh_button)

        connection_layout.addLayout(list_header_layout)

        self.onenote_list_widget = QListWidget()
        self.onenote_list_widget.addItem("자동 재연결 시도 중...")
        self.onenote_list_widget.installEventFilter(self)
        self.onenote_list_widget.viewport().installEventFilter(self)
        self.onenote_list_widget.itemDoubleClicked.connect(
            self.connect_and_center_from_list_item
        )
        connection_layout.addWidget(self.onenote_list_widget)
        right_layout.addWidget(connection_group)

        actions_group = QGroupBox("현재 열린 항목 제어")
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

        search_group = QGroupBox("검색 / 위치정렬")
        search_group_layout = QVBoxLayout(search_group)
        search_group_layout.setSpacing(8)

        project_search_label = QLabel("프로젝트 검색")
        project_search_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        search_group_layout.addWidget(project_search_label)

        module_search_layout = QHBoxLayout()
        self.module_project_search_input = QLineEdit()
        self.module_project_search_input.setPlaceholderText(
            "프로젝트/등록영역 + 모듈영역 검색 (띄어쓰기 무시)..."
        )
        self.module_project_search_input.setClearButtonEnabled(True)
        self.module_project_search_input.textChanged.connect(
            self._schedule_project_buffer_search_highlight
        )
        self.btn_module_project_search_clear = QToolButton()
        self.btn_module_project_search_clear.setText("검색 지우기")
        self.btn_module_project_search_clear.clicked.connect(
            self.module_project_search_input.clear
        )
        module_search_layout.addStretch(1)
        module_search_layout.addWidget(self.module_project_search_input, stretch=4)
        module_search_layout.addWidget(self.btn_module_project_search_clear)
        module_search_layout.addStretch(1)
        search_group_layout.addLayout(module_search_layout)

        project_search_hint = QLabel(
            "입력한 글자가 포함된 항목은 프로젝트/등록영역과 모듈영역에 하이라이트로 표시됩니다."
        )
        project_search_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        project_search_hint.setWordWrap(True)
        project_search_hint.setStyleSheet("color: #B8B8B8; font-size: 9pt;")
        search_group_layout.addWidget(project_search_hint)
        right_layout.addWidget(search_group)

        right_layout.addStretch(1)

        workspace_panel = QWidget()
        workspace_layout = QVBoxLayout(workspace_panel)
        workspace_layout.setContentsMargins(0, 0, 0, 0)
        workspace_layout.setSpacing(0)

        self._workspace_splitter_profiles = {}
        self._active_workspace_splitter_mode = "remocon"
        self.remocon_workspace_tabs = QTabWidget()
        self.remocon_workspace_tabs.setObjectName("RemoconWorkspaceTabs")
        self.remocon_workspace_tabs.addTab(right_panel, "위치정렬")
        self.remocon_workspace_tabs.addTab(self._build_codex_tab("remocon"), "원노트 리모컨")
        self.remocon_workspace_tabs.addTab(self._build_codex_tab("harness"), "원노트 하네스")
        self.remocon_workspace_tabs.currentChanged.connect(
            self._on_remocon_workspace_tab_changed
        )
        workspace_layout.addWidget(self.remocon_workspace_tabs, stretch=1)
        self.main_splitter.addWidget(workspace_panel)

        self.connection_status_label = QLabel(initial_status)
        self.statusBar().addPermanentWidget(self.connection_status_label)
        self.statusBar().setStyleSheet(f"background-color: {COLOR_STATUS_BAR};")

        # --- [START] 스플리터 상태 복원 로직 (수정됨) ---
        # 저장된 스플리터 상태 불러오기
        splitter_states = self.settings.get("splitter_states")
        restored = False
        codex_restored = False
        if isinstance(splitter_states, dict):
            try:
                main_state_b64 = splitter_states.get("main")
                if main_state_b64:
                    main_state = QByteArray.fromBase64(main_state_b64.encode("ascii"))
                    if not main_state.isEmpty():
                        self.main_splitter.restoreState(main_state)

                left_state_b64 = splitter_states.get("left")
                if left_state_b64:
                    left_state = QByteArray.fromBase64(left_state_b64.encode("ascii"))
                    if not left_state.isEmpty():
                        self.left_splitter.restoreState(left_state)

                codex_state_b64 = splitter_states.get("codex")
                if codex_state_b64 and getattr(self, "codex_splitter", None) is not None:
                    codex_state = QByteArray.fromBase64(codex_state_b64.encode("ascii"))
                    if not codex_state.isEmpty():
                        codex_restored = self.codex_splitter.restoreState(codex_state)

                restored = True
            except Exception as e:
                print(f"[WARN] 스플리터 상태 복원 실패: {e}")
                restored = False
                codex_restored = False

        # 복원에 실패했거나 저장된 상태가 없으면 기본값으로 설정
        if not restored:
            self.left_splitter.setSizes([150, 250])
            self.main_splitter.setSizes([400, 560])
        if getattr(self, "codex_splitter", None) is not None and not codex_restored:
            self.codex_splitter.setSizes([208, 920])
        # --- [END] 스플리터 상태 복원 로직 ---

        # 초기 상태 업데이트
        self._update_move_button_state()
        QTimer.singleShot(0, self._sync_codex_target_from_current_fav_item)

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
                self._flush_pending_favorites_save()
            except Exception:
                pass
            self._save_favorites()
            self._flush_pending_buffer_structure_save()
            self._flush_pending_settings_save()
            print("[DBG][FLUSH] Favorites saved on exit")
        except Exception as e:
            print(f"[ERR][FLUSH] Failed to save favorites on exit: {e}")
        super().closeEvent(event)

    def update_status_and_ui(self, status_text: str, is_connected: bool):
        self.connection_status_label.setText(status_text)
        self.center_button.setEnabled(is_connected)
        search_input = getattr(self, "search_input", None)
        if search_input is not None:
            search_input.setEnabled(is_connected)
        search_button = getattr(self, "search_button", None)
        if search_button is not None:
            search_button.setEnabled(is_connected)
        open_all_busy = bool(
            self._open_all_notebooks_worker
            and self._open_all_notebooks_worker.isRunning()
        )
        if hasattr(self, "open_all_notebooks_action"):
            self.open_all_notebooks_action.setEnabled(is_connected and not open_all_busy)

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
            win = None
            handle = sig.get("handle")
            if handle:
                try:
                    candidate = Desktop(backend="uia").window(handle=handle)
                    if candidate.is_visible():
                        win = candidate
                except Exception:
                    win = None
            if win is None:
                win = reacquire_window_by_signature(sig)
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

    def _try_activate_favorite_fastpath_v2(
        self,
        item: QTreeWidgetItem,
        sig: Dict[str, Any],
        target: Dict[str, Any],
        display_name: str,
        *,
        started_at: Optional[float] = None,
    ) -> bool:
        direct_source = "same_window"
        if self._is_sig_same_as_connected_window(sig):
            win = self.onenote_window
        else:
            direct_source = "direct_connect"
            win = None
            handle = sig.get("handle")
            if handle:
                try:
                    candidate = Desktop(backend="uia").window(handle=handle)
                    if candidate.is_visible():
                        win = candidate
                except Exception:
                    win = None
            if win is None:
                win = reacquire_window_by_signature(sig)
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

        notebook_text = _strip_stale_favorite_prefix(
            str((target or {}).get("notebook_text") or "").strip()
        )
        section_text = str((target or {}).get("section_text") or "").strip()
        display_text = _strip_stale_favorite_prefix(display_name)
        target_kind = "section" if section_text else "notebook"
        expected_text = section_text
        resolved_name = ""
        resolved_notebook_id = ""
        resolution_mode = "quick"
        selected_notebook_item = None

        def _attempt_select(kind: str, text: str) -> bool:
            nonlocal selected_notebook_item
            if not text:
                return False
            if kind == "section":
                return select_section_by_text(self.onenote_window, text, tree)
            selected_notebook_item = select_notebook_item_by_text(
                self.onenote_window,
                text,
                tree,
                center_after_select=False,
            )
            return selected_notebook_item is not None

        ok = False
        if target_kind == "section":
            ok = _attempt_select("section", expected_text)
        else:
            quick_candidates = []
            for cand in (notebook_text, display_text):
                if cand and cand not in quick_candidates:
                    quick_candidates.append(cand)
            for cand in quick_candidates:
                if _attempt_select("notebook", cand):
                    expected_text = cand
                    ok = True
                    break

        if not ok:
            requested_notebook_id = str((target or {}).get("notebook_id") or "").strip()
            if target_kind == "notebook" and not requested_notebook_id:
                visible_names = _collect_root_notebook_names_from_tree(tree)
                if visible_names:
                    quick_error = _build_notebook_not_found_error(
                        notebook_text or display_text,
                        visible_names,
                    )
                    stale_name = self._mark_favorite_item_stale(item, display_name)
                    fail_msg = quick_error or f"항목 찾기 실패: '{stale_name or display_name}'"
                    self.update_status_and_ui(
                        fail_msg,
                        self.center_button.isEnabled(),
                    )
                    print(
                        "[DBG][FAV][FASTPATH]",
                        "quick_abort",
                        f"error={fail_msg!r}",
                        f"elapsed_ms={(time.perf_counter() - (started_at or time.perf_counter())) * 1000.0:.1f}",
                    )
                    return True
            target_info = _resolve_favorite_activation_target(target, display_name)
            if not target_info.get("ok", True):
                stale_name = self._mark_favorite_item_stale(item, display_name)
                fail_msg = (
                    target_info.get("error")
                    or f"항목 찾기 실패: '{stale_name or display_name}'"
                )
                self.update_status_and_ui(
                    fail_msg,
                    self.center_button.isEnabled(),
                )
                print(
                    "[DBG][FAV][FASTPATH]",
                    "resolve_abort",
                    f"error={fail_msg!r}",
                    f"elapsed_ms={(time.perf_counter() - (started_at or time.perf_counter())) * 1000.0:.1f}",
                )
                return True
            target_kind = target_info.get("target_kind") or target_kind
            expected_text = target_info.get("expected_center_text") or expected_text
            resolved_name = target_info.get("resolved_name") or ""
            resolved_notebook_id = target_info.get("resolved_notebook_id") or ""
            resolution_mode = "resolved"

        print(
            "[DBG][FAV][FASTPATH]",
            direct_source,
            f"kind={target_kind}",
            f"text={expected_text!r}",
            f"mode={resolution_mode}",
            f"elapsed_ms={(time.perf_counter() - (started_at or time.perf_counter())) * 1000.0:.1f}",
        )

        try:
            self.onenote_window.set_focus()
        except Exception:
            pass

        if not ok:
            ok = _attempt_select(target_kind, expected_text)

        if not ok:
            self._cache_tree_control()
            tree = self.tree_control
            if tree:
                ok = _attempt_select(target_kind, expected_text)

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
                resolved_name,
                resolved_notebook_id,
            )

        self.center_selected_item_action(
            debug_source="fav_fastpath",
            started_at=started_at,
            skip_precheck=True,
            allow_retry=(target_kind != "notebook"),
            preselected_item=selected_notebook_item if target_kind == "notebook" else None,
            preselected_tree_control=tree if target_kind == "notebook" else None,
        )
        return True

    def eventFilter(self, obj, event):
        try:
            list_widget = getattr(self, "onenote_list_widget", None)
            if list_widget is not None and obj is list_widget.viewport():
                event_type = event.type()
                if event_type == QEvent.Type.MouseButtonPress:
                    app = QApplication.instance()
                    delay_ms = 120
                    if app is not None:
                        try:
                            delay_ms = int(app.doubleClickInterval()) + 30
                        except Exception:
                            delay_ms = 120
                    self._schedule_onenote_list_auto_refresh(delay_ms=delay_ms)
                elif event_type == QEvent.Type.MouseButtonDblClick:
                    self._cancel_pending_onenote_list_auto_refresh()
        except Exception:
            pass
        return super().eventFilter(obj, event)

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
        self._last_onenote_list_refresh_at = time.monotonic()
        if self._onenote_list_refresh_timer.isActive():
            self._onenote_list_refresh_timer.stop()

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
        print(f"[DBG][LIST] onenote_windows={len(results)}")
        selection_key = self._pending_onenote_list_selection_key
        self._pending_onenote_list_selection_key = None

        if not results:
            self.onenote_list_widget.addItem("실행 중인 OneNote 창을 찾지 못했습니다.")
        else:
            for info in results:
                item = QListWidgetItem(f'{info["title"]}  [{info["class_name"]}]')
                item.setData(Qt.ItemDataRole.UserRole, copy.deepcopy(info))
                self.onenote_list_widget.addItem(item)
                item_key = (info.get("handle"), info.get("pid"), info.get("title"))
                if selection_key and item_key == selection_key:
                    self.onenote_list_widget.setCurrentItem(item)

        self.onenote_list_widget.setEnabled(True)
        self.refresh_button.setEnabled(True)

    def _cache_tree_control(self):
        self.tree_control = _find_tree_or_list(self.onenote_window)
        if self.tree_control:
            try:
                _ = self.tree_control.children()
            except Exception:
                pass

    def _perform_connection(self, info: Dict) -> bool:
        t0 = time.perf_counter()
        ensure_pywinauto()
        if not _pwa_ready:
            self.update_status_and_ui("pywinauto가 준비되지 않았습니다.", False)
            return False
        try:
            print(
                "[DBG][CONNECT] try",
                f"handle={info.get('handle')}",
                f"pid={info.get('pid')}",
                f"class={info.get('class_name')}",
                f"title={info.get('title')!r}",
            )
            target = None
            handle = info.get("handle")
            if handle:
                try:
                    target = Desktop(backend="uia").window(handle=handle)
                    if not target.is_visible():
                        target = None
                except Exception:
                    target = None
            if target is None:
                target = reacquire_window_by_signature(info)
            if target is None:
                raise ElementNotFoundError

            self.onenote_window = target
            window_title = self.onenote_window.window_text()
            save_connection_info(self.onenote_window)
            elapsed_ms = (time.perf_counter() - t0) * 1000.0
            print(
                f"[DBG][CONNECT] success title={window_title!r} "
                f"elapsed_ms={elapsed_ms:.1f} at_s={(time.perf_counter() - self._t_boot):.3f}"
            )

            status_text = f"연결됨: '{window_title}'"
            self.update_status_and_ui(status_text, True)
            self._cache_tree_control()
            return True

        except ElementNotFoundError:
            elapsed_ms = (time.perf_counter() - t0) * 1000.0
            print(
                "[DBG][CONNECT] fail: target not found/visible "
                f"elapsed_ms={elapsed_ms:.1f} at_s={(time.perf_counter() - self._t_boot):.3f}"
            )
            self.update_status_and_ui("연결 실패: 선택한 창이 보이지 않습니다.", False)
            self.refresh_onenote_list()
            return False
        except Exception as e:
            elapsed_ms = (time.perf_counter() - t0) * 1000.0
            print(
                f"[DBG][CONNECT] exception elapsed_ms={elapsed_ms:.1f} "
                f"at_s={(time.perf_counter() - self._t_boot):.3f} err={e}"
            )
            self.update_status_and_ui(f"연결 실패: {e}", False)
            return False

    def connect_and_center_from_list_item(self, item):
        started_at = time.perf_counter()
        self._cancel_pending_onenote_list_auto_refresh()
        info = None
        row = -1
        item_text = ""

        if item is None:
            try:
                item = self.onenote_list_widget.currentItem()
            except Exception:
                item = None

        if item is not None:
            try:
                item_text = item.text() or ""
            except Exception:
                item_text = ""
            try:
                self.onenote_list_widget.setCurrentItem(item)
            except Exception:
                pass
            try:
                raw = item.data(Qt.ItemDataRole.UserRole)
                if isinstance(raw, dict):
                    info = raw
            except Exception:
                info = None
            if info is None:
                try:
                    row = self.onenote_list_widget.row(item)
                except Exception:
                    row = -1
        else:
            row = self.onenote_list_widget.currentRow()

        if info is None and 0 <= row < len(self.onenote_windows_info):
            info = self.onenote_windows_info[row]

        print(
            "[DBG][LIST][ACTIVATE]",
            f"text={item_text!r}",
            f"row={row}",
            f"has_info={bool(info)}",
            f"at_s={(time.perf_counter() - self._t_boot):.3f}",
        )
        if not info:
            self.update_status_and_ui("OneNote 창 선택 정보를 읽지 못했습니다. 목록을 새로고침해 주세요.", False)
            return

        connect_key = (info.get("handle"), info.get("pid"), info.get("title"))
        now = time.monotonic()
        if (
            self._last_list_connect_key == connect_key
            and (now - self._last_list_connect_at) < 0.35
        ):
            print(f"[DBG][LIST][SKIP] duplicate key={connect_key!r}")
            return
        self._last_list_connect_key = connect_key
        self._last_list_connect_at = now

        current_handle = self._current_onenote_handle()
        target_handle = info.get("handle")
        if current_handle and target_handle and int(target_handle) == current_handle:
            print(
                "[DBG][LIST][FASTPATH] already_connected "
                f"handle={current_handle} elapsed_ms={(time.perf_counter() - started_at) * 1000.0:.1f} "
                f"at_s={(time.perf_counter() - self._t_boot):.3f}"
            )
            self.center_selected_item_action(
                debug_source="list_dblclick_same_window",
                started_at=started_at,
            )
            return

        if self._perform_connection(info):
            self.center_selected_item_action(
                debug_source="list_dblclick_connect",
                started_at=started_at,
            )

    def select_other_window(self):
        dialog = OtherWindowSelectionDialog(self.my_pid, self)
        if dialog.exec():
            info = dialog.selected_info
            if info:
                self._perform_connection(info)

    def disconnect_and_clear_info(self):
        self.onenote_window = None
        self.tree_control = None
        self.update_status_and_ui("연결 해제됨.", False)

        self.settings["connection_signature"] = None
        self._save_settings_to_file(immediate=True)

    def _pre_action_check(self) -> bool:
        """
        OneNote 관련 액션을 실행하기 전 선행 조건 체크.
        False가 나오는 이유를 터미널에 상세히 출력한다.
        """
        print("[DBG][PRECHECK] ENTER")
        try:
            w = getattr(self, "onenote_window", None)
            print(f"[DBG][PRECHECK] onenote_window={w}")
        except Exception as e:
            print(f"[DBG][PRECHECK] onenote_window read EXC: {e}")
            w = None

        # 1) OneNote 윈도우 핸들 확보 여부
        try:
            hwnd = None
            if w is not None:
                hwnd = getattr(w, "handle", None)
                if callable(hwnd):
                    hwnd = w.handle()
            print(f"[DBG][PRECHECK] hwnd={hwnd}")
        except Exception as e:
            print(f"[DBG][PRECHECK] hwnd EXC: {e}")
            hwnd = None

        if not hwnd:
            print("[DBG][PRECHECK] FAIL: hwnd is None/0 (OneNote 창 연결 안됨)")
            try:
                self.update_status_and_ui("OneNote 창이 연결되지 않았습니다. 먼저 OneNote 창 연결/선택을 해주세요.", False)
            except Exception:
                pass
            return False

        # 2) pywinauto backend / wrapper 사용 가능 여부
        try:
            ensure_pywinauto()
            print("[DBG][PRECHECK] ensure_pywinauto OK")
        except Exception as e:
            print(f"[DBG][PRECHECK] FAIL: ensure_pywinauto EXC: {e}")
            return False

        # 3) 포그라인드/활성화 조건이 있으면 여기서 확인
        try:
            # 프로젝트에 기존 함수가 있으면 그대로 호출하되, 실패 사유를 찍는다.
            if hasattr(self, "_bring_onenote_to_front"):
                ok_focus = self._bring_onenote_to_front()
                print(f"[DBG][PRECHECK] _bring_onenote_to_front={ok_focus}")
                if ok_focus is False:
                    print("[DBG][PRECHECK] FAIL: bring_onenote_to_front returned False")
                    return False
        except Exception as e:
            print(f"[DBG][PRECHECK] bring/front EXC: {e}")
            return False

        # 4) 트리 컨트롤 찾기 조건 (기존에 precheck에서 강제하는 경우가 많음)
        try:
            tc = getattr(self, "tree_control", None)
            print(f"[DBG][PRECHECK] tree_control(before)={tc}")
            if not tc:
                finder = globals().get("_find_tree_or_list", None)
                if callable(finder):
                    tc = finder(w)
                    self.tree_control = tc
                print(f"[DBG][PRECHECK] tree_control(after)={tc}")
            if not tc:
                print("[DBG][PRECHECK] FAIL: tree_control not found")
                return False
        except Exception as e:
            print(f"[DBG][PRECHECK] tree_control find EXC: {e}")
            return False

        print("[DBG][PRECHECK] PASS")
        return True

    def center_selected_item_action(
        self,
        checked: bool = False,
        *,
        debug_source: str = "button",
        started_at: Optional[float] = None,
        skip_precheck: bool = False,
        allow_retry: bool = True,
        preselected_item=None,
        preselected_tree_control=None,
    ):
        op_started_at = started_at or time.perf_counter()
        print(
            f"[DBG][CENTER][START] source={debug_source} "
            f"at_s={(time.perf_counter() - self._t_boot):.3f}"
        )
        if not skip_precheck and not self._pre_action_check():
            print(
                f"[DBG][CENTER][ABORT] source={debug_source} "
                f"elapsed_ms={(time.perf_counter() - op_started_at) * 1000.0:.1f} "
                f"at_s={(time.perf_counter() - self._t_boot):.3f}"
            )
            return

        if preselected_tree_control is not None:
            self.tree_control = preselected_tree_control
        elif not self.tree_control:
            self.tree_control = _find_tree_or_list(self.onenote_window)

        success, item_name = scroll_selected_item_to_center(
            self.onenote_window,
            self.tree_control,
            selected_item=preselected_item,
        )

        if success:
            print(
                f"[DBG][CENTER][DONE] source={debug_source} success=True "
                f"item={item_name!r} elapsed_ms={(time.perf_counter() - op_started_at) * 1000.0:.1f} "
                f"at_s={(time.perf_counter() - self._t_boot):.3f}"
            )
            self.update_status_and_ui(f"성공: '{item_name}' 중앙 정렬 완료.", True)
        elif allow_retry:
            self.tree_control = _find_tree_or_list(self.onenote_window)
            success, item_name = scroll_selected_item_to_center(
                self.onenote_window, self.tree_control
            )
            if success:
                print(
                    f"[DBG][CENTER][DONE] source={debug_source} success=True retry=1 "
                    f"item={item_name!r} elapsed_ms={(time.perf_counter() - op_started_at) * 1000.0:.1f} "
                    f"at_s={(time.perf_counter() - self._t_boot):.3f}"
                )
                self.update_status_and_ui(f"성공: '{item_name}' 중앙 정렬 완료.", True)
            else:
                print(
                    f"[DBG][CENTER][DONE] source={debug_source} success=False "
                    f"elapsed_ms={(time.perf_counter() - op_started_at) * 1000.0:.1f} "
                    f"at_s={(time.perf_counter() - self._t_boot):.3f}"
                )
        else:
            print(
                f"[DBG][CENTER][DONE] source={debug_source} success=False retry=skip "
                f"elapsed_ms={(time.perf_counter() - op_started_at) * 1000.0:.1f} "
                f"at_s={(time.perf_counter() - self._t_boot):.3f}"
            )
            self.update_status_and_ui(
                "실패: 선택 항목을 찾거나 정렬하지 못했습니다.", True
            )

    def _open_all_notebooks_from_connected_onenote(self):
        if (
            self._open_all_notebooks_worker is not None
            and self._open_all_notebooks_worker.isRunning()
        ):
            self.update_status_and_ui("이미 실제 OneNote 전체 열기 작업이 실행 중입니다.", True)
            return

        win = getattr(self, "onenote_window", None)
        hwnd = getattr(win, "handle", None) if win is not None else None
        if callable(hwnd):
            try:
                hwnd = hwnd()
            except Exception:
                hwnd = None
        if not hwnd:
            self.update_status_and_ui(
                "OneNote 창이 연결되지 않았습니다. 먼저 OneNote 창을 연결하세요.",
                False,
            )
            return

        ensure_pywinauto()
        if not _pwa_ready:
            self.update_status_and_ui("오류: 자동화 모듈이 로드되지 않았습니다.", True)
            return

        try:
            win.set_focus()
        except Exception:
            pass

        sig = build_window_signature(win)
        if not sig:
            self.update_status_and_ui("오류: OneNote 창 정보를 읽지 못했습니다.", True)
            return

        worker = OpenAllNotebooksWorker(sig, self)
        self._open_all_notebooks_worker = worker
        self.update_status_and_ui("실제 OneNote 전체 열기 준비 중...", True)

        def _on_progress(message: str):
            if self._open_all_notebooks_worker is worker:
                self.connection_status_label.setText(message)

        def _on_done(result: Dict[str, Any]):
            if self._open_all_notebooks_worker is not worker:
                return

            self._open_all_notebooks_worker = None
            try:
                worker.deleteLater()
            except Exception:
                pass

            connected = self._apply_connected_window_info(result.get("window_info"))
            is_connected = connected or bool(getattr(self, "onenote_window", None))
            opened_count = int(result.get("opened_count") or 0)
            remaining = result.get("remaining_names") or []
            error = (result.get("error") or "").strip()

            if result.get("ok"):
                if opened_count > 0:
                    self.update_status_and_ui(
                        f"실제 OneNote 전체 열기 완료: {opened_count}개",
                        is_connected,
                    )
                else:
                    self.update_status_and_ui(
                        "열어야 할 전자필기장이 더 이상 없습니다.",
                        is_connected,
                    )
                return

            if remaining:
                remain_preview = ", ".join(remaining[:3])
                if len(remaining) > 3:
                    remain_preview += " ..."
                suffix = f" 남은 후보: {remain_preview}"
            else:
                suffix = ""
            detail = error or "실제 OneNote 전체 열기에 실패했습니다."
            self.update_status_and_ui(
                f"{detail} (시도 {opened_count}개).{suffix}",
                is_connected,
            )

        worker.progress.connect(_on_progress)
        worker.done.connect(_on_done)
        worker.finished.connect(lambda: None)
        worker.start()

    def _search_and_select_section(self):
        """입력창의 텍스트로 섹션을 검색하고 선택 및 중앙 정렬합니다."""
        if not self._pre_action_check():
            return

        search_text = self.search_input.text().strip()
        if not search_text:
            self.update_status_and_ui("검색할 내용을 입력하세요.", True)
            return

        if not self.tree_control:
            self.tree_control = _find_tree_or_list(self.onenote_window)

        self.update_status_and_ui(f"'{search_text}' 섹션을 검색 중...", True)

        success = select_section_by_text(
            self.onenote_window, search_text, self.tree_control
        )

        if success:
            QTimer.singleShot(100, self.center_selected_item_action)
            self.update_status_and_ui(f"검색 성공: '{search_text}' 선택 완료.", True)
        else:
            self.update_status_and_ui(
                f"검색 실패: '{search_text}' 섹션을 찾을 수 없습니다.", True
            )

    def _calc_nodes_signature(self, obj):
        """리스트/딕트의 안정적인 시그니처를 계산합니다."""
        try:
            raw = json.dumps(obj, sort_keys=True, ensure_ascii=False)
            return hashlib.md5(raw.encode("utf-8")).hexdigest()
        except Exception:
            return None

    def _invalidate_aggregate_cache(self, *, invalidate_classified_keys: bool = True):
        """종합 버퍼 계산/표시 캐시를 무효화합니다."""
        self._aggregate_cache_valid = False
        self._aggregate_cache = []
        self._aggregate_display_cache_sig = None
        self._aggregate_display_cache = []
        self._aggregate_display_cache_kind = None
        self._aggregate_display_cache_source_id = 0
        if invalidate_classified_keys:
            self._aggregate_classified_keys_cache_valid = False
            self._aggregate_classified_keys_cache = set()

    def _get_sorted_aggregate_display_nodes(self, nodes, *, kind: str):
        """
        종합 버퍼 표시용 정렬 결과를 캐시합니다.
        - kind='saved': 종합 버퍼에 직접 저장된 notebook/group 표시본
        - kind='built': 전체 버퍼에서 수집해 만든 fallback 표시본
        """
        source_id = id(nodes)
        if (
            source_id
            and source_id == getattr(self, "_aggregate_display_cache_source_id", 0)
            and kind == getattr(self, "_aggregate_display_cache_kind", None)
            and isinstance(getattr(self, "_aggregate_display_cache", None), list)
        ):
            return self._aggregate_display_cache
        sig = self._calc_nodes_signature(nodes)
        if (
            sig is not None
            and sig == getattr(self, "_aggregate_display_cache_sig", None)
            and kind == getattr(self, "_aggregate_display_cache_kind", None)
            and isinstance(getattr(self, "_aggregate_display_cache", None), list)
        ):
            return self._aggregate_display_cache
        data = self._sorted_copy_nodes_by_name(nodes)
        self._aggregate_display_cache_sig = sig
        self._aggregate_display_cache_kind = kind
        self._aggregate_display_cache = data
        self._aggregate_display_cache_source_id = source_id
        return data

    def _aggregate_notebook_key_from_node(self, node: Any) -> str:
        keys = self._aggregate_notebook_keys_from_node(node)
        if not keys:
            return ""
        id_keys = sorted(k for k in keys if k.startswith("id:"))
        if id_keys:
            return id_keys[0]
        return sorted(keys)[0]

    def _aggregate_notebook_keys_from_node(self, node: Any) -> Set[str]:
        if not isinstance(node, dict):
            return set()
        target = node.get("target") or {}
        keys: Set[str] = set()
        notebook_id = str(target.get("notebook_id") or "").strip()
        if notebook_id:
            keys.add("id:" + notebook_id.casefold())
        name = (
            str(target.get("notebook_text") or "").strip()
            or str(node.get("name") or "").strip()
        )
        name_key = _normalize_notebook_name_key(name)
        if name_key:
            keys.add("name:" + name_key)
        return keys

    def _collect_notebook_nodes_from_nodes(self, nodes: Any) -> List[Dict[str, Any]]:
        if not isinstance(nodes, list):
            return []
        found: List[Dict[str, Any]] = []
        seen: Set[str] = set()
        stack = list(reversed(nodes))
        while stack:
            node = stack.pop()
            if not isinstance(node, dict):
                continue
            if node.get("type") == "notebook":
                node_keys = self._aggregate_notebook_keys_from_node(node)
                if node_keys and not (node_keys & seen):
                    seen.update(node_keys)
                    found.append(
                        {
                            "type": "notebook",
                            "id": node.get("id") or str(uuid.uuid4()),
                            "name": node.get("name") or "전자필기장",
                            "target": copy.deepcopy(node.get("target") or {}),
                        }
                    )
                continue
            children = node.get("children")
            if isinstance(children, list):
                stack.extend(reversed(children))
            data = node.get("data")
            if isinstance(data, list):
                stack.extend(reversed(data))
        try:
            found.sort(key=lambda n: _name_sort_key((n or {}).get("name", "")))
        except Exception:
            pass
        return found

    def _nodes_contain_notebook(self, nodes: Any) -> bool:
        if not isinstance(nodes, list):
            return False
        stack = list(nodes)
        while stack:
            node = stack.pop()
            if not isinstance(node, dict):
                continue
            if node.get("type") == "notebook":
                return True
            children = node.get("children")
            if isinstance(children, list):
                stack.extend(children)
            data = node.get("data")
            if isinstance(data, list):
                stack.extend(data)
        return False

    def _collect_classified_aggregate_notebook_keys(self) -> Set[str]:
        if getattr(self, "_aggregate_classified_keys_cache_valid", False):
            return set(getattr(self, "_aggregate_classified_keys_cache", set()))

        keys: Set[str] = set()

        def _walk_fav_nodes(nodes: Any) -> None:
            if not isinstance(nodes, list):
                return
            for node in nodes:
                if not isinstance(node, dict):
                    continue
                if node.get("type") == "notebook":
                    keys.update(self._aggregate_notebook_keys_from_node(node))
                children = node.get("children")
                if isinstance(children, list):
                    _walk_fav_nodes(children)

        def _walk_buffers(nodes: Any) -> None:
            if not isinstance(nodes, list):
                return
            for node in nodes:
                if not isinstance(node, dict):
                    continue
                if node.get("type") == "buffer":
                    if node.get("id") == AGG_BUFFER_ID:
                        continue
                    _walk_fav_nodes(node.get("data") or [])
                elif node.get("type") == "group":
                    _walk_buffers(node.get("children") or [])

        _walk_buffers(self.settings.get("favorites_buffers", []))
        self._aggregate_classified_keys_cache = set(keys)
        self._aggregate_classified_keys_cache_valid = True
        return keys

    def _build_aggregate_categorized_display_nodes(
        self, source_nodes: Any
    ) -> List[Dict[str, Any]]:
        notebooks = self._collect_notebook_nodes_from_nodes(source_nodes)
        if not notebooks:
            notebooks = self._collect_notebook_nodes_from_nodes(
                _collect_all_sections_dedup(self.settings)
            )

        classified_keys = self._collect_classified_aggregate_notebook_keys()
        unclassified: List[Dict[str, Any]] = []
        classified: List[Dict[str, Any]] = []
        for notebook in notebooks:
            notebook_keys = self._aggregate_notebook_keys_from_node(notebook)
            if notebook_keys and notebook_keys & classified_keys:
                classified.append(notebook)
            else:
                unclassified.append(notebook)

        try:
            unclassified.sort(key=lambda n: _name_sort_key((n or {}).get("name", "")))
            classified.sort(key=lambda n: _name_sort_key((n or {}).get("name", "")))
        except Exception:
            pass

        return [
            {
                "type": "group",
                "id": AGG_UNCLASSIFIED_GROUP_ID,
                "name": AGG_UNCLASSIFIED_GROUP_NAME,
                "children": unclassified,
            },
            {
                "type": "group",
                "id": AGG_CLASSIFIED_GROUP_ID,
                "name": AGG_CLASSIFIED_GROUP_NAME,
                "children": classified,
            },
        ]

    def _aggregate_classification_signature_from_nodes(self, nodes: Any) -> tuple:
        if not isinstance(nodes, list):
            return tuple()

        def _collect_keys(children: Any) -> List[str]:
            keys: List[str] = []
            if not isinstance(children, list):
                return keys
            stack = list(children)
            while stack:
                child = stack.pop()
                if not isinstance(child, dict):
                    continue
                keys.extend(self._aggregate_notebook_keys_from_node(child))
                nested = child.get("children")
                if isinstance(nested, list):
                    stack.extend(nested)
            return keys

        groups: Dict[str, List[str]] = {
            AGG_UNCLASSIFIED_GROUP_ID: [],
            AGG_CLASSIFIED_GROUP_ID: [],
        }
        fallback_keys: List[str] = []

        for node in nodes:
            if not isinstance(node, dict):
                continue
            node_type = node.get("type")
            node_id = node.get("id")
            if node_type == "group" and node_id in groups:
                groups[node_id].extend(_collect_keys(node.get("children") or []))
            elif node_type == "notebook":
                fallback_keys.extend(self._aggregate_notebook_keys_from_node(node))

        if fallback_keys and not (groups[AGG_UNCLASSIFIED_GROUP_ID] or groups[AGG_CLASSIFIED_GROUP_ID]):
            return (("flat", tuple(sorted(set(fallback_keys)))),)

        return (
            (
                AGG_UNCLASSIFIED_GROUP_ID,
                tuple(sorted(set(groups[AGG_UNCLASSIFIED_GROUP_ID]))),
            ),
            (
                AGG_CLASSIFIED_GROUP_ID,
                tuple(sorted(set(groups[AGG_CLASSIFIED_GROUP_ID]))),
            ),
        )

    def _serialize_current_module_tree(self) -> List[Dict[str, Any]]:
        nodes: List[Dict[str, Any]] = []
        try:
            root = self.fav_tree.invisibleRootItem()
            for i in range(root.childCount()):
                nodes.append(self._serialize_fav_item(root.child(i)))
        except Exception:
            pass
        return nodes

    def _aggregate_source_nodes_for_fast_classification(
        self, current_nodes: Optional[List[Dict[str, Any]]] = None
    ) -> List[Dict[str, Any]]:
        if getattr(self, "active_buffer_id", None) == AGG_BUFFER_ID:
            if current_nodes is None:
                current_nodes = self._serialize_current_module_tree()
            if self._nodes_contain_notebook(current_nodes):
                return current_nodes

        agg_node = _find_buffer_node_by_id(
            self.settings.get("favorites_buffers", []),
            AGG_BUFFER_ID,
        )
        saved = (agg_node or {}).get("data", []) if isinstance(agg_node, dict) else []
        if self._nodes_contain_notebook(saved):
            return saved
        return _collect_all_sections_dedup(self.settings)

    def _persist_active_aggregate_data(self, data: List[Dict[str, Any]]) -> None:
        if getattr(self, "active_buffer_id", None) != AGG_BUFFER_ID:
            return
        new_sig = self._calc_nodes_signature(data)
        if self.active_buffer_node is not None:
            self.active_buffer_node["data"] = data
        if self.active_buffer_item is not None:
            payload = self.active_buffer_item.data(0, ROLE_DATA) or {}
            payload["data"] = data
            self.active_buffer_item.setData(0, ROLE_DATA, payload)
        settings_node = _find_buffer_node_by_id(
            self.settings.get("favorites_buffers", []),
            AGG_BUFFER_ID,
        )
        if isinstance(settings_node, dict):
            old_sig = self._calc_nodes_signature(settings_node.get("data", []))
            settings_node["data"] = data
            self._active_buffer_settings_node = settings_node
            if not new_sig or new_sig != old_sig:
                self._save_settings_to_file()

    def _refresh_active_aggregate_classification_from_saved_data(
        self,
        *,
        current_nodes: Optional[List[Dict[str, Any]]] = None,
        persist: bool = True,
        show_status: bool = False,
    ) -> bool:
        if getattr(self, "active_buffer_id", None) != AGG_BUFFER_ID:
            if show_status:
                QMessageBox.information(self, "안내", "종합 버퍼에서만 사용할 수 있습니다.")
            return False
        if getattr(self, "_aggregate_reclassify_in_progress", False):
            return False

        if current_nodes is None:
            current_nodes = self._serialize_current_module_tree()
        source_nodes = self._aggregate_source_nodes_for_fast_classification(current_nodes)
        categorized = self._build_aggregate_categorized_display_nodes(source_nodes)
        current_sig = self._aggregate_classification_signature_from_nodes(current_nodes)
        next_sig = self._aggregate_classification_signature_from_nodes(categorized)

        if current_sig and current_sig == next_sig:
            if show_status:
                try:
                    self.connection_status_label.setText(
                        "종합 분류 상태가 이미 최신입니다."
                    )
                except Exception:
                    pass
            return False

        self._aggregate_reclassify_in_progress = True
        try:
            self._load_favorites_into_center_tree(categorized)
            self._fav_reset_undo_context_from_data(
                categorized,
                reason="aggregate_fast_reclassify",
            )
            if persist:
                self._persist_active_aggregate_data(categorized)
            if show_status:
                try:
                    unclassified_count = len(categorized[0].get("children") or [])
                    classified_count = len(categorized[1].get("children") or [])
                    self.connection_status_label.setText(
                        "종합 분류 새로고침 완료: "
                        f"미분류 {unclassified_count}개, 분류됨 {classified_count}개"
                    )
                except Exception:
                    pass
            return True
        finally:
            self._aggregate_reclassify_in_progress = False

    def _refresh_active_aggregate_classification_action(self) -> None:
        self._refresh_active_aggregate_classification_from_saved_data(
            persist=True,
            show_status=True,
        )

    def _build_aggregate_buffer(self):
        """모든 섹션을 수집하여 종합 버퍼 데이터를 생성합니다."""
        if getattr(self, "_aggregate_cache_valid", False):
            return self._aggregate_cache
        data = _collect_all_sections_dedup(self.settings)
        self._aggregate_cache = data
        self._aggregate_cache_valid = True
        return data

    def _finish_boot_sequence(self):
        """부팅 완료 단계에서 마지막 상태(활성 버퍼 데이터)를 강제 복원합니다."""
        print("[BOOT] Starting final boot sequence...")
        t0 = time.perf_counter()
        try:
            # 활성 버퍼 다시 확인
            active_id = self.settings.get("active_buffer_id")
            found_data = []
            buf_name = "None"
            if active_id and getattr(self, "_last_loaded_center_buffer_id", None) == active_id:
                print(f"[BOOT][PERF] final restore skipped; active buffer already loaded: {active_id}")
                return

            found_item = self._buffer_item_index.get(active_id) if active_id else None
            if found_item is not None:
                payload = found_item.data(0, ROLE_DATA) or {}
                buf_name = found_item.text(0)
                if active_id == AGG_BUFFER_ID:
                    saved = payload.get("data", []) or []
                    source = saved if self._nodes_have_any_type(saved, {"notebook", "group"}) else self._build_aggregate_buffer()
                    found_data = self._build_aggregate_categorized_display_nodes(source)
                else:
                    found_data = payload.get("data", [])
            else:
                iterator = QTreeWidgetItemIterator(self.buffer_tree)
                while iterator.value():
                    item = iterator.value()
                    payload = item.data(0, ROLE_DATA) or {}
                    if payload.get("id") == active_id:
                        buf_name = item.text(0)
                        if active_id == AGG_BUFFER_ID:
                            saved = payload.get("data", []) or []
                            source = saved if self._nodes_have_any_type(saved, {"notebook", "group"}) else self._build_aggregate_buffer()
                            found_data = self._build_aggregate_categorized_display_nodes(source)
                        else:
                            found_data = payload.get("data", [])
                        break
                    iterator += 1

            # 강제 리빌드
            self._rebuild_modules_from_buffer(buf_name, found_data)
        except Exception as e:
            print(f"[BOOT][RESTORE][BUF_REBUILD_FAIL] {e}")
            traceback.print_exc()
        finally:
            self._boot_loading = False
            total_ms = (time.perf_counter() - t0) * 1000.0
            try:
                total_nodes = self._count_nodes_recursive(found_data)
            except Exception:
                total_nodes = len(found_data) if isinstance(found_data, list) else 0
            print(f"[BOOT][PERF] final_restore_ms={total_ms:.1f} total_nodes={total_nodes}")
            print(f"[BOOT] Final boot sequence finished. (Active: {buf_name})")

    def _rebuild_modules_from_buffer(self, buffer_name: str, nodes: list):
        """
        저장된 favorites_buffers 기준으로
        2패널(모듈/전자필기장 영역)을 복원합니다.
        """
        print(f"[BOOT][BUF_RESTORE] buffer='{buffer_name}' count={len(nodes)}")

        if nodes:
            self._load_favorites_into_center_tree(nodes)
        else:
            self._clear_module_panel()
        self._last_loaded_center_buffer_id = (
            self.settings.get("active_buffer_id") if buffer_name != "None" else None
        )

        # 활성 상태바 업데이트
        if buffer_name != "None":
            self.connection_status_label.setText(f"준비됨 (활성 버퍼: {buffer_name})")

    def _clear_module_panel(self):
        """중앙 모듈(즐겨찾기) 패널을 완전히 비웁니다."""
        was_updates_enabled = True
        try:
            self.fav_tree.blockSignals(True)
            was_updates_enabled = self.fav_tree.updatesEnabled()
            self.fav_tree.setUpdatesEnabled(False)
            self.fav_tree.clear()
            self._last_loaded_center_buffer_id = None
            self._last_center_payload_snapshot = None
            self._last_center_payload_source_id = 0
            self._last_center_payload_hash = None # 해시 캐시 초기화
            self._module_search_index = []
            self._module_search_last_match_records = []
            self._module_search_highlighted_by_id = {}
            self._module_search_match_count = 0
        except Exception:
            pass
        finally:
            try:
                self.fav_tree.setUpdatesEnabled(was_updates_enabled)
                if was_updates_enabled:
                    self.fav_tree.viewport().update()
            except Exception:
                pass
            self.fav_tree.blockSignals(False)

    # ----------------- 15. 즐겨찾기 로드/세이브 (계층형 버퍼 시스템 적용) -----------------
    def _load_buffers_and_favorites(self):
        """설정에서 버퍼 트리를 로드합니다."""
        # ✅ 로드 전에 강제 보정 (UI에서 깨져도 복구)
        _ensure_default_and_aggregate_inplace(self.settings)
        self._invalidate_aggregate_cache()

        self.buffer_tree.blockSignals(True)
        self.buffer_tree.clear()
        self._buffer_item_index = {}
        self._first_buffer_item = None
        self._buffer_search_index = []
        self._buffer_search_last_match_records = []
        self._buffer_search_highlighted_by_id = {}
        self._buffer_search_match_count = 0
        self._buffer_search_last_applied_key = ""
        self._buffer_search_last_first_match_id = 0
        self._active_buffer_settings_node = None

        buffers_data = self.settings.get("favorites_buffers", [])

        # ✅ 레거시 설정(딕트/버퍼 없는 리스트) 자동 마이그레이션
        tmp = {
            "favorites_buffers": buffers_data,
            "active_buffer_id": self.settings.get("active_buffer_id"),
            "active_buffer": self.settings.get("active_buffer"),
        }
        if _migrate_favorites_buffers_inplace(tmp):
            self.settings.update(tmp)
            buffers_data = self.settings.get("favorites_buffers", [])
            try:
                self._save_settings_to_file()
            except Exception:
                pass

        # 방어: 그래도 dict면 비워서 크래시 방지
        if isinstance(buffers_data, dict):
            buffers_data = []

        self._boot_loading = True
        try:
            for node in buffers_data:
                self._append_buffer_node(self.buffer_tree.invisibleRootItem(), node)

            try:
                # 시작 시 프로젝트 영역은 항상 전체 펼침 상태로 보여준다.
                # 기존 expandToDepth(1)은 그룹 아래 하위 폴더가 접힌 채 남아서
                # 사용자가 앱을 켤 때마다 다시 열어야 했다.
                self._expand_buffer_groups_always(reason="startup")
            except Exception:
                pass
            self.buffer_tree.blockSignals(False)

            # 활성 버퍼 복원
            active_id = self.settings.get("active_buffer_id")
            found_item = self._buffer_item_index.get(active_id) if active_id else None

            if active_id and not found_item:
                # 트리를 순회하며 ID 찾기
                iterator = QTreeWidgetItemIterator(self.buffer_tree)
                while iterator.value():
                    item = iterator.value()
                    payload = item.data(0, ROLE_DATA)
                    if payload and payload.get("id") == active_id:
                        found_item = item
                        break
                    iterator += 1

            # 못 찾았으면 첫 번째 버퍼 선택
            if not found_item:
                found_item = self._first_buffer_item
            if not found_item:
                iterator = QTreeWidgetItemIterator(self.buffer_tree)
                while iterator.value():
                    item = iterator.value()
                    if item.data(0, ROLE_TYPE) == "buffer":
                        found_item = item
                        break
                    iterator += 1

            if found_item:
                self.buffer_tree.setCurrentItem(found_item)
                self._on_buffer_tree_item_clicked(found_item, 0)
        finally:
            self._boot_loading = False
            try:
                self.buffer_tree.blockSignals(False)
                self.buffer_tree.viewport().update()
            except Exception:
                pass
            self._refresh_project_buffer_search_highlights()

    def _append_buffer_node(self, parent: QTreeWidgetItem, node: Dict[str, Any]) -> QTreeWidgetItem:
        item = QTreeWidgetItem(parent)
        node_type = node.get("type", "buffer")
        name = node.get("name", "이름 없음")
        item.setText(0, name)
        item.setData(0, ROLE_TYPE, node_type)

        # 데이터(즐겨찾기 목록)는 트리에 직접 저장하지 않고,
        # 구조 변경 시 settings에서 다시 읽거나 관리함.
        # 여기서는 ID와 데이터 참조를 위해 payload 저장
        payload = {
            "id": node.get("id", str(uuid.uuid4())),
            "data": node.get("data", []),  # 버퍼인 경우 데이터
            "virtual": node.get("virtual"),  # 종합(aggregate) 등
            "locked": bool(node.get("locked", False)),
        }

        if node_type == "group":
            icon = getattr(self, "_icon_dir", None) or self.style().standardIcon(QApplication.style().StandardPixmap.SP_DirIcon)
            item.setIcon(0, icon)
            for child in node.get("children", []):
                self._append_buffer_node(item, child)
        else:
            # ✅ 종합(가상) 버퍼는 전용 아이콘(컴퓨터)로 표시
            if payload.get("virtual") == "aggregate":
                icon = getattr(self, "_icon_agg", None) or self.style().standardIcon(QApplication.style().StandardPixmap.SP_ComputerIcon)
                item.setIcon(0, icon)
            else:
                icon = getattr(self, "_icon_file", None) or self.style().standardIcon(QApplication.style().StandardPixmap.SP_FileIcon)
                item.setIcon(0, icon)

        item.setData(0, ROLE_DATA, payload)
        item_id = payload.get("id")
        if item_id:
            self._buffer_item_index[item_id] = item
            if node_type == "buffer" and self._first_buffer_item is None:
                self._first_buffer_item = item
        if node_type == "buffer":
            search_key = self._project_buffer_search_key_for_item(item)
            if search_key:
                parents = []
                parent_item = item.parent()
                while parent_item is not None:
                    parents.append(parent_item)
                    parent_item = parent_item.parent()
                self._buffer_search_index.append(
                    {"item": item, "key": search_key, "parents": tuple(parents)}
                )

        # ✅ locked 노드는 편집/이동/드롭 막기 (Default 그룹, 종합)
        if payload.get("locked"):
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable & ~Qt.ItemFlag.ItemIsDragEnabled & ~Qt.ItemFlag.ItemIsDropEnabled)
        else:
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsDragEnabled | Qt.ItemFlag.ItemIsDropEnabled)

        return item

    def _expand_fav_groups_always(self, *, total_nodes: int = -1, reason: str = "") -> None:
        """2패널(중앙 트리)에서 그룹 노드를 기본으로 펼쳐둡니다.

        주의: fav_tree.expandAll()은 노드가 많을 때 렉을 유발할 수 있어,
        '자식이 있는 노드만 expandItem()' 하는 방식으로 그룹만 펼칩니다.
        """
        changed = False
        was_updates_enabled = True
        try:
            root = self.fav_tree.invisibleRootItem()
            stack = [root.child(i) for i in range(root.childCount() - 1, -1, -1)]
            expanded = 0
            was_updates_enabled = self.fav_tree.updatesEnabled()
            self.fav_tree.setUpdatesEnabled(False)
            while stack:
                it = stack.pop()
                # 자식이 있으면 '그룹성 노드'로 간주
                if it.childCount() > 0:
                    if not it.isExpanded():
                        self.fav_tree.expandItem(it)
                        expanded += 1
                        changed = True
                    for j in range(it.childCount()):
                        stack.append(it.child(j))
            tag = f" reason={reason}" if reason else ""
            self._dbg_hot(f"[DBG][FAV][EXPAND_GROUPS]{tag} expanded={expanded} total_nodes={total_nodes}")
        except Exception as e:
            try:
                tag = f" reason={reason}" if reason else ""
                self._dbg_hot(f"[DBG][FAV][EXPAND_GROUPS][FAIL]{tag} {e}")
            except Exception:
                pass
        finally:
            try:
                self.fav_tree.setUpdatesEnabled(was_updates_enabled)
                if changed and was_updates_enabled:
                    self.fav_tree.viewport().update()
            except Exception:
                pass

    def _collapse_fav_groups_always(self, *, reason: str = "") -> None:
        changed = False
        was_updates_enabled = True
        try:
            root = self.fav_tree.invisibleRootItem()
            stack = [root.child(i) for i in range(root.childCount() - 1, -1, -1)]
            collapsed = 0
            was_updates_enabled = self.fav_tree.updatesEnabled()
            self.fav_tree.setUpdatesEnabled(False)
            while stack:
                it = stack.pop()
                if it.childCount() > 0:
                    for j in range(it.childCount() - 1, -1, -1):
                        stack.append(it.child(j))
                    if it.isExpanded():
                        self.fav_tree.collapseItem(it)
                        collapsed += 1
                        changed = True
            tag = f" reason={reason}" if reason else ""
            self._dbg_hot(f"[DBG][FAV][COLLAPSE_GROUPS]{tag} collapsed={collapsed}")
        except Exception as e:
            try:
                tag = f" reason={reason}" if reason else ""
                self._dbg_hot(f"[DBG][FAV][COLLAPSE_GROUPS][FAIL]{tag} {e}")
            except Exception:
                pass
        finally:
            try:
                self.fav_tree.setUpdatesEnabled(was_updates_enabled)
                if changed and was_updates_enabled:
                    self.fav_tree.viewport().update()
            except Exception:
                pass

    def _project_search_text_from_target(self, target: Any) -> List[str]:
        if not isinstance(target, dict):
            return []
        parts = []
        for key in ("notebook_text", "section_text", "display_text", "path"):
            value = target.get(key)
            if isinstance(value, str) and value.strip():
                parts.append(value.strip())
        return parts

    def _project_search_text_from_nodes(
        self, nodes: Any, max_parts: int = 1200
    ) -> List[str]:
        if not isinstance(nodes, list):
            return []
        parts = []
        stack = list(reversed(nodes))
        while stack and len(parts) < max_parts:
            node = stack.pop()
            if not isinstance(node, dict):
                continue
            name = node.get("name")
            if isinstance(name, str) and name.strip():
                parts.append(name.strip())
            parts.extend(self._project_search_text_from_target(node.get("target")))
            children = node.get("children")
            if isinstance(children, list):
                stack.extend(reversed(children))
            data = node.get("data")
            if isinstance(data, list):
                stack.extend(reversed(data))
        return parts

    def _project_search_text_from_fav_tree(self) -> List[str]:
        parts = []
        try:
            root = self.fav_tree.invisibleRootItem()
            stack = [root.child(i) for i in range(root.childCount() - 1, -1, -1)]
            while stack and len(parts) < 1200:
                item = stack.pop()
                text = item.text(0)
                if text:
                    parts.append(text)
                payload = item.data(0, ROLE_DATA) or {}
                parts.extend(self._project_search_text_from_target(payload.get("target")))
                for i in range(item.childCount() - 1, -1, -1):
                    stack.append(item.child(i))
        except Exception:
            pass
        return parts

    def _project_buffer_search_key_for_item(self, item: QTreeWidgetItem) -> str:
        if item is None:
            return ""
        parts = [item.text(0)]
        payload = item.data(0, ROLE_DATA) or {}
        item_id = payload.get("id")
        if item_id and item_id == getattr(self, "active_buffer_id", None):
            parts.extend(self._project_search_text_from_fav_tree())
        else:
            parts.extend(self._project_search_text_from_nodes(payload.get("data")))
        return _normalize_project_search_key(" ".join(p for p in parts if p))

    def _expand_buffer_groups_always(self, *, reason: str = "") -> None:
        changed = False
        was_updates_enabled = True
        try:
            root = self.buffer_tree.invisibleRootItem()
            stack = [root.child(i) for i in range(root.childCount() - 1, -1, -1)]
            expanded = 0
            was_updates_enabled = self.buffer_tree.updatesEnabled()
            self.buffer_tree.setUpdatesEnabled(False)
            while stack:
                it = stack.pop()
                if it.childCount() > 0:
                    if not it.isExpanded():
                        self.buffer_tree.expandItem(it)
                        expanded += 1
                        changed = True
                    for j in range(it.childCount()):
                        stack.append(it.child(j))
            tag = f" reason={reason}" if reason else ""
            self._dbg_hot(f"[DBG][BUF][EXPAND_GROUPS]{tag} expanded={expanded}")
        except Exception as e:
            try:
                tag = f" reason={reason}" if reason else ""
                self._dbg_hot(f"[DBG][BUF][EXPAND_GROUPS][FAIL]{tag} {e}")
            except Exception:
                pass
        finally:
            try:
                self.buffer_tree.setUpdatesEnabled(was_updates_enabled)
                if changed and was_updates_enabled:
                    self.buffer_tree.viewport().update()
            except Exception:
                pass

    def _rebuild_buffer_item_index(self) -> None:
        self._buffer_item_index = {}
        self._first_buffer_item = None
        self._buffer_search_index = []
        self._buffer_search_last_match_records = []
        try:
            root = self.buffer_tree.invisibleRootItem()
            stack = [root.child(i) for i in range(root.childCount() - 1, -1, -1)]
            while stack:
                item = stack.pop()
                payload = item.data(0, ROLE_DATA) or {}
                item_id = payload.get("id")
                if item_id:
                    self._buffer_item_index[item_id] = item
                    if (
                        item.data(0, ROLE_TYPE) == "buffer"
                        and self._first_buffer_item is None
                    ):
                        self._first_buffer_item = item
                if item.data(0, ROLE_TYPE) == "buffer":
                    search_key = self._project_buffer_search_key_for_item(item)
                    if search_key:
                        parents = []
                        parent = item.parent()
                        while parent is not None:
                            parents.append(parent)
                            parent = parent.parent()
                        self._buffer_search_index.append(
                            {"item": item, "key": search_key, "parents": tuple(parents)}
                        )
                for i in range(item.childCount() - 1, -1, -1):
                    stack.append(item.child(i))
        except Exception as e:
            self._dbg_hot(f"[DBG][BUF][INDEX][FAIL] {e}")

    def _rebuild_module_search_index(self) -> None:
        self._module_search_index = []
        self._module_search_last_match_records = []
        try:
            root = self.fav_tree.invisibleRootItem()
            stack = [root.child(i) for i in range(root.childCount() - 1, -1, -1)]
            while stack:
                item = stack.pop()
                search_key = _normalize_project_search_key(item.text(0))
                if search_key:
                    parents = []
                    parent = item.parent()
                    while parent is not None:
                        parents.append(parent)
                        parent = parent.parent()
                    self._module_search_index.append(
                        {"item": item, "key": search_key, "parents": tuple(parents)}
                    )
                for i in range(item.childCount() - 1, -1, -1):
                    stack.append(item.child(i))
        except Exception as e:
            self._dbg_hot(f"[DBG][MOD][SEARCH_INDEX][FAIL] {e}")

    def _set_buffer_search_highlight(self, item: QTreeWidgetItem, enabled: bool) -> None:
        if item is None:
            return
        try:
            if enabled:
                item.setBackground(0, self._buffer_search_highlight_bg)
                item.setForeground(0, self._buffer_search_highlight_fg)
            else:
                item.setBackground(0, self._buffer_search_clear_bg)
                item.setForeground(0, self._buffer_search_clear_fg)
        except Exception:
            pass

    def _set_module_search_highlight(self, item: QTreeWidgetItem, enabled: bool) -> None:
        self._set_buffer_search_highlight(item, enabled)

    def _clear_project_buffer_search_highlights(self) -> None:
        was_updates_enabled = True
        fav_updates_enabled = True
        changed = False
        module_changed = False
        try:
            highlighted = list(self._buffer_search_highlighted_by_id.values())
            module_highlighted = list(self._module_search_highlighted_by_id.values())
            was_updates_enabled = self.buffer_tree.updatesEnabled()
            if highlighted:
                self.buffer_tree.setUpdatesEnabled(False)
                for item in highlighted:
                    self._set_buffer_search_highlight(item, False)
                changed = True
            fav_updates_enabled = self.fav_tree.updatesEnabled()
            if module_highlighted:
                self.fav_tree.setUpdatesEnabled(False)
                for item in module_highlighted:
                    self._set_module_search_highlight(item, False)
                module_changed = True
            self._buffer_search_highlighted_by_id = {}
            self._buffer_search_last_match_records = []
            self._module_search_highlighted_by_id = {}
            self._module_search_last_match_records = []
            self._buffer_search_match_count = 0
            self._module_search_match_count = 0
            self._buffer_search_last_applied_key = ""
            self._buffer_search_last_first_match_id = 0
            self._module_search_last_first_match_id = 0
            self._buffer_search_pending_key = ""
        except Exception:
            pass
        finally:
            try:
                self.buffer_tree.setUpdatesEnabled(was_updates_enabled)
                if changed and was_updates_enabled:
                    self.buffer_tree.viewport().update()
                self.fav_tree.setUpdatesEnabled(fav_updates_enabled)
                if module_changed and fav_updates_enabled:
                    self.fav_tree.viewport().update()
            except Exception:
                pass

    def _schedule_project_buffer_search_highlight(self, text: str = "") -> None:
        self._buffer_search_pending_text = text or ""
        self._buffer_search_pending_key = _normalize_project_search_key(text)
        if not self._buffer_search_pending_key:
            self._buffer_search_timer.stop()
            self._clear_project_buffer_search_highlights()
            return
        self._buffer_search_timer.start(35)

    def _apply_project_search_to_tree(
        self,
        *,
        tree: QTreeWidget,
        records: List[Dict[str, Any]],
        previous_records: List[Dict[str, Any]],
        previous_highlights: Dict[int, QTreeWidgetItem],
        query: str,
        previous_query: str,
        set_highlight: Callable[[QTreeWidgetItem, bool], None],
        previous_first_match_id: int,
        scroll: bool,
    ) -> Dict[str, Any]:
        records_to_scan = records
        if previous_query and query.startswith(previous_query):
            records_to_scan = previous_records

        match_count = 0
        matched_records = []
        match_by_id: Dict[int, QTreeWidgetItem] = {}
        first_match = None
        changed_highlights = False
        was_updates_enabled = True
        try:
            was_updates_enabled = tree.updatesEnabled()
            for record in records_to_scan:
                item = record.get("item")
                item_key = record.get("key") or ""
                if item is not None and item_key and query in item_key:
                    match_count += 1
                    matched_records.append(record)
                    if first_match is None:
                        first_match = item
                    match_by_id[id(item)] = item

            prev_ids = set(previous_highlights)
            next_ids = set(match_by_id)
            removed_ids = prev_ids - next_ids
            added_ids = next_ids - prev_ids
            changed_highlights = bool(removed_ids or added_ids)

            if changed_highlights:
                tree.setUpdatesEnabled(False)
                for item_id in removed_ids:
                    set_highlight(previous_highlights.get(item_id), False)
                expanded_parent_ids = set()
                for record in matched_records:
                    item = record.get("item")
                    if item is None or id(item) not in added_ids:
                        continue
                    set_highlight(item, True)
                    for parent in record.get("parents") or ():
                        parent_id = id(parent)
                        if parent_id in expanded_parent_ids:
                            continue
                        expanded_parent_ids.add(parent_id)
                        if not parent.isExpanded():
                            tree.expandItem(parent)
        finally:
            try:
                tree.setUpdatesEnabled(was_updates_enabled)
                if was_updates_enabled and changed_highlights:
                    tree.viewport().update()
            except Exception:
                pass

        first_match_id = id(first_match) if first_match is not None else 0
        if (
            scroll
            and first_match is not None
            and first_match_id != previous_first_match_id
        ):
            try:
                tree.scrollToItem(
                    first_match,
                    QAbstractItemView.ScrollHint.EnsureVisible,
                )
            except Exception:
                pass

        return {
            "matched_records": matched_records,
            "match_by_id": match_by_id,
            "match_count": match_count,
            "first_match_id": first_match_id,
            "changed": changed_highlights,
        }

    def _highlight_project_buffers_from_module_search(
        self,
        text: str = "",
        *,
        update_status: bool = True,
        scroll: bool = True,
        precomputed_query: Optional[str] = None,
    ) -> None:
        query = precomputed_query if precomputed_query is not None else _normalize_project_search_key(text)
        if not query:
            self._clear_project_buffer_search_highlights()
            return

        if not self._buffer_search_index:
            self._rebuild_buffer_item_index()
        if not self._module_search_index:
            self._rebuild_module_search_index()

        previous_query = self._buffer_search_last_applied_key
        if query == previous_query:
            return
        self._buffer_search_last_applied_key = query

        buffer_result = self._apply_project_search_to_tree(
            tree=self.buffer_tree,
            records=self._buffer_search_index,
            previous_records=getattr(self, "_buffer_search_last_match_records", []),
            previous_highlights=self._buffer_search_highlighted_by_id,
            query=query,
            previous_query=previous_query,
            set_highlight=self._set_buffer_search_highlight,
            previous_first_match_id=self._buffer_search_last_first_match_id,
            scroll=scroll,
        )
        module_result = self._apply_project_search_to_tree(
            tree=self.fav_tree,
            records=self._module_search_index,
            previous_records=getattr(self, "_module_search_last_match_records", []),
            previous_highlights=self._module_search_highlighted_by_id,
            query=query,
            previous_query=previous_query,
            set_highlight=self._set_module_search_highlight,
            previous_first_match_id=self._module_search_last_first_match_id,
            scroll=scroll,
        )

        self._buffer_search_last_match_records = buffer_result["matched_records"]
        self._buffer_search_highlighted_by_id = buffer_result["match_by_id"]
        self._buffer_search_match_count = buffer_result["match_count"]
        self._buffer_search_last_first_match_id = buffer_result["first_match_id"]
        self._module_search_last_match_records = module_result["matched_records"]
        self._module_search_highlighted_by_id = module_result["match_by_id"]
        self._module_search_match_count = module_result["match_count"]
        self._module_search_last_first_match_id = module_result["first_match_id"]

        if update_status:
            try:
                self.connection_status_label.setText(
                    f"프로젝트 검색: '{text}' - 프로젝트 {self._buffer_search_match_count}개, 모듈 {self._module_search_match_count}개 강조"
                )
            except Exception:
                pass

    def _refresh_project_buffer_search_highlights(self) -> None:
        search_input = getattr(self, "module_project_search_input", None)
        if search_input is None:
            return
        query = _normalize_project_search_key(search_input.text())
        if not query:
            if (
                self._buffer_search_highlighted_by_id
                or self._module_search_highlighted_by_id
            ):
                self._clear_project_buffer_search_highlights()
            return
        self._buffer_search_last_applied_key = ""
        self._highlight_project_buffers_from_module_search(
            search_input.text(),
            update_status=False,
            scroll=False,
            precomputed_query=query,
        )

    def _load_favorites_into_center_tree(self, node_data: List):
        """즐겨찾기 데이터를 중앙 트리에 로드합니다."""
        # ✅ 동일 데이터면 rebuild 스킵 (클릭 렉 제거 핵심)
        payload_raw = None
        source_id = id(node_data) if isinstance(node_data, list) else 0
        try:
            payload_raw = json.dumps(node_data, sort_keys=True, ensure_ascii=False)
            new_hash = hashlib.md5(payload_raw.encode("utf-8")).hexdigest()
        except Exception:
            new_hash = None

        if new_hash is not None and getattr(self, "_last_center_payload_hash", None) == new_hash:
            self._last_center_payload_snapshot = payload_raw
            self._last_center_payload_source_id = source_id
            return

        # 로딩 중 clear/append 과정에서 structureChanged/itemChanged가 발생하면
        # 선택 버퍼가 바뀌는 타이밍에 "빈 데이터"가 저장되는 문제가 발생할 수 있다.
        # (재현: 버퍼 A에서 섹션 추가 → 버퍼 B 클릭 → 다시 A 클릭 시 A가 빈 목록으로 덮임)
        self.fav_tree.blockSignals(True)
        was_updates_enabled = self.fav_tree.updatesEnabled()
        self.fav_tree.setUpdatesEnabled(False)
        try:
            self.fav_tree.clear()
            self._module_search_index = []
            self._module_search_last_match_records = []
            self._module_search_highlighted_by_id = {}
            self._module_search_match_count = 0
            t_build0 = time.perf_counter()
            for node in node_data:
                self._append_fav_node(self.fav_tree.invisibleRootItem(), node)

            build_ms = (time.perf_counter() - t_build0) * 1000.0
            total_nodes = -1
            if self._debug_perf_logs or self._debug_hotpaths:
                total_nodes = self._count_nodes_recursive(node_data)
            self._dbg_perf(
                f"[BOOT][PERF][FAV_REBUILD] total_nodes={total_nodes} build_ms={build_ms:.1f}"
            )

            # ✅ 2패널은 '그룹이 항상 펼쳐진 상태'가 기본 UX
            #    - 예전에는 최초 1회만 펼쳤는데, 그 이후엔 항상 접힌 상태로 복원되어 사용성이 나빠짐
            #    - expandAll() 대신 '자식이 있는 노드만 expandItem' 방식으로 그룹만 펼쳐 성능도 방어
            self._expand_fav_groups_always(total_nodes=total_nodes, reason="rebuild")
            self._rebuild_module_search_index()
            self._last_center_payload_hash = new_hash
            self._last_center_payload_snapshot = payload_raw
            self._last_center_payload_source_id = source_id
        finally:
            self.fav_tree.setUpdatesEnabled(was_updates_enabled)
            self.fav_tree.blockSignals(False)
            if was_updates_enabled:
                self.fav_tree.viewport().update()
        self._refresh_project_buffer_search_highlights()

    def _request_favorites_save(self, *_args):
        """
        FavoritesTree 변경 신호 폭주를 짧게 묶어 1회 저장으로 합칩니다.
        - 이름변경(itemChanged 연속)
        - 드래그/구조변경(structureChanged + itemChanged 연쇄)
        - 붙여넣기 직후의 다중 signal
        """
        if not self.active_buffer_node:
            return
        if getattr(self, "_fav_undo_suspended", False):
            return
        self._module_search_index = []
        self._module_search_last_match_records = []
        self._buffer_search_last_applied_key = ""
        self._fav_save_pending = True
        self._fav_save_timer.start(self._fav_save_interval_ms)

    def _flush_pending_favorites_save(self):
        if not getattr(self, "_fav_save_pending", False):
            return False
        self._fav_save_pending = False
        self._save_favorites()
        return True

    def _save_favorites(self):
        """현재 활성화된 중앙 트리의 내용을 버퍼 트리의 해당 노드 데이터에 반영하고 저장합니다."""
        if not self.active_buffer_node:
            return
        self._invalidate_aggregate_cache(
            invalidate_classified_keys=self.active_buffer_id != AGG_BUFFER_ID
        )
        self._module_search_index = []
        self._module_search_last_match_records = []
        self._buffer_search_index = []
        self._buffer_search_last_match_records = []
        self._buffer_search_last_applied_key = ""

        try:
            if getattr(self, "_fav_save_timer", None) is not None and self._fav_save_timer.isActive():
                self._fav_save_timer.stop()
        except Exception:
            pass
        self._fav_save_pending = False

        # ✅ 종합 버퍼도 이제 '노트북 저장'을 위해 저장 허용

        try:
            data = []
            root = self.fav_tree.invisibleRootItem()
            for i in range(root.childCount()):
                data.append(self._serialize_fav_item(root.child(i)))

            try:
                snap = json.dumps(data, sort_keys=True, ensure_ascii=False)
            except Exception:
                snap = "[]"
            try:
                current_hash = None
                if snap == getattr(self, "_last_center_payload_snapshot", None):
                    current_hash = getattr(self, "_last_center_payload_hash", None)
                if not current_hash:
                    current_hash = hashlib.md5(snap.encode("utf-8")).hexdigest()
            except Exception:
                current_hash = None

            # ✅ IMPORTANT:
            # 중앙 트리는 _load_favorites_into_center_tree()에서만 해시를 갱신한다는 가정이 있었는데,
            # 실제로는 paste/cut/delete 등으로 트리를 "직접" 수정한다.
            # 그러면 _last_center_payload_hash가 과거 상태(예: 7개)의 해시로 남아,
            # Undo가 과거 스냅샷(7개)을 로드하려 할 때 "해시가 같다"로 판단되어 리빌드가 스킵된다.
            # => Ctrl+Z가 안 먹는 것처럼 보이는 핵심 원인.
            self._last_center_payload_hash = current_hash
            self._last_center_payload_snapshot = snap
            self._last_center_payload_source_id = id(data)

            # ✅ 같은 상태를 signal 연쇄로 다시 저장하려는 경우 early-return
            #    - rename 편집 종료 / itemChanged 연속 호출 / 구조변경 후 중복 저장 방어
            #    - Undo/Redo 스냅샷은 동일 상태면 원래도 skip이므로 여기서 바로 끊어도 안전
            if current_hash is not None and current_hash == getattr(self, "_fav_last_persisted_hash", None):
                return

            # --- Undo/Redo: FavoritesTree 변경 스냅샷 기록 ---
            try:
                self._fav_record_snapshot(snap)
            except Exception:
                pass

            # 메모리 상의 active_buffer_node 데이터 업데이트
            if self.active_buffer_node is not None:
                self.active_buffer_node["data"] = data
                self._dbg_hot(f"[DBG][FAV][SAVE] Updated active_buffer_node data: count={len(data)}")

            # PyQt의 item.data()로 얻은 dict는 "수정해도 item 내부에 반영되지" 않는 경우가 있다.
            # 따라서 활성 버퍼의 QTreeWidgetItem에도 동일 데이터를 강제 주입한다.
            if self.active_buffer_item is None and self.active_buffer_id:
                self.active_buffer_item = self._buffer_item_index.get(self.active_buffer_id)
            if self.active_buffer_item is None and self.active_buffer_id:
                # 예외 상황 대비: ID로 다시 찾아서 연결
                iterator = QTreeWidgetItemIterator(self.buffer_tree)
                while iterator.value():
                    it = iterator.value()
                    p = it.data(0, ROLE_DATA) or {}
                    if p.get("id") == self.active_buffer_id:
                        self.active_buffer_item = it
                        break
                    iterator += 1

            if self.active_buffer_item is not None:
                p = self.active_buffer_item.data(0, ROLE_DATA) or {}
                p["data"] = data
                self.active_buffer_item.setData(0, ROLE_DATA, p)
                self._dbg_hot(f"[DBG][FAV][SAVE] Updated active_buffer_item payload: id={p.get('id')}")

            self._fav_last_persisted_hash = current_hash
            # 그리고 전체 버퍼 구조 저장
            settings_buffer = self._active_buffer_settings_node
            if settings_buffer is None:
                settings_buffer = _find_buffer_node_by_id(
                    self.settings.get("favorites_buffers", []),
                    self.active_buffer_id,
                )
                self._active_buffer_settings_node = settings_buffer
            if settings_buffer is not None:
                settings_buffer["data"] = data
                self.settings["active_buffer_id"] = self.active_buffer_id
                self._save_settings_to_file()
                self._dbg_hot("[DBG][FAV][SAVE] Active buffer data persisted without full structure rebuild")
            else:
                self._save_buffer_structure()
                self._dbg_hot("[DBG][FAV][SAVE] Fallback full buffer structure persist")

            if (
                self.active_buffer_id == AGG_BUFFER_ID
                and not getattr(self, "_aggregate_reclassify_in_progress", False)
            ):
                self._refresh_active_aggregate_classification_from_saved_data(
                    current_nodes=data,
                    persist=True,
                    show_status=False,
                )

        except Exception as e:
            print(f"[ERROR] 즐겨찾기 저장 실패: {e}")

    def _request_buffer_structure_save(self, *_args):
        if getattr(self, "_boot_loading", False):
            return
        self._buffer_save_timer.start(self._buffer_save_interval_ms)

    def _flush_pending_buffer_structure_save(self) -> None:
        try:
            if self._buffer_save_timer.isActive():
                self._buffer_save_timer.stop()
                self._save_buffer_structure()
        except Exception:
            pass

    def _save_buffer_structure(self):
        """버퍼 트리의 구조(그룹/버퍼)를 settings에 저장합니다."""
        self._invalidate_aggregate_cache()
        try:
            if self._buffer_save_timer.isActive():
                self._buffer_save_timer.stop()
        except Exception:
            pass
        self._buffer_item_index = {}
        self._first_buffer_item = None
        self._buffer_search_index = []
        self._buffer_search_last_match_records = []
        root = self.buffer_tree.invisibleRootItem()
        structure = []
        for i in range(root.childCount()):
            structure.append(self._serialize_buffer_item(root.child(i), rebuild_index=True))

        try:
            structure_sig = hashlib.md5(
                json.dumps(structure, sort_keys=True, ensure_ascii=False).encode("utf-8")
            ).hexdigest()
        except Exception:
            structure_sig = None

        if structure_sig is not None and structure_sig == getattr(self, "_last_saved_buffer_structure_sig", None):
            self.settings["favorites_buffers"] = structure
            self._active_buffer_settings_node = _find_buffer_node_by_id(
                self.settings.get("favorites_buffers", []),
                self.active_buffer_id,
            )
            self._refresh_project_buffer_search_highlights()
            return

        self.settings["favorites_buffers"] = structure
        # ✅ 저장 직전에 구조 강제 보정(순서/락/종합 유지)
        _ensure_default_and_aggregate_inplace(self.settings)
        self._last_saved_buffer_structure_sig = structure_sig
        self._active_buffer_settings_node = _find_buffer_node_by_id(
            self.settings.get("favorites_buffers", []),
            self.active_buffer_id,
        )
        self._save_settings_to_file()
        self._refresh_project_buffer_search_highlights()

    def _serialize_buffer_item(
        self,
        item: QTreeWidgetItem,
        *,
        rebuild_index: bool = False,
    ) -> Dict:
        node_type = item.data(0, ROLE_TYPE)
        payload = item.data(0, ROLE_DATA) or {}
        if rebuild_index:
            item_id = payload.get("id")
            if item_id:
                self._buffer_item_index[item_id] = item
                if node_type == "buffer" and self._first_buffer_item is None:
                    self._first_buffer_item = item
            if node_type == "buffer":
                search_key = self._project_buffer_search_key_for_item(item)
                if search_key:
                    parents = []
                    parent_item = item.parent()
                    while parent_item is not None:
                        parents.append(parent_item)
                        parent_item = parent_item.parent()
                    self._buffer_search_index.append(
                        {"item": item, "key": search_key, "parents": tuple(parents)}
                    )

        node = {
            "type": node_type,
            "id": payload.get("id"),
            "name": item.text(0)
        }

        if node_type == "group":
            if payload.get("locked"):
                node["locked"] = True
            children = []
            for i in range(item.childCount()):
                children.append(
                    self._serialize_buffer_item(
                        item.child(i),
                        rebuild_index=rebuild_index,
                    )
                )
            node["children"] = children
        else:
            # 버퍼인 경우, 현재 메모리 상의 데이터를 유지하거나
            # 활성 상태라면 현재 중앙 트리에서 가져와야 함.
            # payload['data']는 로드 시점의 스냅샷일 수 있으므로 주의.
            # 여기서는 payload['data']를 그대로 쓰고,
            # 활성 버퍼가 변경될 때마다 payload['data']를 갱신해두는 방식을 사용.
            if payload.get("virtual"):
                node["virtual"] = payload.get("virtual")
            if payload.get("locked"):
                node["locked"] = True

            node["data"] = payload.get("data", [])
            # [DBG] 종합 버퍼 저장 스캔
            if node.get("id") == AGG_BUFFER_ID:
                self._dbg_hot(f"[DBG][SSOT][SERIALIZE] Aggregate data count={len(node['data'])}")

        return node

    def _request_settings_save(self):
        self._settings_save_pending = True
        self._settings_save_timer.start(self._settings_save_interval_ms)

    def _flush_pending_settings_save(self):
        if self._settings_save_in_progress:
            return
        if not self._settings_save_pending and self._settings_save_timer.isActive():
            self._settings_save_timer.stop()
        self._settings_save_pending = False
        self._settings_save_in_progress = True
        try:
            save_settings(self.settings)
        finally:
            self._settings_save_in_progress = False

    def _save_settings_to_file(self, immediate: bool = False):
        """현재 self.settings 객체를 파일에 저장합니다."""
        if immediate:
            self._flush_pending_settings_save()
        else:
            self._request_settings_save()


    def _backup_full_settings(self):
        """전체 설정을 백업합니다."""
        last_dir = self.settings.get("last_backup_dir", os.getcwd())
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        default_filename = f"OneNote_Remocon_Backup_{timestamp}.json"

        default_path = os.path.join(last_dir, default_filename)

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "설정 백업",
            default_path,
            "JSON Files (*.json);;All Files (*)",
        )

        if file_path:
            try:
                # 현재 메모리 설정을 파일로 강제 저장 (최신 상태 반영)
                self._save_window_state() # 창 위치 등 업데이트
                self._flush_pending_buffer_structure_save()
                self._save_favorites()    # 즐겨찾기 업데이트

                # 백업 디렉토리 기억
                self.settings["last_backup_dir"] = os.path.dirname(file_path)

                # _write_json을 사용하여 안전하게 저장
                _write_json(file_path, self.settings)

                # 설정 파일에도 last_backup_dir 반영하여 저장
                self._save_settings_to_file(immediate=True)

                QMessageBox.information(
                    self, "백업 완료", f"성공적으로 백업되었습니다.\n\n경로: {file_path}"
                )
            except Exception as e:
                QMessageBox.critical(
                    self, "백업 실패", f"백업 중 오류가 발생했습니다:\n{e}"
                )

    def _reload_settings_after_path_change(self) -> None:
        self.settings = load_settings()
        try:
            self._load_buffers_and_favorites()
            self._update_move_button_state()
        except Exception as e:
            print(f"[WARN] 설정 경로 변경 후 UI 재로드 실패: {e}")
        try:
            self.connection_status_label.setText(_settings_path_mode_label())
        except Exception:
            pass

    def _choose_shared_settings_json(self):
        """프로젝트 실행과 EXE 실행이 같이 쓸 공용 설정 JSON을 지정합니다."""
        current_path = _get_settings_file_path()
        default_dir = os.path.dirname(current_path) if current_path else os.getcwd()
        default_path = os.path.join(default_dir, SETTINGS_FILE)

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "공용 설정 JSON 위치 지정",
            default_path,
            "JSON Files (*.json);;All Files (*)",
        )
        if not file_path:
            return

        file_path = _expand_external_settings_path(file_path)
        try:
            self._save_window_state()
            self._flush_pending_buffer_structure_save()
            self._save_favorites()

            use_existing = False
            if os.path.exists(file_path):
                answer = QMessageBox.question(
                    self,
                    "공용 설정 JSON 선택",
                    (
                        "선택한 JSON 파일이 이미 있습니다.\n\n"
                        "예: 기존 JSON을 불러와 사용합니다.\n"
                        "아니오: 현재 앱 설정으로 해당 JSON을 덮어씁니다.\n"
                        "취소: 변경하지 않습니다."
                    ),
                    (
                        QMessageBox.StandardButton.Yes
                        | QMessageBox.StandardButton.No
                        | QMessageBox.StandardButton.Cancel
                    ),
                    QMessageBox.StandardButton.Yes,
                )
                if answer == QMessageBox.StandardButton.Cancel:
                    return
                use_existing = answer == QMessageBox.StandardButton.Yes

            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            if not use_existing:
                payload = dict(self.settings)
                payload.pop("favorites", None)
                _ensure_default_and_aggregate_inplace(payload)
                _write_json(file_path, payload)

            _set_external_settings_file_path(file_path)
            _SETTINGS_OBJECT_CACHE.clear()
            _JSON_TEXT_CACHE.clear()
            self._reload_settings_after_path_change()
            QMessageBox.information(
                self,
                "공용 설정 JSON 연결 완료",
                (
                    "이제 프로젝트 실행과 EXE 실행이 같은 설정 JSON을 사용합니다.\n\n"
                    f"경로: {file_path}\n\n"
                    "다른 PC에서는 OneDrive가 동기화된 로컬 경로를 한 번 지정하면 됩니다."
                ),
            )
        except Exception as e:
            QMessageBox.critical(self, "공용 설정 JSON 연결 실패", str(e))

    def _show_settings_json_path(self):
        QMessageBox.information(
            self,
            "현재 설정 JSON 위치",
            (
                f"{_settings_path_mode_label()}\n\n"
                f"공용 경로 포인터: {_settings_path_config_file()}\n"
                f"환경변수 우선순위: {SETTINGS_PATH_ENV}"
            ),
        )

    def _open_settings_json_folder(self):
        path = _get_settings_file_path()
        folder = os.path.dirname(path)
        try:
            os.makedirs(folder, exist_ok=True)
            os.startfile(folder)
        except Exception as e:
            QMessageBox.warning(self, "폴더 열기 실패", str(e))

    def _clear_shared_settings_json(self):
        if not _get_external_settings_file_path():
            QMessageBox.information(self, "안내", "현재 공용 설정 JSON 연결이 없습니다.")
            return
        answer = QMessageBox.question(
            self,
            "공용 설정 JSON 연결 해제",
            (
                "공용 설정 JSON 연결을 해제하고 기본 위치의 설정 JSON을 사용합니다.\n"
                "현재 공용 JSON 파일 자체는 삭제하지 않습니다.\n\n"
                "계속하시겠습니까?"
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            self._save_settings_to_file(immediate=True)
            _clear_external_settings_file_path()
            _SETTINGS_OBJECT_CACHE.clear()
            _JSON_TEXT_CACHE.clear()
            self._reload_settings_after_path_change()
            QMessageBox.information(
                self,
                "연결 해제 완료",
                f"이제 기본 설정 JSON을 사용합니다.\n\n{_get_settings_file_path()}",
            )
        except Exception as e:
            QMessageBox.critical(self, "연결 해제 실패", str(e))

    def _restore_full_settings(self):
        """설정 파일을 복원합니다."""
        last_dir = self.settings.get("last_backup_dir", os.getcwd())

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "설정 복원",
            last_dir,
            "JSON Files (*.json);;All Files (*)",
        )

        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    restored_data = json.load(f)

                # 마이그레이션 적용 (구버전 백업일 경우 대비)
                if _migrate_favorites_buffers_inplace(restored_data):
                    print("[INFO] 복원 데이터 마이그레이션 수행됨")

                confirm = QMessageBox.question(
                    self,
                    "복원 확인",
                    "설정을 복원하면 현재 설정이 덮어씌워집니다.\n계속하시겠습니까?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )

                if confirm != QMessageBox.StandardButton.Yes:
                    return

                # 복원 디렉토리 기억
                restored_data["last_backup_dir"] = os.path.dirname(file_path)

                # 현재 설정 교체
                self.settings = DEFAULT_SETTINGS.copy()
                self.settings.update(restored_data)

                # 파일에 즉시 반영
                self._save_settings_to_file(immediate=True)

                # UI 리로드
                # 1. 버퍼/즐겨찾기 트리 갱신
                self._load_buffers_and_favorites()

                # 2. 스플리터 위치 등은 재시작 시 적용되거나 지금 강제 적용 가능하나
                # 여기서는 데이터 갱신에 집중

                QMessageBox.information(
                    self, "복원 완료", "설정이 복원되었습니다."
                )

            except Exception as e:
                QMessageBox.critical(
                    self, "복원 실패", f"복원 중 오류가 발생했습니다:\n{e}"
                )

    def _serialize_fav_item(self, item: QTreeWidgetItem) -> Dict[str, Any]:
        node_type = item.data(0, ROLE_TYPE)
        payload = item.data(0, ROLE_DATA) or {}
        node = {
            "type": node_type,
            "id": payload.get("id") or str(uuid.uuid4()),
            "name": item.text(0),
        }
        if node_type in ("section", "notebook"):
            node["target"] = payload.get("target", {})
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
        if node_type in ("section", "notebook"):
            payload["target"] = node.get("target", {})
            icon = getattr(self, "_icon_file", None) or self.style().standardIcon(QApplication.style().StandardPixmap.SP_FileIcon)
            item.setIcon(0, icon)
        else:
            icon = getattr(self, "_icon_dir", None) or self.style().standardIcon(QApplication.style().StandardPixmap.SP_DirIcon)
            item.setIcon(0, icon)
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

    def _dbg_node_type_counts(self, nodes, tag=""):
        try:
            if not self._debug_hotpaths:
                return
            cnt = {"notebook": 0, "section": 0, "page": 0, "group": 0, "buffer": 0, "unknown": 0}
            def rec(arr):
                if not isinstance(arr, list):
                    return
                for n in arr:
                    if not isinstance(n, dict): continue
                    ty = n.get("type", "unknown")
                    cnt[ty] = cnt.get(ty or "unknown", 0) + 1
                    rec(n.get("children", []))
                    rec(n.get("data", []))
            rec(nodes)
            self._dbg_hot(f"[DBG][NODE_TYPES]{'['+tag+']' if tag else ''} total={len(nodes) if isinstance(nodes, list) else 'NA'} {cnt}")
        except Exception as e:
            print(f"[DBG][NODE_TYPES][FAIL] {e}")

    def _nodes_have_type(self, nodes, ty: str) -> bool:
        if not isinstance(nodes, list):
            return False
        for n in nodes:
            if isinstance(n, dict) and n.get("type") == ty:
                return True
        return False

    def _nodes_have_any_type(self, nodes, types) -> bool:
        return isinstance(nodes, list) and any(isinstance(n, dict) and n.get("type") in types for n in nodes)

    def _count_nodes_recursive(self, nodes) -> int:
        """node(list[dict])의 총 노드 수(자식 포함)를 빠르게 계산한다."""
        if not isinstance(nodes, list):
            return 0
        total = 0
        stack = list(nodes)
        while stack:
            n = stack.pop()
            if not isinstance(n, dict):
                continue
            total += 1
            ch = n.get("children")
            if isinstance(ch, list) and ch:
                stack.extend(ch)
        return total

    def _sorted_copy_nodes_by_name(self, nodes: Any) -> List:
        """
        종합(aggregate) 버퍼에서만 사용하는 표시용 정렬.
        - 원본 nodes를 변형하지 않기 위해 deepcopy 후 정렬
        - group이 있으면 children도 재귀 정렬
        """
        if not isinstance(nodes, list):
            return nodes if isinstance(nodes, list) else (nodes or [])
        try:
            copied = copy.deepcopy(nodes)
        except Exception:
            copied = list(nodes)

        def _disp_name(n: Any) -> str:
            if not isinstance(n, dict):
                return ""
            # group / buffer / section / notebook 모두 name 우선
            name = n.get("name")
            if name:
                return name
            t = n.get("target") or {}
            return t.get("section_text") or t.get("notebook_text") or ""

        def _rec(lst: List) -> List:
            # children 먼저 정렬
            for it in lst:
                if isinstance(it, dict) and isinstance(it.get("children"), list):
                    it["children"] = _rec(it["children"])
            try:
                lst.sort(key=lambda n: _name_sort_key(_disp_name(n)))
            except Exception:
                pass
            return lst

        return _rec(copied)

    # ----------------- FavoritesTree Undo/Redo helpers -----------------
    def _fav_reset_undo_context_from_data(self, data, *, reason: str = "") -> None:
        """
        2패널(모듈영역) Undo/Redo가 이상해지는 핵심 원인은
        '버퍼 전환 후에도 이전 버퍼의 _fav_last_snapshot / undo stack이 유지'되어
        Ctrl+Z가 다른 버퍼 스냅샷을 불러오는 케이스가 생기는 것입니다.

        버퍼/그룹 선택으로 중앙 트리를 로드한 직후,
        로드된 데이터 기준으로 undo/redo 컨텍스트를 초기화합니다.

        - _fav_last_snapshot: 현재 상태(초기 스냅샷)
        - _fav_undo_stack / _fav_redo_stack: 비움

        이렇게 해야 첫 변경 저장 시 "초기 스냅샷"이 undo에 들어가며,
        Ctrl+Z가 현재 버퍼 내부에서만 정상 동작합니다.
        """
        try:
            data_list = data if isinstance(data, list) else []
            snap = None
            if (
                isinstance(data, list)
                and id(data) == getattr(self, "_last_center_payload_source_id", 0)
            ):
                snap = getattr(self, "_last_center_payload_snapshot", None)
            if snap is None:
                snap = json.dumps(data_list, sort_keys=True, ensure_ascii=False)
        except Exception:
            snap = ""
        try:
            self._fav_undo_stack.clear()
            self._fav_redo_stack.clear()
        except Exception:
            self._fav_undo_stack = []
            self._fav_redo_stack = []
        self._fav_last_snapshot = snap
        try:
            self._fav_last_persisted_hash = None
            if snap == getattr(self, "_last_center_payload_snapshot", None):
                self._fav_last_persisted_hash = getattr(
                    self, "_last_center_payload_hash", None
                )
            if not self._fav_last_persisted_hash:
                self._fav_last_persisted_hash = hashlib.md5(
                    snap.encode("utf-8")
                ).hexdigest()
        except Exception:
            self._fav_last_persisted_hash = None
        try:
            tag = f" reason={reason}" if reason else ""
            self._dbg_hot(f"[DBG][FAV][UNDO_CTX] reset{tag} undo=0 redo=0 snap_len={len(snap)}")
        except Exception:
            pass

    # ----------------- 15-3. 버퍼 트리 이벤트 핸들러 -----------------
    def _on_buffer_tree_item_clicked(self, item, col):
        """버퍼 트리 항목 클릭 시 처리"""
        if not item:
            return
        node_type = item.data(0, ROLE_TYPE)
        payload = item.data(0, ROLE_DATA) or {}

        if node_type == "buffer":
            # 버퍼 전환 직전: 현재 중앙 트리 내용을 "이전 버퍼"에 반드시 저장
            # (그렇지 않으면 버퍼를 다시 클릭했을 때 섹션/그룹이 사라지는 현상 발생)
            flushed_current_buffer = False
            try:
                flushed_current_buffer = self._flush_pending_favorites_save()
            except Exception:
                pass
            if (
                self.active_buffer_id
                and payload
                and payload.get("id") != self.active_buffer_id
                and not flushed_current_buffer
            ):
                self._save_favorites()

            self.active_buffer_id = payload.get("id")
            self.active_buffer_node = payload  # Dict payload(스냅샷)
            self.active_buffer_item = item
            self._active_buffer_settings_node = _find_buffer_node_by_id(
                self.settings.get("favorites_buffers", []),
                self.active_buffer_id,
            )
            self.settings["active_buffer_id"] = self.active_buffer_id

            try:
                # ✅ 종합 버퍼: 전자필기장(노트북) 저장용
                if payload.get("id") == AGG_BUFFER_ID:
                    saved = payload.get("data", []) or []
                    self._dbg_node_type_counts(saved, "AGG_SAVED")

                    # 1) 종합 버퍼에 notebook/group 저장이 있으면: 그걸 기준으로 분류 표시한다.
                    if self._nodes_have_any_type(saved, {"notebook", "group"}):
                        self._dbg_hot(f"[DBG][AGG] load SAVED data (len={len(saved)})")
                        data_to_load = self._build_aggregate_categorized_display_nodes(saved)
                        self._load_favorites_into_center_tree(data_to_load)
                    else:
                        # 2) 저장된게 없을 때만: 기존 종합 계산 fallback
                        agg_data = self._build_aggregate_buffer()
                        self._dbg_node_type_counts(agg_data, "AGG_BUILT")
                        self._dbg_hot(f"[DBG][AGG] load BUILT aggregate (len={len(agg_data)})")
                        data_to_load = self._build_aggregate_categorized_display_nodes(agg_data)
                        self._load_favorites_into_center_tree(data_to_load)

                    # ✅ 버퍼 전환 직후: Undo/Redo 컨텍스트를 현재 버퍼 데이터로 리셋
                    self._fav_reset_undo_context_from_data(data_to_load, reason="buffer_switch:AGG")
                    self.btn_add_section_current.setEnabled(False)
                    self.btn_add_group.setEnabled(True)  # 원하면 그룹도 허용
                    if hasattr(self, "btn_register_all_notebooks"):
                        self.btn_register_all_notebooks.setEnabled(True)
                        self.btn_register_all_notebooks.setVisible(True)
                else:
                    data_to_load = payload.get("data", []) or []
                    self._load_favorites_into_center_tree(data_to_load)

                    # ✅ 버퍼 전환 직후: Undo/Redo 컨텍스트를 현재 버퍼 데이터로 리셋
                    self._fav_reset_undo_context_from_data(data_to_load, reason="buffer_switch")
                    self.btn_add_section_current.setEnabled(True)
                    if hasattr(self, "btn_register_all_notebooks"):
                        self.btn_register_all_notebooks.setEnabled(False)
                        self.btn_register_all_notebooks.setVisible(False)
                    self.btn_add_group.setEnabled(True)
                self._last_loaded_center_buffer_id = self.active_buffer_id
            finally:
                try:
                    self.buffer_tree.setUpdatesEnabled(True)
                    self.fav_tree.setUpdatesEnabled(True)
                    self.buffer_tree.viewport().update()
                    self.fav_tree.viewport().update()
                except Exception:
                    pass
        else:
            # 그룹 선택 시
            # 현재 버퍼 내용이 남아있을 수 있으므로 먼저 저장
            flushed_current_buffer = False
            try:
                flushed_current_buffer = self._flush_pending_favorites_save()
            except Exception:
                pass
            if self.active_buffer_id and not flushed_current_buffer:
                self._save_favorites()
            if hasattr(self, "btn_register_all_notebooks"):
                self.btn_register_all_notebooks.setEnabled(False)
                self.btn_register_all_notebooks.setVisible(False)
            self.btn_add_section_current.setEnabled(False)
            self.btn_add_group.setEnabled(False)
            self.active_buffer_id = None
            self.active_buffer_item = None
            self._active_buffer_settings_node = None
            self._last_loaded_center_buffer_id = None
            self._load_favorites_into_center_tree([])

            # ✅ 버퍼가 아닌(그룹/빈) 상태에서도 Undo 컨텍스트를 리셋 (이전 버퍼 스냅샷 혼입 방지)
            self._fav_reset_undo_context_from_data([], reason="buffer_switch:group_or_none")
        self._update_buffer_move_button_state()

    def _on_buffer_tree_selection_changed(self):
        """
        1패널에서 클릭/키보드 이동 등으로 "선택"만 바뀐 경우에도
        2패널(모듈/섹션)이 즉시 갱신되도록 한다.
        """
        if getattr(self, "_buf_sel_guard", False):
            return
        item = self.buffer_tree.currentItem()
        if not item:
            return

        node_type = item.data(0, ROLE_TYPE)
        payload = item.data(0, ROLE_DATA) or {}

        # 이미 활성 버퍼면 스킵(불필요 리로드 방지)
        if node_type == "buffer":
            cur_id = payload.get("id")
            if cur_id and self.active_buffer_id == cur_id:
                return

        self._buf_sel_guard = True
        try:
            # 기존 클릭 로직 재사용
            self._on_buffer_tree_item_clicked(item, 0)
        finally:
            self._buf_sel_guard = False

    def _on_buffer_tree_double_clicked(self, item, col):
        node_type = item.data(0, ROLE_TYPE)
        if node_type == "group":
            # 그룹이면 확장/축소 (기본 동작)
            pass
        else:
            # 버퍼면 이름 편집
            pass

    def _add_buffer_group(self):
        """새 버퍼 그룹 추가"""
        parent = self.buffer_tree.currentItem()
        # 버퍼가 선택되어 있으면 그 부모(그룹 또는 루트)에 추가
        if parent and parent.data(0, ROLE_TYPE) == "buffer":
            parent = parent.parent()

        parent = parent or self.buffer_tree.invisibleRootItem()

        node = {"type": "group", "name": "새 그룹", "children": []}
        item = self._append_buffer_node(parent, node)
        self.buffer_tree.setCurrentItem(item)
        self.buffer_tree.editItem(item, 0)
        self._save_buffer_structure()

    def _add_buffer(self):
        """새 버퍼 추가"""
        parent = self.buffer_tree.currentItem()
        # 버퍼가 선택되어 있으면 그 부모에 추가
        if parent and parent.data(0, ROLE_TYPE) == "buffer":
            parent = parent.parent()

        parent = parent or self.buffer_tree.invisibleRootItem()

        node = {"type": "buffer", "name": "새 버퍼", "data": []}
        item = self._append_buffer_node(parent, node)
        self.buffer_tree.setCurrentItem(item)
        self.buffer_tree.editItem(item, 0)
        # 새 버퍼가 생성되면 클릭 이벤트 강제 호출하여 활성화
        self._on_buffer_tree_item_clicked(item, 0)
        self._save_buffer_structure()

    def _rename_buffer(self):
        item = self.buffer_tree.currentItem()
        if item:
            self.buffer_tree.editItem(item, 0)

    def _delete_buffer(self):
        print("[DBG][BUF][DEL] _delete_buffer: ENTER")
        try:
            item = self.buffer_tree.currentItem()
            print(f"[DBG][BUF][DEL] currentItem={item}")
            if not item:
                print("[DBG][BUF][DEL] no currentItem -> RETURN")
                return

            node_type = item.data(0, ROLE_TYPE)
            payload = item.data(0, ROLE_DATA) or {}
            name = item.text(0) or "(no-name)"
            deleting_id = payload.get("id")
            locked = bool(payload.get("locked"))

            parent = item.parent() or self.buffer_tree.invisibleRootItem()
            idx = parent.indexOfChild(item)
            print(f"[DBG][BUF][DEL] node_type={node_type} name='{name}' id={deleting_id} locked={locked} parent={parent} idx={idx}")

            if locked:
                print("[DBG][BUF][DEL] locked item -> blocked")
                QMessageBox.information(self, "삭제 불가", "이 항목은 고정 항목이라 삭제할 수 없습니다.")
                return

            deleting_active = bool(self.active_buffer_id and deleting_id == self.active_buffer_id)
            print(f"[DBG][BUF][DEL] deleting_active={deleting_active} active_buffer_id={self.active_buffer_id}")

            # ✅ 확인창
            if node_type == "group":
                child_cnt = item.childCount()
                msg = f"그룹 '{name}' 을(를) 삭제할까요?\n\n하위 항목 {child_cnt}개도 함께 삭제됩니다."
            else:
                msg = f"버퍼 '{name}' 을(를) 삭제할까요?"

            reply = QMessageBox.question(
                self,
                "삭제 확인",
                msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            print(f"[DBG][BUF][DEL] confirm reply={reply}")
            if reply != QMessageBox.StandardButton.Yes:
                print("[DBG][BUF][DEL] user cancelled -> RETURN")
                return

            # ✅ 활성 버퍼 삭제면 중앙 트리 저장(유실 방지)
            if deleting_active:
                print("[DBG][BUF][DEL] deleting active buffer -> _save_favorites()")
                try:
                    self._save_favorites()
                except Exception:
                    print("[ERR][BUF][DEL] _save_favorites failed (ignored)")
                    traceback.print_exc()

            # ✅ 실제 트리에서 제거
            taken = parent.takeChild(idx)
            print(f"[DBG][BUF][DEL] takeChild result={taken}")
            del taken

            # ✅ 구조 저장
            print("[DBG][BUF][DEL] _save_buffer_structure()")
            try:
                self._save_buffer_structure()
            except Exception:
                print("[ERR][BUF][DEL] _save_buffer_structure failed")
                traceback.print_exc()

            # ✅ 활성 버퍼였다면 다른 버퍼로 자동 전환
            if deleting_active:
                print("[DBG][BUF][DEL] deleted active -> reset active and auto-select next buffer")
                self.active_buffer_id = None
                self.active_buffer_item = None
                self.active_buffer_node = None
                self._active_buffer_settings_node = None
                self.settings["active_buffer_id"] = None

                found_item = self._first_buffer_item
                print(f"[DBG][BUF][DEL] next buffer found_item={found_item}")
                if found_item:
                    self.buffer_tree.setCurrentItem(found_item)
                    self._on_buffer_tree_item_clicked(found_item, 0)
                else:
                    print("[DBG][BUF][DEL] no buffer remains -> clear fav_tree")
                    try:
                        self.fav_tree.clear()
                    except Exception:
                        traceback.print_exc()

            try:
                self._update_buffer_move_button_state()
            except Exception:
                pass

            print("[DBG][BUF][DEL] DONE")
        except Exception:
            print("[ERR][BUF][DEL] exception in _delete_buffer")
            traceback.print_exc()

    # ----------------- 15-2. 즐겨찾기 버퍼 순서 변경 로직 (수정) -----------------
    def _update_buffer_move_button_state(self):
        """버퍼 트리 이동 버튼 상태 업데이트"""
        item = self.buffer_tree.currentItem()
        if not item:
            self.btn_buffer_move_up.setEnabled(False)
            self.btn_buffer_move_down.setEnabled(False)
            return

        parent = item.parent() or self.buffer_tree.invisibleRootItem()
        index = parent.indexOfChild(item)

        self.btn_buffer_move_up.setEnabled(index > 0)
        self.btn_buffer_move_down.setEnabled(index < parent.childCount() - 1)

    def _move_buffer_up(self):
        item = self.buffer_tree.currentItem()
        if not item: return

        parent = item.parent() or self.buffer_tree.invisibleRootItem()
        index = parent.indexOfChild(item)
        if index > 0:
            taken = parent.takeChild(index)
            parent.insertChild(index - 1, taken)
            self.buffer_tree.setCurrentItem(taken)
            self._save_buffer_structure()
            self._update_buffer_move_button_state()

    def _move_buffer_down(self):
        item = self.buffer_tree.currentItem()
        if not item: return

        parent = item.parent() or self.buffer_tree.invisibleRootItem()
        index = parent.indexOfChild(item)
        if index < parent.childCount() - 1:
            taken = parent.takeChild(index)
            parent.insertChild(index + 1, taken)
            self.buffer_tree.setCurrentItem(taken)
            self._save_buffer_structure()
            self._update_buffer_move_button_state()

    def _on_buffer_context_menu(self, pos):
        """버퍼 트리 컨텍스트 메뉴"""
        item = self.buffer_tree.currentItem()
        menu = QMenu(self)

        act_add_grp = QAction("그룹 추가", self)
        act_add_grp.triggered.connect(self._add_buffer_group)
        menu.addAction(act_add_grp)

        act_add_buf = QAction("버퍼 추가", self)
        act_add_buf.triggered.connect(self._add_buffer)
        menu.addAction(act_add_buf)

        if item:
            menu.addSeparator()
            act_rename = QAction("이름 변경 (F2)", self)
            act_rename.triggered.connect(self._rename_buffer)
            menu.addAction(act_rename)

            act_del = QAction("삭제", self)
            act_del.triggered.connect(self._delete_buffer)
            menu.addAction(act_del)

        menu.exec(self.buffer_tree.viewport().mapToGlobal(pos))


    # ----------------- 16-1. 즐겨찾기 복사/붙여넣기 로직 -----------------
    def _copy_favorite_item(self):
        """선택된 즐겨찾기 항목(다중 가능)을 복사합니다."""
        items = self._selected_fav_items_top()
        if not items:
            return
        payload = [self._serialize_fav_item(it) for it in items]
        # 단일이면 dict로, 다중이면 list로 저장
        self.clipboard_data = payload[0] if len(payload) == 1 else payload
        self.connection_status_label.setText(
            f"{len(items)}개 항목 복사됨."
        )

    def _paste_favorite_item(self):
        """클립보드에 있는 즐겨찾기 항목을 붙여넣습니다."""
        if not self.clipboard_data:
            QMessageBox.warning(
                self, "붙여넣기 오류", "클립보드에 복사된 항목이 없습니다."
            )
            return

        # ✅ 붙여넣기 대상 정규화: notebook/section 선택 상태에서 붙여넣으면
        #    항목(노트북/섹션) 안에 항목이 들어가 트리가 꼬입니다.
        #    따라서 선택이 notebook/section이면 자동으로 group 레벨로 올려 붙여넣습니다.
        parent = self._normalize_fav_paste_parent(self._current_fav_item())

        def _deep_copy_node(node: Dict[str, Any]) -> Dict[str, Any]:
            new_node = node.copy()
            new_node["id"] = str(uuid.uuid4())
            # new_node["name"] = f"복사본 - {new_node['name']}" # 이 줄을 제거하거나 주석 처리
            if "children" in new_node:
                new_node["children"] = [
                    _deep_copy_node(child) for child in new_node["children"]
                ]
            return new_node

        try:
            nodes = self.clipboard_data
            if isinstance(nodes, dict):
                nodes = [nodes]
            if not isinstance(nodes, list):
                nodes = []
            new_items = []

            # ✅ 다중 붙여넣기를 '한 번의 Undo'로 묶기
            with self._fav_bulk_edit(reason=f"paste:{len(nodes)}"):
                for node in nodes:
                    copied_node = _deep_copy_node(node)
                    new_item = self._append_fav_node(parent, copied_node)
                    new_items.append(new_item)
                if new_items:
                    self.fav_tree.setCurrentItem(new_items[-1])

            self.connection_status_label.setText(f"{len(new_items)}개 항목 붙여넣기 완료.")

        except Exception as e:
            QMessageBox.critical(
                self, "붙여넣기 오류", f"항목을 붙여넣는 중 오류가 발생했습니다: {e}"
            )

    def _selected_fav_items_top(self) -> List[QTreeWidgetItem]:
        """
        다중 선택 시 '상위 선택만' 반환합니다.
        (부모와 자식을 동시에 선택했을 때 중복 복사/붙여넣기 방지)
        """
        items = self.fav_tree.selectedItems()
        if not items:
            return []
        # QTreeWidgetItem은 unhashable이므로 id()로 membership set 구성
        selected_ids = {id(it) for it in items}
        top_items: List[QTreeWidgetItem] = []
        for it in items:
            p = it.parent()
            skip = False
            while p is not None:
                if id(p) in selected_ids:
                    skip = True
                    break
                p = p.parent()
            if not skip:
                top_items.append(it)
        return top_items

    def _normalize_fav_paste_parent(self, item: Optional[QTreeWidgetItem]) -> QTreeWidgetItem:
        """
        붙여넣기 대상 정규화:
        - group 선택: 그대로 group 안에 붙여넣기
        - notebook/section 선택: 자동으로 상위 group 레벨에 붙여넣기 (항목-항목 중첩 방지)
        - 그 외/None: 루트
        """
        root = self.fav_tree.invisibleRootItem()
        if not item:
            return root
        try:
            t = item.data(0, ROLE_TYPE)
        except Exception:
            return root
        if t == "group":
            return item
        # notebook/section 등은 group까지 올라간다
        p = item
        while p is not None:
            try:
                if p.data(0, ROLE_TYPE) == "group":
                    return p
            except Exception:
                break
            p = p.parent()
        return root

    def _fav_capture_center_tree_snapshot(self) -> str:
        """
        현재 중앙 FavoritesTree의 상태를 JSON 스냅샷으로 캡쳐한다.
        Undo 그룹(base/final)을 _fav_last_snapshot에 의존하면
        버퍼 전환/리빌드 스킵/hash 최적화/중간 save 타이밍에 의해 base==final로 잡혀
        Ctrl+Z가 '안 먹는' 상태가 생길 수 있어서, 트리에서 직접 캡쳐로 고정한다.
        """
        try:
            data = []
            root = self.fav_tree.invisibleRootItem()
            for i in range(root.childCount()):
                data.append(self._serialize_fav_item(root.child(i)))
            return json.dumps(data, sort_keys=True, ensure_ascii=False)
        except Exception:
            return "[]"

    def _fav_record_snapshot(self, new_snapshot: str):
        """FavoritesTree 변경사항을 Undo/Redo 스택에 기록합니다."""
        # (A) 로드/undo apply 중에는 기록하지 않는다.
        if getattr(self, "_fav_undo_suspended", False):
            self._fav_last_snapshot = new_snapshot
            return

        # (B) 최초 스냅샷
        if self._fav_last_snapshot is None:
            self._fav_last_snapshot = new_snapshot
            return

        # (C) 동일하면 skip
        if new_snapshot == self._fav_last_snapshot:
            return

        # (D) 다중 붙여넣기/다중 삭제 같은 "벌크 변경"에서는
        #     itemChanged가 여러 번 발생하며 _save_favorites()가 연속 호출된다.
        #     이때 매번 undo 스택에 쌓이면 Ctrl+Z가 "한 개씩" 되돌아가서 답답해진다.
        #     => 트랜잭션(depth>0)에서는 _fav_last_snapshot만 갱신하고,
        #        최종 커밋은 _fav_end_undo_group()에서 1회만 수행한다.
        if getattr(self, "_fav_undo_batch_depth", 0) > 0:
            self._fav_last_snapshot = new_snapshot
            self._fav_undo_batch_final_snapshot = new_snapshot
            return

        # (E) 일반 단일 변경
        self._fav_undo_stack.append(self._fav_last_snapshot)
        if len(self._fav_undo_stack) > self._fav_undo_max:
            self._fav_undo_stack = self._fav_undo_stack[-self._fav_undo_max:]
        self._fav_redo_stack.clear()
        self._fav_last_snapshot = new_snapshot

    def _fav_begin_undo_group(self, *, reason: str = "") -> None:
        """여러 변경을 한 번의 Undo step으로 묶기 시작."""
        if self._fav_undo_batch_depth == 0:
            # ✅ base는 _fav_last_snapshot 대신 '현재 트리'에서 직접 캡쳐 (base==final 문제 원천 차단)
            base = self._fav_capture_center_tree_snapshot()
            self._fav_undo_batch_base_snapshot = base
            self._fav_undo_batch_final_snapshot = None
            self._fav_undo_batch_reason = reason or ""
            # base를 last_snapshot에도 맞춰둬야 이후 비교가 흔들리지 않는다.
            self._fav_last_snapshot = base
        self._fav_undo_batch_depth += 1
        try:
            if reason:
                print(f"[DBG][FAV][UNDO_GRP] begin depth={self._fav_undo_batch_depth} reason={reason} base_len={len(self._fav_undo_batch_base_snapshot or '')}")
        except Exception:
            pass

    def _fav_end_undo_group(self) -> None:
        """Undo group 종료: 변경이 있었으면 undo 스택에 1회만 커밋."""
        if self._fav_undo_batch_depth <= 0:
            return
        self._fav_undo_batch_depth -= 1
        if self._fav_undo_batch_depth != 0:
            return

        base = self._fav_undo_batch_base_snapshot or ""
        # ✅ final도 last_snapshot 의존 X: 트리에서 직접 캡쳐
        final = self._fav_undo_batch_final_snapshot
        if final is None:
            final = self._fav_capture_center_tree_snapshot()
        # last_snapshot을 final로 동기화 (Undo/Redo 비교 흔들림 방지)
        self._fav_last_snapshot = final

        changed = (final != base)
        if changed:
            self._fav_undo_stack.append(base)
            if len(self._fav_undo_stack) > self._fav_undo_max:
                self._fav_undo_stack = self._fav_undo_stack[-self._fav_undo_max:]
            self._fav_redo_stack.clear()

        try:
            r = self._fav_undo_batch_reason
            print(
                f"[DBG][FAV][UNDO_GRP] end changed={int(changed)} undo={len(self._fav_undo_stack)} redo={len(self._fav_redo_stack)} reason={r} base_len={len(base)} final_len={len(final)}"
            )
        except Exception:
            pass

        self._fav_undo_batch_base_snapshot = None
        self._fav_undo_batch_final_snapshot = None
        self._fav_undo_batch_reason = ""

    @contextmanager
    def _fav_bulk_edit(self, *, reason: str = ""):
        """
        FavoritesTree를 벌크로 수정할 때 사용.
        - Qt itemChanged 연쇄 save를 막기 위해 fav_tree signals를 잠깐 막고
        - Undo/Redo는 begin/end로 한 번의 step으로 묶는다.
        """
        self._fav_begin_undo_group(reason=reason)
        was_updates_enabled = self.fav_tree.updatesEnabled()
        self.fav_tree.blockSignals(True)
        self.fav_tree.setUpdatesEnabled(False)
        try:
            yield
            # 벌크 작업이 끝나면 딱 1번만 저장(=스냅샷 갱신)
            try:
                self._save_favorites()
            except Exception:
                pass
        finally:
            self.fav_tree.setUpdatesEnabled(was_updates_enabled)
            self.fav_tree.blockSignals(False)
            try:
                if was_updates_enabled:
                    self.fav_tree.viewport().update()
            except Exception:
                pass
            self._fav_end_undo_group()

    def _fav_apply_snapshot(self, snapshot: str):
        """스냅샷(JSON 문자열)을 중앙 즐겨찾기 트리에 적용합니다."""
        try:
            data = json.loads(snapshot) if snapshot else []
            if not isinstance(data, list):
                data = []
        except Exception:
            print("[ERR][FAV][UNDO] invalid snapshot")
            traceback.print_exc()
            return
        self._fav_undo_suspended = True
        try:
            # ✅ Undo/Redo는 무조건 리빌드되어야 한다.
            # 해시 스킵 최적화가 undo 적용을 막는 케이스를 원천 차단.
            self._last_center_payload_hash = None
            self._last_center_payload_snapshot = None
            self._last_center_payload_source_id = 0
            self._fav_last_persisted_hash = None
            self._load_favorites_into_center_tree(data)
            self._save_favorites()
        finally:
            self._fav_undo_suspended = False

    def _undo_favorite_tree(self):
        try:
            print(f"[DBG][FAV][UNDO] called undo={len(self._fav_undo_stack)} redo={len(self._fav_redo_stack)} last_len={len(self._fav_last_snapshot or '')}")
        except Exception:
            pass
        if not self._fav_undo_stack:
            try:
                self.connection_status_label.setText("되돌릴 작업이 없습니다.")
            except Exception:
                pass
            return
        cur = self._fav_last_snapshot
        if cur is None:
            try:
                cur = json.dumps(self.active_buffer_node.get("data", []), sort_keys=True, ensure_ascii=False)
            except Exception:
                cur = ""
        self._fav_redo_stack.append(cur or "")
        snap = self._fav_undo_stack.pop()
        self._fav_apply_snapshot(snap)
        try:
            self.connection_status_label.setText("되돌리기 완료 (Ctrl+Z)")
        except Exception:
            pass

    def _redo_favorite_tree(self):
        try:
            print(f"[DBG][FAV][REDO] called undo={len(self._fav_undo_stack)} redo={len(self._fav_redo_stack)} last_len={len(self._fav_last_snapshot or '')}")
        except Exception:
            pass
        if not self._fav_redo_stack:
            try:
                self.connection_status_label.setText("다시 실행할 작업이 없습니다.")
            except Exception:
                pass
            return
        cur = self._fav_last_snapshot
        if cur is None:
            try:
                cur = json.dumps(self.active_buffer_node.get("data", []), sort_keys=True, ensure_ascii=False)
            except Exception:
                cur = ""
        self._fav_undo_stack.append(cur or "")
        snap = self._fav_redo_stack.pop()
        self._fav_apply_snapshot(snap)
        try:
            self.connection_status_label.setText("다시 실행 완료 (Ctrl+Shift+Z)")
        except Exception:
            pass

    def _cut_favorite_item(self):
        """선택 항목 잘라내기(Ctrl+X): 복사 + 삭제."""
        items = self._selected_fav_items_top()
        if not items:
            cur = self._current_fav_item()
            if cur:
                items = [cur]
        if not items:
            return
        payload_nodes = [self._serialize_fav_item(it) for it in items]
        self.clipboard_data = payload_nodes[0] if len(payload_nodes) == 1 else payload_nodes

        # ✅ 다중 잘라내기를 '한 번의 Undo'로 묶기
        with self._fav_bulk_edit(reason=f"cut:{len(items)}"):
            # 실제 삭제 (부모 기준으로 takeChild)
            for it in items:
                parent = it.parent() or self.fav_tree.invisibleRootItem()
                idx = parent.indexOfChild(it)
                if idx >= 0:
                    parent.takeChild(idx)
        try:
            self.connection_status_label.setText(f"{len(items)}개 항목 잘라내기 완료.")
        except Exception:
            pass

    # ----------------- 16. 즐겨찾기 조작 -----------------
    def _current_fav_item(self) -> Optional[QTreeWidgetItem]:
        items = self.fav_tree.selectedItems()
        return items[0] if items else None

    def _move_item_up(self):
        item = self._current_fav_item()
        if not item:
            return

        parent = item.parent() or self.fav_tree.invisibleRootItem()
        index = parent.indexOfChild(item)

        if index > 0:
            is_expanded = item.isExpanded()
            taken_item = parent.takeChild(index)
            parent.insertChild(index - 1, taken_item)
            taken_item.setExpanded(is_expanded)
            self.fav_tree.setCurrentItem(taken_item)
            self._save_favorites()
            self._update_move_button_state()

    def _move_item_down(self):
        item = self._current_fav_item()
        if not item:
            return

        parent = item.parent() or self.fav_tree.invisibleRootItem()
        index = parent.indexOfChild(item)

        if index < parent.childCount() - 1:
            is_expanded = item.isExpanded()
            taken_item = parent.takeChild(index)
            parent.insertChild(index + 1, taken_item)
            taken_item.setExpanded(is_expanded)
            self.fav_tree.setCurrentItem(taken_item)
            self._save_favorites()
            self._update_move_button_state()

    def _update_move_button_state(self):
        item = self._current_fav_item()

        if not item:
            self.btn_move_up.setEnabled(False)
            self.btn_move_down.setEnabled(False)
            return

        parent = item.parent() or self.fav_tree.invisibleRootItem()
        index = parent.indexOfChild(item)

        self.btn_move_up.setEnabled(index > 0)
        self.btn_move_down.setEnabled(index < parent.childCount() - 1)

    def _add_group(self):
        parent = self._current_fav_item()
        if parent and parent.data(0, ROLE_TYPE) == "section":
            parent = parent.parent()
        parent = parent or self.fav_tree.invisibleRootItem()
        node = {"type": "group", "name": "새 그룹", "children": []}
        item = self._append_fav_node(parent, node)
        self.fav_tree.editItem(item, 0)
        self._save_favorites()

    def _register_all_notebooks_from_current_onenote(self):
        """종합 버퍼를 OneNote의 열린 전자필기장 목록으로 새로고침하고 분류까지 갱신합니다."""
        started_at = time.perf_counter()
        refresh_button = getattr(self, "btn_register_all_notebooks", None)
        if refresh_button is not None:
            refresh_button.setEnabled(False)

        try:
            cur_item = self.buffer_tree.currentItem()
            cur_payload = cur_item.data(0, ROLE_DATA) if cur_item else {}
            is_agg = (
                getattr(self, "active_buffer_id", None) == AGG_BUFFER_ID
                or bool(isinstance(cur_payload, dict) and cur_payload.get("id") == AGG_BUFFER_ID)
            )
            if not is_agg:
                QMessageBox.information(self, "안내", "이 기능은 '종합' 버퍼에서만 사용할 수 있습니다.")
                return

            onenote_window = getattr(self, "onenote_window", None)
            hwnd = None
            try:
                if onenote_window is not None:
                    hwnd = getattr(onenote_window, "handle", None)
                    if callable(hwnd):
                        hwnd = onenote_window.handle()
            except Exception:
                hwnd = None
            if not hwnd:
                self.update_status_and_ui("OneNote 창이 연결되지 않았습니다. 먼저 OneNote 창 연결/선택을 해주세요.", False)
                return

            try:
                self.connection_status_label.setText("종합 새로고침 중...")
                QApplication.processEvents()
            except Exception:
                pass

            try:
                sig = build_window_signature(onenote_window)
            except Exception as e:
                print(f"[WARN][AGG_REFRESH] signature build failed: {e}")
                sig = {}

            notebook_nodes: List[Dict[str, Any]] = []
            seen_keys: Set[str] = set()
            com_error = ""

            def _append_notebook_node(
                nb_name: str,
                *,
                notebook_id: str = "",
                notebook_path: str = "",
            ) -> None:
                nb_name_clean = _strip_stale_favorite_prefix(str(nb_name or "").strip())
                name_key = _normalize_notebook_name_key(nb_name_clean)
                id_key = str(notebook_id or "").strip().casefold()
                dedupe_key = f"id:{id_key}" if id_key else f"name:{name_key}"
                if not nb_name_clean or not name_key or dedupe_key in seen_keys:
                    return
                seen_keys.add(dedupe_key)
                target = {"sig": sig, "notebook_text": nb_name_clean}
                if notebook_id:
                    target["notebook_id"] = str(notebook_id).strip()
                if notebook_path:
                    target["path"] = str(notebook_path).strip()
                notebook_nodes.append(
                    {
                        "type": "notebook",
                        "id": str(uuid.uuid4()),
                        "name": nb_name_clean,
                        "target": target,
                    }
                )

            try:
                for record in _get_open_notebook_records_via_com(refresh=True):
                    _append_notebook_node(
                        record.get("name", ""),
                        notebook_id=record.get("id", ""),
                        notebook_path=record.get("path", ""),
                    )
            except Exception as e:
                com_error = str(e)
                print(f"[WARN][AGG_REFRESH][COM] {e}")

            source = "COM"
            if not notebook_nodes:
                source = "UI"
                try:
                    ensure_pywinauto()
                    if hasattr(self, "_bring_onenote_to_front"):
                        self._bring_onenote_to_front()
                    if not getattr(self, "tree_control", None):
                        self.tree_control = _find_tree_or_list(onenote_window)
                    for nb_name in _collect_root_notebook_names_from_tree(
                        self.tree_control,
                        limit=512,
                    ):
                        _append_notebook_node(nb_name)
                except Exception as e:
                    print(f"[WARN][AGG_REFRESH][UI_FALLBACK] {e}")

            if not notebook_nodes:
                message = "등록할 전자필기장을 찾지 못했습니다."
                if com_error:
                    message += f"\n\nCOM 조회 오류: {com_error}"
                QMessageBox.information(self, "안내", message)
                return

            self._invalidate_aggregate_cache(invalidate_classified_keys=True)
            categorized = self._build_aggregate_categorized_display_nodes(notebook_nodes)

            self._aggregate_reclassify_in_progress = True
            try:
                self._load_favorites_into_center_tree(categorized)
                self._fav_reset_undo_context_from_data(
                    categorized,
                    reason="aggregate_onenote_refresh",
                )
                self._persist_active_aggregate_data(categorized)
            finally:
                self._aggregate_reclassify_in_progress = False

            unclassified_count = len(categorized[0].get("children") or []) if categorized else 0
            classified_count = len(categorized[1].get("children") or []) if len(categorized) > 1 else 0
            elapsed_ms = (time.perf_counter() - started_at) * 1000.0
            self.update_status_and_ui(
                (
                    "종합 새로고침 완료: "
                    f"전체 {len(notebook_nodes)}개, "
                    f"미분류 {unclassified_count}개, "
                    f"분류됨 {classified_count}개 "
                    f"({source}, {elapsed_ms:.0f}ms)"
                ),
                True,
            )
        except Exception as e:
            print(f"[ERROR][AGG_REFRESH] {e}")
            traceback.print_exc()
            QMessageBox.warning(self, "오류", f"종합 새로고침 실패: {e}")
        finally:
            if refresh_button is not None:
                refresh_button.setEnabled(
                    getattr(self, "active_buffer_id", None) == AGG_BUFFER_ID
                )

    def _add_section_from_current(self):
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

        default_name = (info.get("title") or "새 섹션").strip() or "새 섹션"
        name, ok = QInputDialog.getText(
            self, "섹션 즐겨찾기 추가", "표시 이름:", text=default_name
        )
        if not ok or not name.strip():
            return

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

    def _delete_favorite_item(self):
        print("[DBG][FAV][DEL] _delete_favorite_item: ENTER")
        try:
            # ✅ 다중선택 삭제: 상위 선택만 남김(부모/자식 중복 선택 방지)
            targets = self._selected_fav_items_top()
            if not targets:
                item = self._current_fav_item()
                if item:
                    targets = [item]
            print(f"[DBG][FAV][DEL] targets_count={len(targets)}")
            if not targets:
                return

            # ✅ 확인 메시지(한 번만)
            names = [t.text(0) for t in targets[:5]]
            more = "" if len(targets) <= 5 else f" 외 {len(targets)-5}개"
            msg = f"선택한 {len(targets)}개 항목을 삭제할까요?\n- " + "\n- ".join(names) + more
            ret = QMessageBox.question(
                self,
                "삭제 확인",
                msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ret != QMessageBox.StandardButton.Yes:
                return

            # ✅ 안전한 삭제 순서: 깊은 항목부터(자식 먼저)
            def _depth(it: QTreeWidgetItem) -> int:
                d = 0
                p = it.parent()
                while p:
                    d += 1
                    p = p.parent()
                return d
            targets.sort(key=_depth, reverse=True)

            # ✅ 다중 삭제를 '한 번의 Undo'로 묶기
            with self._fav_bulk_edit(reason=f"delete:{len(targets)}"):
                for it in targets:
                    parent = it.parent() or self.fav_tree.invisibleRootItem()
                    idx = parent.indexOfChild(it)
                    print(f"[DBG][FAV][DEL] remove name='{it.text(0)}' depth={_depth(it)} idx={idx}")
                    parent.takeChild(idx)
            print("[DBG][FAV][DEL] DONE multi")
        except Exception:
            print("[ERR][FAV][DEL] exception")
            traceback.print_exc()

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

        # 복사/붙여넣기 메뉴
        menu.addSeparator()

        act_copy = QAction("복사 (Ctrl+C)", self)
        act_copy.triggered.connect(self._copy_favorite_item)
        act_copy.setEnabled(item is not None)
        menu.addAction(act_copy)

        act_paste = QAction("붙여넣기 (Ctrl+V)", self)
        act_paste.triggered.connect(self._paste_favorite_item)
        act_paste.setEnabled(self.clipboard_data is not None)
        menu.addAction(act_paste)

        if item:
            try:
                item_type = item.data(0, ROLE_TYPE)
            except Exception:
                item_type = None
            if item_type in ("section", "notebook"):
                menu.addSeparator()
                act_send_to_codex = QAction("코덱스 작업 위치로 보내기", self)
                act_send_to_codex.triggered.connect(
                    lambda checked=False, fav_item=item: self._sync_codex_target_from_fav_item(
                        fav_item,
                        switch_to_codex=True,
                    )
                )
                menu.addAction(act_send_to_codex)

            menu.addSeparator()
            act_rename = QAction("이름바꾸기", self)
            act_rename.triggered.connect(self._rename_favorite_item)
            menu.addAction(act_rename)

            act_delete = QAction("삭제", self)
            act_delete.triggered.connect(self._delete_favorite_item)
            menu.addAction(act_delete)

        menu.exec(self.fav_tree.viewport().mapToGlobal(pos))

    def _on_fav_item_double_clicked(self, item: QTreeWidgetItem, col: int):
        if not item:
            return
        started_at = time.perf_counter()
        node_type = item.data(0, ROLE_TYPE)
        print(f"[DBG][FAV][DBLCLK] type={node_type} text='{item.text(0)}'")
        # ✅ notebook 타입도 더블클릭 동작해야 함
        if node_type not in ("section", "notebook"):
            return
        self._sync_codex_target_from_fav_item(item)
        self._activate_favorite_section(item, started_at=started_at)

    # _activate_favorite_notebook 제거됨 (기능 통합)

    def _cancel_pending_center_after_activate(self):
        try:
            if self._pending_center_timer is not None:
                self._pending_center_timer.stop()
                self._pending_center_timer.deleteLater()
        except Exception:
            pass
        self._pending_center_timer = None
        worker = self._center_worker
        self._center_worker = None
        if worker is not None:
            try:
                if worker.isRunning():
                    worker.requestInterruption()
                    worker.wait(500)
            except Exception:
                pass

    def _cancel_pending_favorite_activation(self):
        worker = self._favorite_activation_worker
        self._favorite_activation_worker = None
        if worker is not None:
            try:
                if worker.isRunning():
                    worker.requestInterruption()
                    worker.wait(500)
            except Exception:
                pass

    def _retain_qthread_until_finished(self, worker: Optional[QThread], attr_name: str):
        if worker is None:
            return
        self._retained_qthreads.append(worker)

        def _cleanup():
            try:
                if getattr(self, attr_name, None) is worker:
                    setattr(self, attr_name, None)
            except Exception:
                pass
            try:
                self._retained_qthreads.remove(worker)
            except ValueError:
                pass
            try:
                worker.deleteLater()
            except Exception:
                pass

        worker.finished.connect(_cleanup)

    def _mark_favorite_item_stale(
        self,
        item: Optional[QTreeWidgetItem],
        fallback_name: str,
    ) -> str:
        current_name = ""
        if item is not None:
            try:
                current_name = item.text(0) or ""
            except Exception:
                current_name = ""

        base_name = current_name or fallback_name or ""
        stale_prefixes = ("(구) ", "(old) ")
        if not base_name:
            return ""
        if any(base_name.startswith(prefix) for prefix in stale_prefixes):
            return base_name

        new_name = f"(구) {base_name}"
        if item is not None:
            try:
                item.setText(0, new_name)
                self._save_favorites()
            except Exception:
                pass
        return new_name

    def _sync_favorite_notebook_target(
        self,
        item: Optional[QTreeWidgetItem],
        resolved_name: str,
        resolved_notebook_id: str,
    ) -> Dict[str, Any]:
        current_name = ""
        if item is not None:
            try:
                current_name = item.text(0) or ""
            except Exception:
                current_name = ""

        result = {
            "display_name": current_name,
            "changed": False,
            "name_changed": False,
            "was_stale": False,
        }
        if item is None:
            return result

        try:
            payload = item.data(0, ROLE_DATA) or {}
            if not isinstance(payload, dict):
                payload = {}

            target = payload.get("target") or {}
            if not isinstance(target, dict):
                target = {}

            display_name = current_name
            stale_prefixes = ("(구) ", "(old) ")
            for prefix in stale_prefixes:
                if display_name.startswith(prefix):
                    display_name = display_name[len(prefix):]
                    result["was_stale"] = True
                    break

            clean_name = display_name
            resolved_name = (resolved_name or "").strip()
            resolved_notebook_id = (resolved_notebook_id or "").strip()

            if resolved_name and clean_name != resolved_name:
                clean_name = resolved_name
                result["name_changed"] = True

            if current_name != clean_name:
                item.setText(0, clean_name)
                result["changed"] = True

            if resolved_name and target.get("notebook_text") != resolved_name:
                target["notebook_text"] = resolved_name
                result["changed"] = True

            if resolved_notebook_id and target.get("notebook_id") != resolved_notebook_id:
                target["notebook_id"] = resolved_notebook_id
                result["changed"] = True

            payload["target"] = target
            item.setData(0, ROLE_DATA, payload)
            result["display_name"] = clean_name or current_name

            if result["changed"]:
                self._save_favorites()
        except Exception:
            print("[ERR][FAV][SYNC] exception")
            traceback.print_exc()

        return result

    def _handle_favorite_activation_result(
        self,
        item: Optional[QTreeWidgetItem],
        sig: Dict[str, Any],
        display_name: str,
        result: Dict[str, Any],
    ) -> None:
        try:
            connected = self._apply_connected_window_info(result.get("window_info"))
            ok = bool(result.get("ok"))
            target_kind = result.get("target_kind")
            expected_center_text = result.get("expected_center_text")
            resolved_name = str(result.get("resolved_name") or "").strip()
            resolved_notebook_id = str(result.get("resolved_notebook_id") or "").strip()

            if connected and ok:
                notebook_sync = {
                    "display_name": display_name,
                    "changed": False,
                    "name_changed": False,
                    "was_stale": False,
                }
                if target_kind == "notebook":
                    notebook_sync = self._sync_favorite_notebook_target(
                        item, resolved_name, resolved_notebook_id
                    )

                aligned_now = False
                if (
                    target_kind == "notebook"
                    and expected_center_text
                    and getattr(self, "onenote_window", None) is not None
                ):
                    try:
                        aligned_now, _ = scroll_selected_item_to_center(
                            self.onenote_window,
                            self.tree_control,
                        )
                    except Exception:
                        aligned_now = False

                if target_kind in ("section", "notebook") and expected_center_text:
                    if not aligned_now:
                        self._schedule_center_after_activate(
                            sig,
                            expected_center_text,
                            target_kind=target_kind,
                        )
                else:
                    self._cancel_pending_center_after_activate()

                final_name = notebook_sync.get("display_name") or display_name
                if notebook_sync.get("was_stale"):
                    self.update_status_and_ui(
                        f"활성화: '{final_name}' (이름 복원)", True
                    )
                elif notebook_sync.get("name_changed"):
                    self.update_status_and_ui(
                        f"활성화: '{final_name}' (이름 갱신)", True
                    )
                else:
                    self.update_status_and_ui(f"활성화: '{final_name}'", True)
                return

            current_name = ""
            if item is not None:
                try:
                    current_name = item.text(0) or ""
                except Exception:
                    current_name = ""

            stale_prefixes = ("(구) ", "(old) ")
            if item is not None and not any(
                current_name.startswith(prefix) for prefix in stale_prefixes
            ):
                new_name = f"(구) {current_name}"
                item.setText(0, new_name)
                self._save_favorites()
                fail_msg = result.get("error") or f"항목 찾기 실패: '{new_name}'"
                self.update_status_and_ui(fail_msg, True)
            else:
                fail_msg = result.get("error") or f"항목 찾기 실패: '{display_name}'"
                self.update_status_and_ui(fail_msg, True)
        except Exception as e:
            print("[ERR][FAV][ACTIVATE][RESULT] exception")
            traceback.print_exc()
            self.update_status_and_ui(f"즐겨찾기 처리 중 오류: {e}", True)

    def _apply_connected_window_info(self, info: Optional[Dict[str, Any]]) -> bool:
        if not info or not info.get("handle"):
            return False
        try:
            self.onenote_window = Desktop(backend="uia").window(handle=info["handle"])
            if not self.onenote_window.is_visible():
                raise ElementNotFoundError
            save_connection_info(self.onenote_window)
            self._cache_tree_control()
            return True
        except Exception:
            return False

    def _schedule_center_after_activate(
        self,
        sig: Dict[str, Any],
        expected_text: str,
        *,
        target_kind: str = "",
    ):
        self._cancel_pending_center_after_activate()
        if not sig or not expected_text:
            return

        self._center_request_seq += 1
        request_seq = self._center_request_seq
        target_kind = (target_kind or "").strip().lower()
        timer = QTimer(self)
        timer.setSingleShot(True)
        self._pending_center_timer = timer

        def _start_worker():
            if self._pending_center_timer is not timer:
                return

            self._pending_center_timer = None
            worker = CenterAfterActivateWorker(
                sig,
                expected_text,
                target_kind=target_kind,
                parent=self,
            )
            self._center_worker = worker
            self._retain_qthread_until_finished(worker, "_center_worker")

            def _on_done(ok: bool, selected_name: str):
                if self._center_worker is not worker:
                    return

                if ok and selected_name:
                    print(
                        f"[DBG][CENTER][DONE] selected='{selected_name}' req={request_seq}"
                    )
                else:
                    print(f"[DBG][CENTER][SKIP] req={request_seq}")

            worker.done.connect(_on_done)
            worker.start()

        timer.timeout.connect(_start_worker)
        timer.start(40 if target_kind == "notebook" else 120)

    def _activate_favorite_section(
        self,
        item: QTreeWidgetItem,
        *,
        started_at: Optional[float] = None,
    ):
        ensure_pywinauto()
        if not _pwa_ready:
            self.update_status_and_ui(
                "오류: 자동화 모듈이 로드되지 않았습니다.",
                self.center_button.isEnabled(),
            )
            return

        node_type = item.data(0, ROLE_TYPE)
        payload = item.data(0, ROLE_DATA) or {}
        target = payload.get("target") or {}
        display_name = item.text(0)

        sig = target.get("sig") or {}
        if not sig:
            self.update_status_and_ui(
                "오류: 즐겨찾기에 대상 창 정보가 없습니다.",
                self.center_button.isEnabled(),
            )
            self._cancel_pending_favorite_activation()
            return

        self._cancel_pending_favorite_activation()
        self._cancel_pending_center_after_activate()
        if self._try_activate_favorite_fastpath(
            item,
            sig,
            target,
            display_name,
            started_at=started_at,
        ):
            return
        worker = FavoriteActivationWorker(
            sig=sig,
            target=target,
            display_name=display_name,
            auto_center_after_activate=self._auto_center_after_activate,
            parent=self,
        )
        self._favorite_activation_worker = worker
        self._retain_qthread_until_finished(worker, "_favorite_activation_worker")

        def _on_done(result: Dict[str, Any]):
            if self._favorite_activation_worker is not worker:
                return
            return self._handle_favorite_activation_result(
                item=item,
                sig=sig,
                display_name=display_name,
                result=result,
            )

            connected = self._apply_connected_window_info(result.get("window_info"))
            ok = bool(result.get("ok"))
            target_kind = result.get("target_kind")
            expected_center_text = result.get("expected_center_text")

            if connected and ok:
                is_name_restored = False
                current_name = item.text(0)
                restored_name = current_name
                if current_name.startswith("(구) "):
                    restored_name = current_name[4:]
                    item.setText(0, restored_name)
                    self._save_favorites()
                    is_name_restored = True

                if target_kind in ("section", "notebook") and expected_center_text:
                    self._schedule_center_after_activate(sig, expected_center_text)
                else:
                    self._cancel_pending_center_after_activate()

                if is_name_restored:
                    self.update_status_and_ui(f"활성화: '{restored_name}' (이름 복원)", True)
                else:
                    self.update_status_and_ui(f"활성화: '{display_name}'", True)
                return

            current_name = item.text(0)
            if not current_name.startswith("(구) "):
                new_name = f"(구) {current_name}"
                item.setText(0, new_name)
                self._save_favorites()
                fail_msg = result.get("error") or f"섹션 찾기 실패: '{new_name}'"
                self.update_status_and_ui(fail_msg, True)
            else:
                fail_msg = result.get("error") or f"섹션 찾기 실패: '{display_name}' 을 찾을 수 없음"
                self.update_status_and_ui(fail_msg, True)

        worker.done.connect(_on_done)
        worker.start()


# ----------------- 17. 엔트리 포인트 -----------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = OneNoteScrollRemoconApp()
    ex.show()

    try:
        ex.fav_tree.itemDoubleClicked.disconnect()
    except TypeError:
        pass

    def _toggle_group_and_activate_section(item, col):
        node_type = item.data(0, ROLE_TYPE)
        if node_type != "section":
            item.setExpanded(not item.isExpanded())
        else:
            ex._on_fav_item_double_clicked(item, col)

    ex.fav_tree.itemDoubleClicked.connect(_toggle_group_and_activate_section)

    def _toggle_group_and_activate_section_safe(item, col):
        try:
            _toggle_group_and_activate_section(item, col)
        except Exception:
            print("[ERR][FAV][DBLCLK][STANDALONE] exception")
            traceback.print_exc()
            try:
                ex.update_status_and_ui("즐겨찾기 실행 중 오류가 발생했습니다.", True)
            except Exception:
                pass

    try:
        ex.fav_tree.itemDoubleClicked.disconnect(_toggle_group_and_activate_section)
    except Exception:
        pass
    ex.fav_tree.itemDoubleClicked.connect(_toggle_group_and_activate_section_safe)

    sys.exit(app.exec())

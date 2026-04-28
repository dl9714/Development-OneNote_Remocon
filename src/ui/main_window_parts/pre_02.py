# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _mac_record_is_app_only_without_launch_info(record: Dict[str, Any]) -> bool:
    if not IS_MACOS:
        return False
    if str((record or {}).get("url") or "").strip():
        return False
    if _mac_record_has_ui_open_hint(record):
        return False
    hints = _notebook_record_source_hints(record)
    return not hints or bool(hints & _APP_ONLY_NOTEBOOK_SOURCES)


def _make_open_notebook_check_icon(size: int = 18) -> QIcon:
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    try:
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        pen = QPen(QColor("#B9FF5A"))
        pen.setWidth(max(2, size // 6))
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        painter.drawLine(size // 5, size // 2, size // 2 - 1, size - size // 4)
        painter.drawLine(size // 2 - 1, size - size // 4, size - size // 5, size // 4)
    finally:
        painter.end()
    return QIcon(pixmap)


def _onenote_list_hint_text() -> str:
    if IS_MACOS:
        return "더블클릭 또는 Enter로 연결 후 현재 전자필기장 보기 열기"
    return "더블클릭 또는 Enter로 연결 및 중앙 정렬"


def _search_group_title() -> str:
    return "찾기 / 빠른 이동" if IS_MACOS else "검색 / 위치정렬"


def _project_search_label_text() -> str:
    return "전자필기장/섹션 검색" if IS_MACOS else "프로젝트 검색"


def _project_search_placeholder_text() -> str:
    if IS_MACOS:
        return "전자필기장 분류 + 섹션 바로가기 검색 (띄어쓰기 무시)..."
    return "프로젝트/등록영역 + 모듈영역 검색 (띄어쓰기 무시)..."


def _project_search_hint_text() -> str:
    if IS_MACOS:
        return "입력한 글자가 포함된 항목은 전자필기장 분류와 섹션 바로가기에 하이라이트로 표시됩니다."
    return "입력한 글자가 포함된 항목은 프로젝트/등록영역과 모듈영역에 하이라이트로 표시됩니다."


def _project_search_status_text(
    raw_query: str,
    buffer_count: int,
    module_count: int,
) -> str:
    if IS_MACOS:
        return (
            f"항목 검색: '{raw_query}' - 전자필기장 분류 {buffer_count}개, "
            f"섹션 바로가기 {module_count}개 강조"
        )
    return f"프로젝트 검색: '{raw_query}' - 프로젝트 {buffer_count}개, 모듈 {module_count}개 강조"


def _primary_restore_button_text() -> str:
    if IS_MACOS:
        return "현재 전자필기장 보기"
    if IS_WINDOWS:
        return "선택 위치정렬"
    return f"현재 선택된 {_center_target_ui_name()} 중앙으로 정렬"


def _mac_context_summary_text(context: Optional[Dict[str, Any]], fallback: str = "") -> str:
    context = context or {}
    notebook = str(context.get("notebook") or "").strip()
    section = str(context.get("section") or "").strip()
    page = str(context.get("page") or "").strip()
    parts = [value for value in (notebook, section, page) if value]
    if not parts and fallback:
        parts = [str(fallback).strip()]
    return " > ".join(part for part in parts if part)


def _codex_platform_skill_aliases(platform_key: str) -> Dict[str, str]:
    if platform_key == CODEX_PLATFORM_MACOS:
        return {
            "링크 생성": "앱 링크 생성",
            "부모 ID 조회": "상위 위치 조회",
        }
    return {}


def _canonical_codex_platform_skill(platform_key: str, value: str) -> str:
    normalized = str(value or "").strip()
    if not normalized:
        return ""
    return _codex_platform_skill_aliases(platform_key).get(normalized, normalized)


def _codex_active_platform_key() -> str:
    return CODEX_PLATFORM_MACOS if IS_MACOS else CODEX_PLATFORM_WINDOWS


def _codex_platform_variants() -> List[Tuple[str, str]]:
    return [
        (CODEX_PLATFORM_WINDOWS, "Windows"),
        (CODEX_PLATFORM_MACOS, "macOS"),
    ]


def _codex_platform_display_name(platform_key: str) -> str:
    for key, label in _codex_platform_variants():
        if key == platform_key:
            return label
    return str(platform_key or "").strip() or "플랫폼"


def _codex_platform_engine_summary(platform_key: str) -> str:
    if platform_key == CODEX_PLATFORM_MACOS:
        return "OneNote for Mac 접근성/UI 자동화"
    return "Windows OneNote COM API"


def _codex_platform_structure_summary(platform_key: str) -> str:
    if platform_key == CODEX_PLATFORM_MACOS:
        return "왼쪽 패널의 전자필기장/섹션/페이지 구조와 현재 보이는 UI 상태를 기준으로 작업"
    return "COM ID와 GetHierarchy/GetPageContent 결과를 기준으로 작업"


# ----------------- 0.0 설정 파일 경로 헬퍼 -----------------
def _get_app_base_path() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))


def _get_default_settings_file_path() -> str:
    if IS_MACOS:
        return os.path.join(_settings_path_config_dir(), SETTINGS_FILE)
    return os.path.join(_get_app_base_path(), SETTINGS_FILE)


def _settings_path_config_dir() -> str:
    if IS_WINDOWS:
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(base, "OneNote_Remocon")
    if IS_MACOS:
        return os.path.join(
            os.path.expanduser("~/Library/Application Support"),
            "OneNote_Remocon",
        )
    return os.path.join(os.path.expanduser("~/.config"), "OneNote_Remocon")


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
_OPEN_ALL_DEBUG_LOG_PATH = os.path.expanduser(
    "~/Library/Logs/OneNote_Remocon/open_all_debug.log"
)


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

_publish_context(globals())

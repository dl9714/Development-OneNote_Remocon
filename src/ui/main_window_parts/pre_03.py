# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _append_open_all_debug_log(line: str) -> None:
    try:
        os.makedirs(os.path.dirname(_OPEN_ALL_DEBUG_LOG_PATH), exist_ok=True)
        with open(_OPEN_ALL_DEBUG_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(
                f"{time.strftime('%Y-%m-%d %H:%M:%S')} {str(line or '').rstrip()}\n"
            )
    except Exception:
        pass


def _dump_json_text(obj: Dict[str, Any]) -> str:
    return json.dumps(obj, ensure_ascii=False, indent=2)

def _write_json_text(path: str, text: str) -> bool:
    """내용이 바뀐 경우에만 .bak 백업 후 원자적으로 저장합니다."""
    parent_dir = os.path.dirname(path)
    if parent_dir:
        os.makedirs(parent_dir, exist_ok=True)
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


def _sanitize_connection_signature_for_platform(sig: Any) -> Optional[Dict[str, Any]]:
    if not isinstance(sig, dict):
        return None

    bundle_id = str(sig.get("bundle_id") or "").strip()
    class_name = str(sig.get("class_name") or "").strip()
    exe_name = str(sig.get("exe_name") or "").strip().lower()
    exe_path = str(sig.get("exe_path") or "").strip()

    if IS_MACOS:
        if ONENOTE_MAC_BUNDLE_ID in (bundle_id, class_name, exe_name):
            return sig
        if exe_path or exe_name.endswith(".exe") or "\\" in exe_path:
            return None
        return None

    if IS_WINDOWS:
        if bundle_id == ONENOTE_MAC_BUNDLE_ID or class_name == ONENOTE_MAC_BUNDLE_ID:
            return None
        return sig

    return sig


def _sanitize_settings_for_platform_inplace(settings: Dict[str, Any]) -> bool:
    migrated = False
    current_sig = settings.get("connection_signature")
    sanitized_sig = _sanitize_connection_signature_for_platform(current_sig)
    if sanitized_sig != current_sig:
        settings["connection_signature"] = sanitized_sig
        migrated = True
    return migrated


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


def load_settings(cache_object: bool = True) -> Dict[str, Any]:
    # 설정 파일 경로를 실행 파일 위치 기준으로 가져옴
    settings_path = _get_settings_file_path()

    file_sig = _get_file_signature(settings_path)
    cache_entry = _SETTINGS_OBJECT_CACHE.get(settings_path)
    if (
        cache_object
        and file_sig is not None
        and cache_entry
        and cache_entry.get("sig") == file_sig
    ):
        cached = copy.deepcopy(cache_entry.get("data") or DEFAULT_SETTINGS)
        _sanitize_settings_for_platform_inplace(cached)
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
                _sanitize_settings_for_platform_inplace(settings)
                _ensure_default_and_aggregate_inplace(settings)
                try:
                    _write_json(settings_path, settings)
                except Exception as e:
                    print(f"[WARN] 초기 설정 복사 실패(메모리 로드는 계속): {e}")
                if cache_object:
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
        if cache_object:
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
        migrated = _sanitize_settings_for_platform_inplace(settings) or migrated
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
                    migrated = _sanitize_settings_for_platform_inplace(seed_settings) or migrated
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

        if cache_object:
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

_publish_context(globals())

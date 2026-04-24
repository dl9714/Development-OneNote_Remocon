# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin01:
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
        self._open_notebooks_refresh_worker: Optional[OpenNotebookRecordsWorker] = None
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
        self._open_all_candidate_count: Optional[int] = None
        self._open_all_candidate_scope: str = ""
        self._open_all_candidate_stats: Dict[str, int] = {}
        self._open_all_candidate_count_dirty = True
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
        self._mac_empty_scan_retry_attempts = 0
        self._mac_empty_scan_retry_timer = QTimer(self)
        self._mac_empty_scan_retry_timer.setSingleShot(True)
        self._mac_empty_scan_retry_timer.timeout.connect(
            self._retry_onenote_list_after_empty_macos_scan
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

            # 저장된 연결이 있으면 자동 재연결을 먼저 끝내고, 목록은 그 결과에 맞춰 갱신한다.
            # macOS에서 초기 스캔/재연결을 동시에 돌리면 잠깐 앱명만 보이거나 상태가 흔들릴 수 있다.
            has_saved_sig = bool(self.settings.get("connection_signature"))
            if has_saved_sig:
                reconnect_delay_ms = 400 if IS_MACOS else 0
                QTimer.singleShot(reconnect_delay_ms, self._start_auto_reconnect)
            else:
                QTimer.singleShot(0, self.refresh_onenote_list)
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

        if self.settings.get("connection_signature"):
            self.connection_status_label.setText("준비됨 (자동 재연결 중...)")
        else:
            self.connection_status_label.setText("준비됨")

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

_publish_context(globals())

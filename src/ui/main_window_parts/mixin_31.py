# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin31:

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
            if IS_MACOS:
                refreshed = self._coerce_macos_window(w)
                if refreshed is not None:
                    self.onenote_window = refreshed
                    w = refreshed
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
        expected_text: str = "",
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

        if IS_MACOS:
            refreshed = self._coerce_macos_window(getattr(self, "onenote_window", None))
            if refreshed is not None:
                self.onenote_window = refreshed

        if preselected_tree_control is not None:
            self.tree_control = preselected_tree_control
        elif not self.tree_control:
            self.tree_control = _find_tree_or_list(self.onenote_window)

        success, item_name = scroll_selected_item_to_center(
            self.onenote_window,
            self.tree_control,
            selected_item=preselected_item,
            expected_text=expected_text,
        )

        if success:
            print(
                f"[DBG][CENTER][DONE] source={debug_source} success=True "
                f"item={item_name!r} elapsed_ms={(time.perf_counter() - op_started_at) * 1000.0:.1f} "
                f"at_s={(time.perf_counter() - self._t_boot):.3f}"
            )
            if IS_MACOS:
                summary = _mac_context_summary_text(
                    self._mac_selected_outline_context(self.onenote_window),
                    fallback=str(item_name or ""),
                )
                success_message = (
                    f"성공: 현재 전자필기장 보기 열기 완료."
                    + (f" {summary}" if summary else "")
                )
            else:
                success_message = f"성공: '{item_name}' 중앙 정렬 완료."
            self.update_status_and_ui(success_message, True)
        elif allow_retry:
            self.tree_control = _find_tree_or_list(self.onenote_window)
            success, item_name = scroll_selected_item_to_center(
                self.onenote_window,
                self.tree_control,
                expected_text=expected_text,
            )
            if success:
                print(
                    f"[DBG][CENTER][DONE] source={debug_source} success=True retry=1 "
                    f"item={item_name!r} elapsed_ms={(time.perf_counter() - op_started_at) * 1000.0:.1f} "
                    f"at_s={(time.perf_counter() - self._t_boot):.3f}"
                )
                if IS_MACOS:
                    summary = _mac_context_summary_text(
                        self._mac_selected_outline_context(self.onenote_window),
                        fallback=str(item_name or ""),
                    )
                    success_message = (
                        f"성공: 현재 전자필기장 보기 열기 완료."
                        + (f" {summary}" if summary else "")
                    )
                else:
                    success_message = f"성공: '{item_name}' 중앙 정렬 완료."
                self.update_status_and_ui(success_message, True)
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
                (
                    "실패: 현재 전자필기장 보기 열기를 완료하지 못했습니다."
                    if IS_MACOS
                    else "실패: 선택 항목을 찾거나 정렬하지 못했습니다."
                ),
                True,
            )

_publish_context(globals())

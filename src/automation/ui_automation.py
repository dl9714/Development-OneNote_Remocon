# -*- coding: utf-8 -*-
"""
UI Automation 모듈

pywinauto를 사용한 UI 자동화 기능을 제공합니다.
"""

from typing import Optional, List


class UIAutomationClient:
    """UI Automation 작업을 수행하는 클래스 (싱글톤 패턴)"""

    _instance = None
    _initialized = False

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        """UI Automation 클라이언트를 초기화합니다."""
        if not UIAutomationClient._initialized:
            self._pwa_ready = False
            self.Desktop = None
            self.WindowNotFoundError = None
            self.ElementNotFoundError = None
            self.TimeoutError = None
            self.UIAWrapper = None
            self.UIAElementInfo = None
            self.mouse = None
            self.keyboard = None
            UIAutomationClient._initialized = True

    def ensure_loaded(self) -> bool:
        """
        pywinauto 라이브러리를 지연 로딩합니다.

        Returns:
            bool: 로딩 성공 여부
        """
        if self._pwa_ready:
            return True

        try:
            from pywinauto import Desktop, mouse, keyboard
            from pywinauto.findwindows import (
                WindowNotFoundError,
                ElementNotFoundError,
            )
            from pywinauto.timings import TimeoutError
            from pywinauto.controls.uiawrapper import UIAWrapper
            from pywinauto.uia_element_info import UIAElementInfo

            self.Desktop = Desktop
            self.WindowNotFoundError = WindowNotFoundError
            self.ElementNotFoundError = ElementNotFoundError
            self.TimeoutError = TimeoutError
            self.UIAWrapper = UIAWrapper
            self.UIAElementInfo = UIAElementInfo
            self.mouse = mouse
            self.keyboard = keyboard
            self._pwa_ready = True
            return True

        except ImportError as e:
            print(f"[ERROR] pywinauto 임포트 실패: {e}")
            return False

    def is_ready(self) -> bool:
        """
        pywinauto가 로드되었는지 확인합니다.

        Returns:
            bool: 로드 여부
        """
        return self._pwa_ready

    # ==================== 텍스트 정규화 ====================

    @staticmethod
    def normalize_text(text: Optional[str]) -> str:
        """
        텍스트를 정규화합니다 (공백 정리 및 소문자 변환).

        Args:
            text: 정규화할 텍스트

        Returns:
            str: 정규화된 텍스트
        """
        return " ".join(((text or "").strip().split())).lower()

    # ==================== Tree/List 컨트롤 찾기 ====================

    def find_tree_or_list(self, window) -> Optional[object]:
        """
        윈도우에서 Tree 또는 List 컨트롤을 찾습니다.

        Args:
            window: pywinauto 윈도우 객체

        Returns:
            Optional[object]: Tree 또는 List 컨트롤 객체 또는 None
        """
        if not self.ensure_loaded():
            return None

        for control_type in ("Tree", "List"):
            try:
                control = window.child_window(
                    control_type=control_type, found_index=0
                ).wrapper_object()
                return control
            except Exception:
                continue

        return None

    # ==================== 선택된 항목 가져오기 ====================

    def get_selected_tree_item(self, tree_control) -> Optional[object]:
        """
        트리 컨트롤에서 선택된 항목을 빠르게 가져옵니다.
        여러 fallback 전략을 시도합니다.

        Args:
            tree_control: 트리 컨트롤 객체

        Returns:
            Optional[object]: 선택된 항목 또는 None
        """
        if not self.ensure_loaded():
            return None

        # 전략 1: get_selection() 메소드 사용
        try:
            if hasattr(tree_control, "get_selection"):
                sel = tree_control.get_selection()
                if sel:
                    return sel[0]
        except Exception:
            pass

        # 전략 2: iface_selection 인터페이스 사용
        try:
            iface_sel = getattr(tree_control, "iface_selection", None)
            if iface_sel:
                arr = iface_sel.GetSelection()
                length = getattr(arr, "Length", 0)
                if length and length > 0:
                    el = arr.GetElement(0)
                    return self.UIAWrapper(self.UIAElementInfo(el))
        except Exception:
            pass

        # 전략 3: children()을 순회하며 is_selected() 확인
        try:
            for item in tree_control.children():
                try:
                    if item.is_selected():
                        return item
                except Exception:
                    pass
        except Exception:
            pass

        # 전략 4: descendants(TreeItem)을 순회하며 is_selected() 확인
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

    # ==================== 섹션 검색 및 선택 ====================

    def select_section_by_text(
        self, window, text: str, tree_control: Optional[object] = None
    ) -> bool:
        """
        텍스트로 섹션을 검색하고 선택합니다.

        Args:
            window: OneNote 윈도우 객체
            text: 검색할 섹션 텍스트
            tree_control: 트리 컨트롤 (None이면 자동 탐색)

        Returns:
            bool: 선택 성공 여부
        """
        if not self.ensure_loaded():
            return False

        try:
            # 트리 컨트롤 찾기
            if tree_control is None:
                tree_control = self.find_tree_or_list(window)
            if not tree_control:
                return False

            # 정규화된 타겟 텍스트
            target_norm = self.normalize_text(text)

            def _scan(types: List[str]) -> bool:
                """주어진 컨트롤 타입들을 스캔합니다."""
                for control_type in types:
                    try:
                        for item in tree_control.descendants(control_type=control_type):
                            try:
                                if (
                                    self.normalize_text(item.window_text())
                                    == target_norm
                                ):
                                    # 선택 시도
                                    try:
                                        item.select()
                                        return True
                                    except Exception:
                                        # select() 실패 시 click_input() 시도
                                        try:
                                            item.click_input()
                                            return True
                                        except Exception:
                                            pass
                            except Exception:
                                pass
                    except Exception:
                        pass
                return False

            # TreeItem, ListItem 순서로 검색
            if _scan(["TreeItem"]):
                return True
            if _scan(["ListItem"]):
                return True

            return False

        except Exception as e:
            print(f"[ERROR] 섹션 선택 실패: {e}")
            return False

    def select_notebook_by_text(
        self, window, text: str, tree_control: Optional[object] = None
    ) -> bool:
        """
        텍스트로 전자필기장(노트북)을 검색하고 선택합니다.
        NOTE: OneNote UI 구조에 따라 TreeItem/ListItem로 보일 수 있어 둘 다 스캔합니다.
        """
        if not self.ensure_loaded():
            return False
        try:
            if tree_control is None:
                tree_control = self.find_tree_or_list(window)
            if not tree_control:
                return False

            target_norm = self.normalize_text(text)

            def _scan(types: List[str]) -> bool:
                for control_type in types:
                    try:
                        for item in tree_control.descendants(control_type=control_type):
                            try:
                                if (
                                    self.normalize_text(item.window_text())
                                    == target_norm
                                ):
                                    try:
                                        item.select()
                                        return True
                                    except Exception:
                                        try:
                                            item.click_input()
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
            print(f"[ERROR] 전자필기장 선택 실패: {e}")
            return False

    # ==================== 키보드 및 마우스 제어 ====================

    def send_keys(self, keys: str) -> bool:
        """
        키보드 입력을 전송합니다.

        Args:
            keys: 전송할 키 문자열 (예: "{UP}", "{DOWN}")

        Returns:
            bool: 성공 여부
        """
        if not self.ensure_loaded():
            return False

        try:
            self.keyboard.send_keys(keys)
            return True
        except Exception as e:
            print(f"[ERROR] 키 전송 실패: {e}")
            return False

    def wheel_scroll(self, coords, wheel_dist: int) -> bool:
        """
        마우스 휠 스크롤을 수행합니다.

        Args:
            coords: 좌표 튜플 (x, y)
            wheel_dist: 휠 거리 (양수: 위로, 음수: 아래로)

        Returns:
            bool: 성공 여부
        """
        if not self.ensure_loaded():
            return False

        try:
            # scroll() 메소드 시도
            try:
                self.mouse.scroll(coords=coords, wheel_dist=wheel_dist)
                return True
            except Exception:
                pass

            # wheel() 메소드 시도
            try:
                self.mouse.wheel(coords=coords, wheel_dist=wheel_dist)
                return True
            except Exception:
                pass

            return False

        except Exception as e:
            print(f"[ERROR] 마우스 휠 스크롤 실패: {e}")
            return False

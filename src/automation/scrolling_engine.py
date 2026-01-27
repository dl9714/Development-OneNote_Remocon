# -*- coding: utf-8 -*-
"""
스크롤 엔진 모듈

OneNote 섹션 트리의 스크롤 기능을 제공합니다.
"""

import time
from typing import Optional, Tuple
from src.constants import (
    SCROLL_WAIT_TIMEOUT,
    SCROLL_POLL_INTERVAL,
    SCROLL_CENTER_TOLERANCE,
    SCROLL_MAX_REPEATS,
)
from src.automation.ui_automation import UIAutomationClient


class ScrollingEngine:
    """스크롤 작업을 수행하는 클래스"""

    def __init__(self, ui_automation: Optional[UIAutomationClient] = None):
        """
        스크롤 엔진을 초기화합니다.

        Args:
            ui_automation: UI Automation 클라이언트 (None이면 새로 생성)
        """
        self.ui_automation = ui_automation or UIAutomationClient()

    # ==================== 위치 안정화 대기 ====================

    @staticmethod
    def wait_rect_settle(
        get_rect_func, timeout: float = SCROLL_WAIT_TIMEOUT, interval: float = SCROLL_POLL_INTERVAL
    ) -> None:
        """
        요소의 위치가 안정화될 때까지 대기합니다.

        Args:
            get_rect_func: 요소의 Rectangle를 반환하는 함수
            timeout: 최대 대기 시간 (초)
            interval: 폴링 간격 (초)
        """
        start = time.perf_counter()
        prev = get_rect_func()

        while time.perf_counter() - start < timeout:
            time.sleep(interval)
            cur = get_rect_func()
            # top과 bottom이 2픽셀 이내로 안정되면 종료
            if abs(cur.top - prev.top) < 2 and abs(cur.bottom - prev.bottom) < 2:
                break
            prev = cur

    # ==================== 패턴 기반 스크롤 ====================

    def scroll_via_pattern(
        self, container, direction: str, small: bool = True, repeats: int = 1
    ) -> bool:
        """
        UIAutomation 스크롤 패턴을 사용하여 스크롤합니다.

        Args:
            container: 스크롤할 컨테이너
            direction: 스크롤 방향 ("up" 또는 "down")
            small: 작은 단위로 스크롤할지 여부
            repeats: 반복 횟수

        Returns:
            bool: 성공 여부
        """
        if not self.ui_automation.ensure_loaded():
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

            # 스크롤 양 결정
            v_inc = ScrollAmount_SmallIncrement if small else ScrollAmount_LargeIncrement
            v_dec = ScrollAmount_SmallDecrement if small else ScrollAmount_LargeDecrement
            v_amount = v_inc if direction == "down" else v_dec

            # 스크롤 실행
            for _ in range(max(1, repeats)):
                iface.Scroll(ScrollAmount_NoAmount, v_amount)

            return True

        except Exception:
            return False

    # ==================== 휠 기반 스크롤 (Fallback) ====================

    def scroll_via_wheel(self, container, steps: int) -> bool:
        """
        마우스 휠을 사용하여 스크롤합니다.
        여러 fallback 전략을 시도합니다.

        Args:
            container: 스크롤할 컨테이너
            steps: 스크롤 단계 (양수: 위로, 음수: 아래로)

        Returns:
            bool: 성공 여부
        """
        if steps == 0:
            return True

        if not self.ui_automation.ensure_loaded():
            return False

        # 전략 1: wheel_scroll() 메소드
        try:
            if hasattr(container, "wheel_scroll"):
                container.wheel_scroll(steps)
                return True
        except Exception:
            pass

        # 전략 2: wheel_mouse_input() 메소드
        try:
            if hasattr(container, "wheel_mouse_input"):
                container.wheel_mouse_input(wheel_dist=steps)
                return True
        except Exception:
            pass

        # 전략 3: 마우스 좌표로 직접 휠 스크롤
        try:
            rect = container.rectangle()
            center = rect.mid_point()
            if self.ui_automation.wheel_scroll((center.x, center.y), steps):
                return True
        except Exception:
            pass

        # 전략 4: 키보드 입력 사용
        try:
            container.set_focus()
            if steps > 0:
                self.ui_automation.send_keys("{UP %d}" % steps)
            else:
                self.ui_automation.send_keys("{DOWN %d}" % abs(steps))
            return True
        except Exception:
            pass

        return False

    # ==================== 요소를 중앙으로 스크롤 ====================

    def center_element_in_view(self, element, container) -> bool:
        """
        요소를 컨테이너의 중앙에 위치하도록 스크롤합니다.

        Args:
            element: 중앙에 배치할 요소
            container: 스크롤 컨테이너

        Returns:
            bool: 성공 여부
        """
        if not self.ui_automation.ensure_loaded():
            return False

        try:
            # ScrollIntoView 호출
            try:
                element.iface_scroll_item.ScrollIntoView()
            except AttributeError:
                return False

            # 위치 안정화 대기
            self.wait_rect_settle(lambda: element.rectangle())

            # 요소와 컨테이너의 중심 좌표 계산
            rect_container = container.rectangle()
            rect_item = element.rectangle()
            item_center_y = (rect_item.top + rect_item.bottom) / 2
            container_center_y = (rect_container.top + rect_container.bottom) / 2
            offset = item_center_y - container_center_y

            # 이미 중앙에 있으면 종료
            if abs(offset) <= SCROLL_CENTER_TOLERANCE:
                return True

            # 스크롤 반복 횟수 계산 함수
            def calculate_steps(dy):
                return max(1, min(SCROLL_MAX_REPEATS, int(abs(dy) / 150)))

            # 최대 3회 반복하여 중앙에 맞춤
            for _ in range(3):
                if abs(offset) <= SCROLL_CENTER_TOLERANCE:
                    break

                direction = "down" if offset > 0 else "up"
                repeats = calculate_steps(offset)

                # 패턴 기반 스크롤 시도
                used_pattern = self.scroll_via_pattern(
                    container, direction=direction, small=True, repeats=repeats
                )

                # 패턴 실패 시 휠 스크롤 사용
                if not used_pattern:
                    wheel_steps = -repeats if offset > 0 else repeats
                    self.scroll_via_wheel(container, wheel_steps)

                # 짧은 대기 후 위치 재계산
                time.sleep(SCROLL_POLL_INTERVAL)

                rect_container = container.rectangle()
                rect_item = element.rectangle()
                item_center_y = (rect_item.top + rect_item.bottom) / 2
                container_center_y = (rect_container.top + rect_container.bottom) / 2
                offset = item_center_y - container_center_y

            return True

        except Exception as e:
            print(f"[WARN] 중앙 정렬 중 오류: {e}")
            return False

    # ==================== 선택된 항목을 중앙으로 스크롤 ====================

    def scroll_selected_to_center(
        self, window, tree_control: Optional[object] = None
    ) -> Tuple[bool, Optional[str]]:
        """
        트리 컨트롤에서 선택된 항목을 중앙으로 스크롤합니다.

        Args:
            window: OneNote 윈도우 객체
            tree_control: 트리 컨트롤 (None이면 자동 탐색)

        Returns:
            Tuple[bool, Optional[str]]: (성공 여부, 항목 이름)
        """
        if not self.ui_automation.ensure_loaded():
            return False, None

        try:
            # 트리 컨트롤 찾기
            if tree_control is None:
                tree_control = self.ui_automation.find_tree_or_list(window)
            if not tree_control:
                return False, None

            # 선택된 항목 가져오기
            selected_item = self.ui_automation.get_selected_tree_item(tree_control)
            if not selected_item:
                return False, None

            # 항목 이름 가져오기
            item_name = selected_item.window_text()

            # 중앙으로 스크롤
            success = self.center_element_in_view(selected_item, tree_control)

            return success, item_name if success else None

        except Exception:
            return False, None

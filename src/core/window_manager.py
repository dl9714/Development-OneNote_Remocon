# -*- coding: utf-8 -*-
"""
윈도우 관리 모듈

OneNote 윈도우 검색, 검증, 연결 시그니처 관리를 담당합니다.
"""

import os
import ctypes
from ctypes import wintypes
from typing import Optional, List, Dict, Any
from src.constants import (
    ONENOTE_CLASS_NAME,
    ONENOTE_KEYWORDS,
    ONENOTE_EXE_NAMES,
    SCORE_WEIGHT_PID,
    SCORE_WEIGHT_TITLE,
    SCORE_WEIGHT_CLASS,
    SCORE_WEIGHT_EXE,
)


class WindowManager:
    """OneNote 윈도우를 검색하고 관리하는 클래스"""

    def __init__(self):
        """윈도우 관리자를 초기화합니다."""
        self._user32 = ctypes.windll.user32
        self._current_pid = os.getpid()

    # ==================== Win32 API 헬퍼 메소드 ====================

    def _get_window_text(self, hwnd: int) -> str:
        """
        윈도우 핸들로부터 윈도우 제목을 가져옵니다.

        Args:
            hwnd: 윈도우 핸들

        Returns:
            str: 윈도우 제목
        """
        length = self._user32.GetWindowTextLengthW(hwnd)
        buf = ctypes.create_unicode_buffer(length + 1 if length > 0 else 1)
        self._user32.GetWindowTextW(hwnd, buf, len(buf))
        return buf.value

    def _get_class_name(self, hwnd: int) -> str:
        """
        윈도우 핸들로부터 클래스 이름을 가져옵니다.

        Args:
            hwnd: 윈도우 핸들

        Returns:
            str: 클래스 이름
        """
        buf = ctypes.create_unicode_buffer(256)
        self._user32.GetClassNameW(hwnd, buf, 256)
        return buf.value

    def get_process_image_path(self, pid: int) -> Optional[str]:
        """
        프로세스 ID로부터 실행 파일의 전체 경로를 가져옵니다.

        Args:
            pid: 프로세스 ID

        Returns:
            Optional[str]: 실행 파일 경로 또는 None
        """
        if not pid:
            return None

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
            return None

        try:
            # 1차 버퍼
            size = 512
            while True:
                buf_len = wintypes.DWORD(size)
                buf = ctypes.create_unicode_buffer(buf_len.value)
                ok = QueryFullProcessImageNameW(hProcess, 0, buf, ctypes.byref(buf_len))
                if ok:
                    return buf.value
                # 버퍼 부족 시 한 번 정도 키워 봄
                err = ctypes.get_last_error()
                # ERROR_INSUFFICIENT_BUFFER = 122
                if err == 122 and size < 4096:
                    size *= 2
                    continue
                return None
        finally:
            CloseHandle(hProcess)

    # ==================== 윈도우 열거 ====================

    def enumerate_windows(self, filter_title_substr=None) -> List[Dict[str, Any]]:
        """
        모든 보이는 윈도우를 빠르게 열거합니다.

        Args:
            filter_title_substr: 제목 필터 (문자열 또는 문자열 리스트)

        Returns:
            List[Dict[str, Any]]: 윈도우 정보 딕셔너리 리스트
                - handle: 윈도우 핸들
                - title: 윈도우 제목
                - class_name: 윈도우 클래스 이름
                - pid: 프로세스 ID
        """
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
                if not self._user32.IsWindowVisible(hwnd):
                    return True
                title = self._get_window_text(hwnd)
                if not title:
                    return True
                if filters and not any(f in title.lower() for f in filters):
                    return True

                cls = self._get_class_name(hwnd)
                pid = wintypes.DWORD()
                self._user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
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

        self._user32.EnumWindows(_enum_proc, 0)
        return results

    # ==================== OneNote 윈도우 검증 ====================

    def is_onenote_window(self, window_info: Dict[str, Any]) -> bool:
        """
        주어진 창 정보가 실제로 OneNote 앱 창인지 엄격하게 확인합니다.

        Args:
            window_info: 윈도우 정보 딕셔너리 (handle, title, class_name, pid)

        Returns:
            bool: OneNote 윈도우 여부
        """
        # 자기 자신의 프로세스는 제외
        if window_info.get("pid") == self._current_pid:
            return False

        title_lower = window_info.get("title", "").lower()
        cls = window_info.get("class_name", "")
        pid = window_info.get("pid")

        # 1. Classic Desktop (OMain*) - 레거시 OneNote
        if "omain" in (cls or "").lower():
            return True

        # 2. Modern App (ApplicationFrameWindow) + 타이틀 키워드
        if cls == ONENOTE_CLASS_NAME:
            if any(keyword in title_lower for keyword in ONENOTE_KEYWORDS):
                return True

        # 3. Fallback: 제목에 키워드 + EXE 확인
        if any(keyword in title_lower for keyword in ONENOTE_KEYWORDS):
            exe_path = self.get_process_image_path(pid)
            if exe_path:
                exe_name = os.path.basename(exe_path).lower()
                if any(onenote_exe in exe_name for onenote_exe in ONENOTE_EXE_NAMES):
                    return True

        return False

    def enumerate_onenote_windows(self) -> List[Dict[str, Any]]:
        """
        시스템에서 모든 OneNote 윈도우를 찾습니다.

        Returns:
            List[Dict[str, Any]]: OneNote 윈도우 정보 리스트
        """
        all_windows = self.enumerate_windows(filter_title_substr=ONENOTE_KEYWORDS)
        onenote_windows = [w for w in all_windows if self.is_onenote_window(w)]
        return onenote_windows

    # ==================== 윈도우 시그니처 관리 ====================

    def create_window_signature(self, window_element) -> Dict[str, Any]:
        """
        pywinauto 윈도우 요소로부터 시그니처를 생성합니다.

        Args:
            window_element: pywinauto 윈도우 객체

        Returns:
            Dict[str, Any]: 윈도우 시그니처
                - handle: 윈도우 핸들
                - pid: 프로세스 ID
                - class_name: 클래스 이름
                - title: 윈도우 제목
                - exe_path: 실행 파일 경로
                - exe_name: 실행 파일 이름
        """
        try:
            pid = window_element.process_id()
        except Exception:
            pid = None

        exe_path = self.get_process_image_path(pid) if pid else None
        exe_name = os.path.basename(exe_path).lower() if exe_path else None

        try:
            handle = window_element.handle
        except Exception:
            handle = None

        try:
            title = window_element.window_text()
        except Exception:
            title = None

        try:
            cls_name = window_element.class_name()
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

    def score_candidate(self, candidate: Dict[str, Any], signature: Dict[str, Any]) -> int:
        """
        후보 윈도우가 시그니처와 얼마나 일치하는지 점수를 계산합니다.

        Args:
            candidate: 후보 윈도우 정보
            signature: 찾으려는 윈도우 시그니처

        Returns:
            int: 일치 점수 (높을수록 더 일치)
        """
        try:
            title = (candidate.get("title") or "").lower()
            cls = candidate.get("class_name") or ""
            pid = candidate.get("pid")
            exe_path = self.get_process_image_path(pid) or ""
            exe_name = os.path.basename(exe_path).lower() if exe_path else ""

            score = 0

            # 핸들이 정확히 일치하면 최고점
            if signature.get("handle") and candidate.get("handle") == signature["handle"]:
                score += 100

            # 실행 파일 이름 일치
            if signature.get("exe_name") and exe_name == signature["exe_name"]:
                score += 50

            # OneNote EXE인지 확인
            if "onenote.exe" in exe_name:
                score += 50

            # 제목에 OneNote 키워드 포함
            if any(keyword in title for keyword in ONENOTE_KEYWORDS):
                score += 25

            # 클래스 이름 일치
            if signature.get("class_name") and cls == signature["class_name"]:
                score += 10

            # PID 일치
            if signature.get("pid") and pid == signature["pid"]:
                score += 8

            # 이전 제목과 유사성
            prev_title = (signature.get("title") or "").lower()
            if prev_title:
                if prev_title in title:
                    score += 6
                else:
                    for keyword in ONENOTE_KEYWORDS:
                        if keyword in prev_title and keyword in title:
                            score += 4
                            break

            # OneNote 앱 프레임 윈도우
            if cls == ONENOTE_CLASS_NAME:
                score += 5

            return score

        except Exception:
            return -1

    def find_window_by_signature(
        self, signature: Dict[str, Any], min_score: int = 30
    ) -> Optional[object]:
        """
        시그니처를 사용하여 윈도우를 찾습니다.

        Args:
            signature: 윈도우 시그니처
            min_score: 최소 일치 점수 (기본값: 30)

        Returns:
            Optional[object]: pywinauto 윈도우 객체 또는 None
        """
        # pywinauto를 이 시점에서 임포트 (지연 로딩)
        try:
            from pywinauto import Desktop
        except ImportError:
            print("[ERROR] pywinauto를 임포트할 수 없습니다.")
            return None

        # 핸들로 먼저 시도
        h = signature.get("handle")
        if h:
            try:
                w = Desktop(backend="uia").window(handle=h)
                if w.is_visible():
                    return w
            except Exception:
                pass

        # 모든 윈도우를 열거하고 점수 계산
        candidates = self.enumerate_windows(filter_title_substr=None)
        best, best_score = None, -1

        for c in candidates:
            s = self.score_candidate(c, signature)
            if s > best_score:
                best, best_score = c, s

        # 최소 점수 이상이면 윈도우 반환
        if best and best_score >= min_score:
            try:
                w = Desktop(backend="uia").window(handle=best["handle"])
                if w.is_visible():
                    return w
            except Exception:
                return None

        return None

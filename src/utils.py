# -*- coding: utf-8 -*-
"""
유틸리티 함수 모듈

애플리케이션 전반에서 사용되는 유틸리티 함수들을 제공합니다.
"""

import sys
import os


def app_base_path() -> str:
    """실행 환경에 맞는 애플리케이션 기준 경로."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def resource_path(relative_path: str) -> str:
    """
    PyInstaller에서 묶인 리소스 파일을 찾는 경로를 반환합니다.

    Args:
        relative_path: 리소스의 상대 경로

    Returns:
        str: 리소스의 전체 경로

    Notes:
        - PyInstaller로 패키징된 경우: sys._MEIPASS 디렉토리 사용
        - 스크립트 실행인 경우: 현재 작업 디렉토리 사용
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = app_base_path()

    return os.path.join(base_path, relative_path)

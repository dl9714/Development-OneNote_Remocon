# -*- coding: utf-8 -*-
"""
플랫폼 공통 헬퍼 모음.

Windows 전용 구현을 유지하면서 macOS용 경로/아이콘/파일 열기 동작을 분리한다.
"""

from __future__ import annotations

import os
import sys
from typing import Optional


IS_WINDOWS = sys.platform.startswith("win")
IS_MACOS = sys.platform == "darwin"

ONENOTE_MAC_BUNDLE_ID = "com.microsoft.onenote.mac"
ONENOTE_MAC_APP_NAMES = ("Microsoft OneNote", "OneNote")
MAC_OPEN_PATH = "/usr/bin/open"
MAC_OSASCRIPT_PATH = "/usr/bin/osascript"


def default_icon_path() -> str:
    """현재 플랫폼에 맞는 기본 앱 아이콘 경로를 반환한다."""
    return "assets/app_icon.ico" if IS_WINDOWS else "assets/app_icon.png"


def settings_config_dir(app_name: str = "OneNote_Remocon") -> str:
    """플랫폼별 사용자 설정 포인터 디렉토리."""
    if IS_WINDOWS:
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(base, app_name)
    if IS_MACOS:
        return os.path.join(
            os.path.expanduser("~/Library/Application Support"),
            app_name,
        )
    return os.path.join(os.path.expanduser("~/.config"), app_name)


def open_path_in_system(path: str) -> None:
    """플랫폼 기본 앱으로 파일/폴더를 연다."""
    if not path:
        return
    if IS_WINDOWS:
        os.startfile(path)  # type: ignore[attr-defined]
        return
    import subprocess

    if IS_MACOS:
        subprocess.Popen([MAC_OPEN_PATH, path])
        return
    subprocess.Popen(["xdg-open", path])


def open_url_in_system(url: str) -> None:
    """플랫폼 기본 브라우저/앱으로 URL을 연다."""
    open_path_in_system(url)


def is_macos_accessibility_trusted() -> bool:
    """현재 프로세스가 macOS 접근성 제어 권한을 가졌는지 확인한다."""
    if not IS_MACOS:
        return True
    try:
        import ctypes
        from ctypes.util import find_library

        lib_path = find_library("ApplicationServices") or "ApplicationServices"
        app_services = ctypes.CDLL(lib_path)
        checker = app_services.AXIsProcessTrusted
        checker.restype = ctypes.c_bool
        return bool(checker())
    except Exception:
        return True


def open_macos_accessibility_settings() -> None:
    """macOS 접근성 권한 설정 화면을 연다."""
    if not IS_MACOS:
        return
    open_url_in_system(
        "x-apple.systempreferences:com.apple.preference.security?Privacy_Accessibility"
    )


def platform_label() -> str:
    if IS_WINDOWS:
        return "windows"
    if IS_MACOS:
        return "macos"
    return sys.platform


def onenote_app_identifier() -> Optional[str]:
    if IS_WINDOWS:
        return None
    if IS_MACOS:
        return ONENOTE_MAC_BUNDLE_ID
    return None

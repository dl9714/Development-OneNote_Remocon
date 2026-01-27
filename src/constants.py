# -*- coding: utf-8 -*-
"""
상수 정의 모듈

애플리케이션 전반에서 사용되는 상수들을 정의합니다.
"""

from PyQt6.QtCore import Qt

# ----------------- 파일 경로 관련 -----------------
SETTINGS_FILE = "OneNote_Remocon_Setting.json"
APP_ICON_PATH = "app_icon.ico"

# ----------------- 윈도우 관련 -----------------
ONENOTE_CLASS_NAME = "ApplicationFrameWindow"

# ----------------- 스크롤 관련 -----------------
SCROLL_STEP_SENSITIVITY = 40
SCROLL_WAIT_TIMEOUT = 0.3
SCROLL_POLL_INTERVAL = 0.03
SCROLL_MAX_REPEATS = 5
SCROLL_CENTER_TOLERANCE = 10  # 픽셀 단위

# ----------------- 즐겨찾기 관련 -----------------
DEFAULT_FAVORITES_BUFFER = "기본 즐겨찾기 버퍼"

# ----------------- Qt ItemDataRole 관련 -----------------
ROLE_TYPE = Qt.ItemDataRole.UserRole + 1
ROLE_DATA = Qt.ItemDataRole.UserRole + 2

# ----------------- 기본 설정 -----------------
DEFAULT_SETTINGS = {
    "window_geometry": {"x": 200, "y": 180, "width": 960, "height": 540},
    "splitter_states": None,
    "connection_signature": None,
    "favorites_buffers": {DEFAULT_FAVORITES_BUFFER: []},
    "active_buffer": DEFAULT_FAVORITES_BUFFER,
}

# ----------------- OneNote 키워드 -----------------
ONENOTE_KEYWORDS = ["onenote", "원노트"]
ONENOTE_EXE_NAMES = ["onenote.exe", "onenoteim.exe"]

# ----------------- 윈도우 시그니처 점수 가중치 -----------------
SCORE_WEIGHT_PID = 100
SCORE_WEIGHT_TITLE = 50
SCORE_WEIGHT_CLASS = 25
SCORE_WEIGHT_EXE = 10

# ----------------- UI 관련 -----------------
DEFAULT_WINDOW_WIDTH = 960
DEFAULT_WINDOW_HEIGHT = 540
DEFAULT_WINDOW_X = 200
DEFAULT_WINDOW_Y = 180

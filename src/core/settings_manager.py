# -*- coding: utf-8 -*-
"""
설정 관리 모듈

애플리케이션 설정 파일의 로드, 저장, 마이그레이션을 담당합니다.
"""

import sys
import json
import os
from typing import Dict, Any, Optional
from src.constants import SETTINGS_FILE, DEFAULT_SETTINGS, DEFAULT_FAVORITES_BUFFER


class SettingsManager:
    """애플리케이션 설정을 관리하는 클래스"""

    def __init__(self):
        """설정 관리자를 초기화합니다."""
        self._settings: Dict[str, Any] = {}
        self._settings_path: Optional[str] = None

    def get_settings_path(self) -> str:
        """
        설정 파일(쓰기 가능)의 경로를 반환합니다.

        Returns:
            str: 설정 파일의 전체 경로

        Notes:
            - PyInstaller로 패키징된 경우: 실행 파일(.exe)이 위치한 디렉토리
            - 스크립트 실행인 경우: 현재 작업 디렉토리
        """
        if self._settings_path:
            return self._settings_path

        # sys.frozen은 PyInstaller에 의해 생성된 실행 파일인지 확인하는 일반적인 방법입니다.
        if getattr(sys, "frozen", False):
            # 실행 파일(.exe)이 있는 디렉토리
            base_path = os.path.dirname(sys.executable)
        else:
            # 스크립트 실행 환경 (현재 작업 디렉토리)
            base_path = os.path.abspath(".")

        self._settings_path = os.path.join(base_path, SETTINGS_FILE)
        return self._settings_path

    def load(self) -> Dict[str, Any]:
        """
        설정 파일을 로드합니다.

        Returns:
            Dict[str, Any]: 로드된 설정 딕셔너리

        Notes:
            - 파일이 없으면 기본 설정 반환
            - 구버전 설정 자동 마이그레이션
            - 오류 발생 시 기본 설정 반환
        """
        settings_path = self.get_settings_path()

        if not os.path.exists(settings_path):
            self._settings = DEFAULT_SETTINGS.copy()
            return self._settings

        try:
            with open(settings_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            # 하위 호환성을 위한 마이그레이션 로직
            data = self._migrate_settings(data)

            # 기본 설정으로 시작하고 로드된 데이터로 업데이트
            self._settings = DEFAULT_SETTINGS.copy()
            self._settings.update(data)

            return self._settings

        except Exception as e:
            print(f"[ERROR] 설정 파일 로드 실패: {e}")
            self._settings = DEFAULT_SETTINGS.copy()
            return self._settings

    def save(self, data: Optional[Dict[str, Any]] = None) -> bool:
        """
        설정을 파일에 저장합니다.

        Args:
            data: 저장할 설정 딕셔너리 (None이면 내부 설정 사용)

        Returns:
            bool: 저장 성공 여부
        """
        if data is None:
            data = self._settings
        else:
            self._settings = data

        settings_path = self.get_settings_path()

        try:
            # 구버전 favorites 키 제거 (마이그레이션 완료)
            save_data = data.copy()
            if "favorites" in save_data:
                del save_data["favorites"]

            with open(settings_path, "w", encoding="utf-8") as f:
                json.dump(save_data, f, ensure_ascii=False, indent=2)

            return True

        except Exception as e:
            print(f"[ERROR] 설정 파일 저장 실패: {e}")
            return False

    def get(self, key: str, default: Any = None) -> Any:
        """
        설정 값을 가져옵니다.

        Args:
            key: 설정 키
            default: 키가 없을 때 반환할 기본값

        Returns:
            Any: 설정 값
        """
        return self._settings.get(key, default)

    def set(self, key: str, value: Any) -> None:
        """
        설정 값을 설정합니다.

        Args:
            key: 설정 키
            value: 설정 값
        """
        self._settings[key] = value

    def update(self, data: Dict[str, Any]) -> None:
        """
        여러 설정 값을 한 번에 업데이트합니다.

        Args:
            data: 업데이트할 설정 딕셔너리
        """
        self._settings.update(data)

    def get_all(self) -> Dict[str, Any]:
        """
        모든 설정을 반환합니다.

        Returns:
            Dict[str, Any]: 전체 설정 딕셔너리
        """
        return self._settings.copy()

    def _migrate_settings(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        구버전 설정을 새 버전으로 마이그레이션합니다.

        Args:
            data: 로드된 설정 데이터

        Returns:
            Dict[str, Any]: 마이그레이션된 설정 데이터
        """
        # 구버전 favorites → favorites_buffers 마이그레이션
        if "favorites" in data and "favorites_buffers" not in data:
            print("[INFO] 구 버전 설정 감지. 새 즐겨찾기 버퍼 구조로 마이그레이션합니다.")
            data["favorites_buffers"] = {DEFAULT_FAVORITES_BUFFER: data["favorites"]}
            data["active_buffer"] = DEFAULT_FAVORITES_BUFFER
            del data["favorites"]

        return data

    # 편의 메소드들
    def get_window_geometry(self) -> Dict[str, int]:
        """윈도우 지오메트리 설정을 가져옵니다."""
        return self.get("window_geometry", DEFAULT_SETTINGS["window_geometry"])

    def set_window_geometry(self, x: int, y: int, width: int, height: int) -> None:
        """윈도우 지오메트리 설정을 저장합니다."""
        self.set("window_geometry", {"x": x, "y": y, "width": width, "height": height})

    def get_splitter_states(self) -> Optional[Any]:
        """스플리터 상태를 가져옵니다."""
        return self.get("splitter_states")

    def set_splitter_states(self, states: Any) -> None:
        """스플리터 상태를 저장합니다."""
        self.set("splitter_states", states)

    def get_connection_signature(self) -> Optional[Dict[str, Any]]:
        """윈도우 연결 시그니처를 가져옵니다."""
        return self.get("connection_signature")

    def set_connection_signature(self, signature: Optional[Dict[str, Any]]) -> None:
        """윈도우 연결 시그니처를 저장합니다."""
        self.set("connection_signature", signature)

    def get_favorites_buffers(self) -> Dict[str, list]:
        """즐겨찾기 버퍼들을 가져옵니다."""
        return self.get("favorites_buffers", {DEFAULT_FAVORITES_BUFFER: []})

    def set_favorites_buffers(self, buffers: Dict[str, list]) -> None:
        """즐겨찾기 버퍼들을 저장합니다."""
        self.set("favorites_buffers", buffers)

    def get_active_buffer(self) -> str:
        """활성 즐겨찾기 버퍼 이름을 가져옵니다."""
        return self.get("active_buffer", DEFAULT_FAVORITES_BUFFER)

    def set_active_buffer(self, buffer_name: str) -> None:
        """활성 즐겨찾기 버퍼를 설정합니다."""
        self.set("active_buffer", buffer_name)

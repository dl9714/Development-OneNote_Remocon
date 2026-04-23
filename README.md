# OneNote 전자필기장 스크롤 리모컨

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![GUI](https://img.shields.io/badge/GUI-PyQt6-green.svg)
![OS](https://img.shields.io/badge/OS-Windows%20%7C%20macOS-blue)

OneNote 전자필기장 스크롤 리모컨은 OneNote 창에 연결해서 현재 선택된 항목을 다시 찾고, 즐겨찾기에서 빠르게 이동하고, 왼쪽 목록의 선택 항목을 화면 안으로 정렬해 주는 데 초점을 둔 데스크톱 유틸리티입니다.

이 저장소는 이제 Windows와 macOS를 함께 지원합니다.

- Windows: 기존 `pywinauto` + PowerShell COM 흐름을 유지합니다.
- macOS: `System Events` 접근성 자동화 기반으로 창 연결, 항목 선택, 위치 정렬, 경로 조회 흐름을 분리해서 동작합니다.

## 주요 기능

- OneNote 창 자동/수동 연결
- 현재 선택된 항목 중앙 정렬
- 텍스트 기반 섹션 검색 및 선택
- 즐겨찾기 등록, 그룹화, 재정렬, 내보내기/가져오기
- Codex용 OneNote 작업 위치 조회
- 플랫폼별 OneNote 작업 템플릿 생성

## 플랫폼별 백엔드

### Windows

- 창 탐색: Win32 + `pywinauto`
- UI 제어: `pywinauto`
- OneNote 작업 템플릿: PowerShell COM

### macOS

- 창 탐색: `System Events`
- UI 제어: 접근성(`osascript`) 기반
- OneNote 작업 템플릿: macOS UI 자동화/경로 기반 안내

## 설치

### Windows

```bash
pip install PyQt6 pywinauto
```

### macOS

```bash
pip install PyQt6
```

추가로 macOS에서는 앱을 제어하려면 터미널/파이썬 실행 환경에 접근성 권한이 필요합니다.

## 실행

```bash
python main.py
```

## macOS에서 달라진 점

- Windows 전용 `pywinauto`, Win32, COM 호출은 그대로 남겨두고 macOS 경로를 별도로 추가했습니다.
- 아이콘, 설정 포인터 경로, 파일 열기 동작, 빌드 스펙이 플랫폼별로 분기됩니다.
- OneNote 위치 조회는 macOS에서는 현재 열린 창의 전자필기장/섹션 경로를 읽는 방식으로 동작합니다.
- Codex 템플릿은 macOS에서는 COM ID 대신 경로 문자열과 UI 단계 중심으로 생성됩니다.

## 알려진 제약

- macOS OneNote는 Windows COM API가 없어서 ID 기반 조작보다는 경로/접근성 기반으로 동작합니다.
- macOS에서 긴 목록의 중앙 정렬은 접근성 스크롤바 값에 의존하므로 OneNote UI 구조가 크게 바뀌면 보정이 필요할 수 있습니다.
- `실제 OneNote 전체 열기`처럼 Windows COM에 크게 기대던 기능은 macOS에서 best-effort 방식으로 동작합니다.

## 빌드

PyInstaller 스펙 파일은 이제 상대 경로와 플랫폼별 아이콘을 사용합니다.

```bash
pyinstaller OneNote_Remocon.spec
```

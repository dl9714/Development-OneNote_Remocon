@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: ============================================
:: OneNote_Remocon Build Automation Script
:: Mode: onedir (faster startup, folder output)
:: ============================================

echo.
echo ╔══════════════════════════════════════════════════╗
echo ║   OneNote Remocon Build Automation (onedir)      ║
echo ╚══════════════════════════════════════════════════╝
echo.

:: 현재 스크립트 위치로 이동
cd /d "%~dp0"
echo [INFO] Working directory: %CD%
echo.

:: Python 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python이 설치되어 있지 않거나 PATH에 없습니다.
    pause
    exit /b 1
)

:: PyInstaller 확인
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [WARN] PyInstaller가 설치되어 있지 않습니다. 설치를 시작합니다...
    pip install pyinstaller
    if errorlevel 1 (
        echo [ERROR] PyInstaller 설치에 실패했습니다.
        pause
        exit /b 1
    )
)

:: 이전 빌드 결과물 삭제
echo [INFO] Cleaning previous build artifacts...
if exist "build" rmdir /s /q "build"
if exist "dist\OneNote_Remocon" rmdir /s /q "dist\OneNote_Remocon"

:: 빌드 실행
echo.
echo [INFO] Building OneNote_Remocon (onedir mode)...
echo ─────────────────────────────────────────────────
echo.

pyinstaller --onedir ^
    --windowed ^
    --name "OneNote_Remocon" ^
    --icon="assets/app_icon.ico" ^
    --add-data "assets/app_icon.ico;assets" ^
    --add-data "assets/app_icon.png;assets" ^
    main.py

if errorlevel 1 (
    echo.
    echo [ERROR] 빌드에 실패했습니다.
    pause
    exit /b 1
)

echo.
echo ─────────────────────────────────────────────────
echo [SUCCESS] 빌드 완료!
echo.
echo 출력 경로: %CD%\dist\OneNote_Remocon\
echo 실행 파일: %CD%\dist\OneNote_Remocon\OneNote_Remocon.exe
echo.

:: 빌드 결과 폴더 열기 (선택)
set /p OPEN_FOLDER="결과 폴더를 열까요? (Y/N): "
if /i "!OPEN_FOLDER!"=="Y" (
    explorer "dist\OneNote_Remocon"
)

echo.
echo [DONE] 빌드 자동화 스크립트 종료
pause

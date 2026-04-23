# -*- mode: python ; coding: utf-8 -*-

import os
import subprocess
import sys


SPEC_PATH = os.path.abspath(
    globals().get("__file__", os.path.join(os.getcwd(), "OneNote_Remocon.spec"))
)
ROOT = os.path.dirname(SPEC_PATH)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from src.app_version import APP_BUILD_VERSION, APP_VERSION

WINDOWS_ICON_PATH = os.path.join(ROOT, "assets", "app_icon.ico")
MAC_PNG_ICON_PATH = os.path.join(ROOT, "assets", "app_icon.png")
MAC_BUNDLE_ICON_PATH = os.path.join(ROOT, "assets", "app_icon.icns")
MAC_ENTITLEMENTS_PATH = os.path.join(ROOT, "assets", "macos-entitlements.plist")
MAC_DEFAULT_CODESIGN_IDENTITY = "OneNote Remocon Local Code Signing"


def _mac_codesign_identity():
    if sys.platform != "darwin":
        return None
    identity = os.environ.get("ONENOTE_REMOCON_CODESIGN_IDENTITY", MAC_DEFAULT_CODESIGN_IDENTITY)
    if not identity:
        return None
    try:
        result = subprocess.run(
            ["/usr/bin/security", "find-certificate", "-c", identity],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
    except OSError:
        return None
    return identity if result.returncode == 0 else None

datas = [
    (WINDOWS_ICON_PATH, "assets"),
    (MAC_PNG_ICON_PATH, "assets"),
]

if os.path.exists(MAC_BUNDLE_ICON_PATH):
    datas.append((MAC_BUNDLE_ICON_PATH, "assets"))

settings_path = os.path.join(ROOT, "OneNote_Remocon_Setting.json")
if os.path.exists(settings_path):
    datas.append((settings_path, "."))

exe_icon = None
if sys.platform.startswith("win") and os.path.exists(WINDOWS_ICON_PATH):
    exe_icon = [WINDOWS_ICON_PATH]
elif sys.platform == "darwin" and os.path.exists(MAC_PNG_ICON_PATH):
    exe_icon = [MAC_PNG_ICON_PATH]

codesign_identity = _mac_codesign_identity()
entitlements_file = (
    MAC_ENTITLEMENTS_PATH
    if sys.platform == "darwin" and codesign_identity and os.path.exists(MAC_ENTITLEMENTS_PATH)
    else None
)

a = Analysis(
    [os.path.join(ROOT, "main.py")],
    pathex=[ROOT],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="OneNote_Remocon",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=codesign_identity,
    entitlements_file=entitlements_file,
    icon=exe_icon,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="OneNote_Remocon",
)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="OneNote_Remocon.app",
        icon=MAC_BUNDLE_ICON_PATH if os.path.exists(MAC_BUNDLE_ICON_PATH) else None,
        bundle_identifier="com.codex.onenote-remocon",
        info_plist={
            "CFBundleDisplayName": "OneNote Remocon",
            "CFBundleShortVersionString": APP_VERSION,
            "CFBundleVersion": APP_BUILD_VERSION,
            "NSAppleEventsUsageDescription": (
                "OneNote 창을 찾고 이동하며 정렬 기능을 수행하기 위해 Apple Events 권한이 필요합니다."
            ),
        },
    )

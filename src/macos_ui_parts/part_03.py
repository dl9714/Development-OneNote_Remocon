# -*- coding: utf-8 -*-
from __future__ import annotations

from src.macos_ui_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _enumerate_macos_windows_via_system_events(filter_title_substr=None) -> List[Dict[str, Any]]:
    """보이는 macOS 앱 윈도우 목록."""
    if not IS_MACOS:
        return []

    script = r'''
tell application "System Events"
    set resultItems to {}
    repeat with p in (every application process whose background only is false)
        try
            set procName to name of p as text
            set procPid to unix id of p
            set procFront to frontmost of p as text
            try
                set procBundle to bundle identifier of p as text
            on error
                set procBundle to ""
            end try
            repeat with w in windows of p
                try
                    set winName to name of w as text
                on error
                    set winName to ""
                end try
                set winName to my cleanText(winName)
                set procName to my cleanText(procName)
                set procBundle to my cleanText(procBundle)
                set end of resultItems to procName & tab & procPid & tab & procFront & tab & procBundle & tab & winName
            end repeat
        end try
    end repeat
    set AppleScript's text item delimiters to linefeed
    return resultItems as text
end tell

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    raw = _run_osascript(script, timeout=20)
    if not raw:
        return []

    if isinstance(filter_title_substr, str):
        filters = [_normalize_text(filter_title_substr)]
    elif filter_title_substr:
        filters = [_normalize_text(str(item)) for item in filter_title_substr]
    else:
        filters = []

    results: List[Dict[str, Any]] = []
    for line in raw.splitlines():
        parts = line.split("\t")
        if len(parts) < 5:
            continue
        app_name, pid_text, front_text, bundle_id, title = parts[:5]
        try:
            pid = int(pid_text)
        except Exception:
            continue
        if filters and not any(token in _normalize_text(title) for token in filters):
            continue
        handle = _synthetic_window_handle(pid, title, bundle_id)
        results.append(
            {
                "handle": handle,
                "title": title,
                "class_name": bundle_id or app_name,
                "pid": pid,
                "bundle_id": bundle_id,
                "app_name": app_name,
                "frontmost": front_text == "true",
            }
        )
    return results


def _read_onenote_current_notebook_name(pid: int, title_hint: str = "") -> str:
    if not pid:
        return ""
    locator = _applescript_window_locator(pid, title_hint or "")
    script = locator + r'''
        set currentNotebook to ""
        repeat with e in (entire contents of targetWindow)
            try
                if role of e is "AXButton" then
                    set candidateText to ""
                    try
                        set candidateText to description of e as text
                    end try
                    if candidateText does not contain "(현재 전자필기장)" and candidateText does not contain "(현재 전자 필기장)" then
                        try
                            set candidateText to value of e as text
                        end try
                    end if
                    if candidateText does not contain "(현재 전자필기장)" and candidateText does not contain "(현재 전자 필기장)" then
                        try
                            set candidateText to title of e as text
                        end try
                    end if
                    set candidateText to my cleanText(candidateText)
                    if candidateText contains "(현재 전자필기장)" or candidateText contains "(현재 전자 필기장)" then
                        set currentNotebook to candidateText
                        exit repeat
                    end if
                end if
            end try
        end repeat
        return my cleanText(currentNotebook)

on cleanText(v)
    if v is missing value then return ""
    set t to v as text
    set t to my replaceText(tab, " ", t)
    set t to my replaceText(return, " ", t)
    set t to my replaceText(linefeed, " ", t)
    return t
end cleanText

on replaceText(findText, replaceText, sourceText)
    set AppleScript's text item delimiters to findText
    set parts to every text item of sourceText
    set AppleScript's text item delimiters to replaceText
    set newText to parts as text
    set AppleScript's text item delimiters to ""
    return newText
end replaceText
'''
    try:
        raw = _run_osascript(script, timeout=10)
    except Exception:
        return ""
    return _extract_current_notebook_name(raw)


def _hydrate_missing_macos_window_titles(
    results: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    if not results:
        return results

    missing = [
        info for info in results if not _clean_field(str(info.get("title") or ""))
    ]
    if not missing:
        return results

    fallback_results = _enumerate_macos_windows_via_system_events(
        filter_title_substr=None
    )
    fallback_titles: Dict[Tuple[int, str], List[str]] = {}
    for info in fallback_results:
        title = _clean_field(str(info.get("title") or ""))
        if not title:
            continue
        pid = int(info.get("pid") or 0)
        if not pid:
            continue
        bundle_id = str(info.get("bundle_id") or "")
        app_name = str(info.get("app_name") or "")
        for key in (
            (pid, bundle_id),
            (pid, app_name),
            (pid, ""),
        ):
            fallback_titles.setdefault(key, []).append(title)

    notebook_name_cache: Dict[int, str] = {}
    for info in results:
        if _clean_field(str(info.get("title") or "")):
            continue

        pid = int(info.get("pid") or 0)
        bundle_id = str(info.get("bundle_id") or "")
        app_name = str(info.get("app_name") or "")

        matched_titles = (
            fallback_titles.get((pid, bundle_id))
            or fallback_titles.get((pid, app_name))
            or fallback_titles.get((pid, ""))
            or []
        )
        if matched_titles:
            info["title"] = matched_titles[0]
            continue

        if bundle_id == ONENOTE_MAC_BUNDLE_ID or app_name in ONENOTE_MAC_APP_NAMES:
            if pid not in notebook_name_cache:
                notebook_name_cache[pid] = _read_onenote_current_notebook_name(pid)
            notebook_name = notebook_name_cache.get(pid) or ""
            if notebook_name:
                info["title"] = notebook_name

    return results


def enumerate_macos_windows(filter_title_substr=None) -> List[Dict[str, Any]]:
    results = _enumerate_macos_windows_via_coregraphics(
        filter_title_substr=filter_title_substr
    )
    if results:
        return _hydrate_missing_macos_window_titles(results)
    return _hydrate_missing_macos_window_titles(
        _enumerate_macos_windows_via_system_events(
        filter_title_substr=filter_title_substr
        )
    )


def enumerate_macos_windows_quick(filter_title_substr=None) -> List[Dict[str, Any]]:
    """Finder 실행 초기처럼 민감한 구간에서 쓰는 가벼운 창 열거."""
    return _enumerate_macos_windows_via_coregraphics(
        filter_title_substr=filter_title_substr
    )


class MacDesktop:
    """pywinauto Desktop 대체."""

    def __init__(self, backend: str = "uia"):
        self.backend = backend

    def window(self, handle=None):
        if handle is None:
            raise MacAutomationError("window handle이 필요합니다.")
        for info in enumerate_macos_windows():
            if int(info.get("handle") or 0) == int(handle):
                return MacWindow(info)
        raise MacAutomationError("해당 macOS 윈도우를 찾지 못했습니다.")


@dataclass
class MacWindow:
    info: Dict[str, Any]

    @property
    def handle(self) -> int:
        return int(self.info.get("handle") or 0)

    def window_text(self) -> str:
        title = _clean_field(str(self.info.get("title") or ""))
        if title:
            return title

        handle = self.handle
        pid = self.process_id()
        bundle_id = self.bundle_id()
        for info in enumerate_macos_windows():
            same_handle = handle and int(info.get("handle") or 0) == handle
            same_pid = pid and int(info.get("pid") or 0) == pid
            same_bundle = bundle_id and str(info.get("bundle_id") or "") == bundle_id
            if same_handle or (same_pid and same_bundle):
                self.info.update(info)
                title = _clean_field(str(info.get("title") or ""))
                if title:
                    return title

        if bundle_id == ONENOTE_MAC_BUNDLE_ID or self.app_name() in ONENOTE_MAC_APP_NAMES:
            title = _read_onenote_current_notebook_name(pid)
            if title:
                self.info["title"] = title
                return title

        return ""

    def class_name(self) -> str:
        return str(self.info.get("bundle_id") or self.info.get("class_name") or self.info.get("app_name") or "")

    def process_id(self) -> int:
        return int(self.info.get("pid") or 0)

    def bundle_id(self) -> str:
        return str(self.info.get("bundle_id") or "")

    def app_name(self) -> str:
        return str(self.info.get("app_name") or "")

    def is_visible(self) -> bool:
        handle = self.handle
        pid = self.process_id()
        title = self.window_text()
        bundle_id = self.bundle_id()
        for info in enumerate_macos_windows():
            if handle and int(info.get("handle") or 0) == handle:
                self.info.update(info)
                return True
            if (
                pid
                and bundle_id
                and int(info.get("pid") or 0) == pid
                and str(info.get("bundle_id") or "") == bundle_id
                and (
                    not title
                    or _clean_field(str(info.get("title") or "")) == title
                )
            ):
                self.info.update(info)
                return True
            if int(info.get("pid") or 0) == pid and str(info.get("title") or "") == title:
                self.info.update(info)
                return True
        return False

    def set_focus(self) -> None:
        bundle_id = self.bundle_id()
        app_name = self.app_name()
        if not bundle_id and not app_name:
            raise MacAutomationError("활성화할 앱 식별자를 찾지 못했습니다.")
        app_activate = (
            f'tell application id {_quote_applescript_text(bundle_id)} to activate'
            if bundle_id
            else f'tell application {_quote_applescript_text(app_name)} to activate'
        )
        script = f'''
{app_activate}
delay 0.1
tell application "System Events"
    tell first application process whose unix id is {self.process_id()}
        try
            perform action "AXRaise" of (first window whose name is {_quote_applescript_text(self.window_text())})
        end try
    end tell
end tell
'''
        _run_osascript(script, timeout=10)

    def child_window(self, control_type=None, found_index=0):
        if control_type in ("Tree", "List", None):
            return MacTreeControl(self)
        raise MacAutomationError(f"지원하지 않는 macOS child_window 요청: {control_type}")

_publish_context(globals())

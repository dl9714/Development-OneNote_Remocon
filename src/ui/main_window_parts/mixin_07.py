# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin07:

    def _codex_onenote_templates_macos(
        self, values: Dict[str, str]
    ) -> Dict[str, str]:
        target = values.get("target", "") or "현재 선택된 작업 위치"
        title = values.get("title", "") or "(제목 없음)"
        body = values.get("body", "") or "(본문 없음)"

        def _manual(title_text: str, body_lines: List[str]) -> str:
            content = "\n".join(f"- {line}" for line in body_lines if line)
            return f"# OneNote macOS: {title_text}\n# 대상: {target}\n{content}\n"

        return {
            "add_page": _manual(
                "페이지 추가",
                [
                    "Microsoft OneNote for Mac을 연다.",
                    f"왼쪽 패널에서 `{target}` 위치를 찾는다. OneNote 조회 ON이면 해당 경로를 우선 사용한다.",
                    "페이지 목록의 `페이지 추가` 버튼을 눌러 새 페이지를 만든다.",
                    f"새 페이지 제목을 `{title}`로 입력한다.",
                    f"본문을 아래 내용으로 붙여넣는다.\n\n{body}",
                    "생성 후 같은 제목 페이지가 보이는지 확인한다.",
                ],
            ),
            "add_section": _manual(
                "새 섹션 생성",
                [
                    f"현재 전자필기장 또는 그룹 `{target}`을 연다.",
                    "섹션 추가 버튼을 눌러 새 섹션을 만든다.",
                    f"섹션 이름을 `{title}`로 지정한다.",
                    "생성 후 왼쪽 섹션 목록에 같은 이름이 나타나는지 확인한다.",
                ],
            ),
            "add_section_group": _manual(
                "새 섹션 그룹 생성",
                [
                    f"대상 전자필기장 `{target}`에서 섹션 그룹을 만들 위치를 연다.",
                    "macOS에서는 리본/컨텍스트 메뉴 UI 자동화로 섹션 그룹을 생성한다.",
                    f"새 그룹 이름을 `{title}`로 입력한다.",
                    "생성 후 왼쪽 패널에서 그룹이 보이는지 확인한다.",
                ],
            ),
            "add_notebook": _manual(
                "새 전자필기장 생성",
                [
                    f"새 전자필기장 저장 경로 또는 이름: `{target}`",
                    "OneNote for Mac의 `파일 > 새 전자 필기장` 흐름으로 생성한다.",
                    "OneDrive 또는 로컬 저장 위치를 명확히 확인한다.",
                    "생성 후 해당 전자필기장이 열린 상태인지 확인한다.",
                ],
            ),
            "list_hierarchy": _manual(
                "계층 구조 조회",
                [
                    "현재 열린 OneNote for Mac 창에서 전자필기장/섹션 목록을 읽는다.",
                    "결과는 JSON이나 구조화된 목록으로 정리한다.",
                    "macOS에서는 COM ID 대신 경로 문자열을 우선 식별자로 사용한다.",
                ],
            ),
            "read_section_pages": _manual(
                "섹션 페이지 목록 읽기",
                [
                    f"대상 섹션 경로: `{target}`",
                    f"제목 필터가 있으면 `{title}` 포함 여부로 페이지를 추린다.",
                    "페이지 제목 목록과 보이는 순서를 정리한다.",
                ],
            ),
            "read_page_xml": _manual(
                "페이지 내용 읽기",
                [
                    f"대상 페이지 또는 URL: `{target}`",
                    "macOS에서는 XML 대신 현재 페이지의 텍스트/구조를 읽어 요약하거나 원문을 추출한다.",
                    "제목과 본문이 비어 있지 않은지 확인한다.",
                ],
            ),
            "append_page_body": _manual(
                "페이지 본문 추가",
                [
                    f"대상 페이지: `{target}`",
                    f"추가할 본문:\n\n{body}",
                    "현재 페이지 끝으로 이동해 본문을 추가한다.",
                    "추가 후 마지막 단락이 실제로 반영됐는지 확인한다.",
                ],
            ),
            "rename_page": _manual(
                "페이지 제목 변경",
                [
                    f"대상 페이지: `{target}`",
                    f"새 제목: `{title}`",
                    "페이지 제목 영역을 수정한 뒤 새 제목이 목록과 본문 상단에 모두 반영됐는지 확인한다.",
                ],
            ),
            "open_notebook": _manual(
                "전자필기장 열기",
                [
                    f"대상 경로 또는 URL: `{target}`",
                    "macOS에서는 `onenote:`/웹 링크/파일 경로를 열어 OneNote를 활성화한다.",
                    "열린 뒤 왼쪽 패널에 전자필기장이 표시되는지 확인한다.",
                ],
            ),
            "navigate_to_id": _manual(
                "대상으로 이동",
                [
                    f"대상 입력값: `{target}`",
                    "macOS에서는 ID 대신 경로 또는 `onenote:` URL을 우선 사용한다.",
                    "이동 후 해당 전자필기장/섹션/페이지가 화면에 보이는지 확인한다.",
                ],
            ),
            "find_pages": _manual(
                "페이지 검색",
                [
                    f"검색어: `{title}`",
                    "OneNote for Mac 검색 UI를 열고 검색어를 입력한다.",
                    "검색 결과 페이지 제목과 위치를 최대 50개까지 정리한다.",
                ],
            ),
            "get_object_link": _manual(
                "앱 링크 생성",
                [
                    f"대상 경로 또는 페이지: `{target}`",
                    "macOS에서는 가능한 경우 OneNote 앱 링크/공유 링크를 복사한다.",
                    "직접 ID 링크를 만들 수 없으면 현재 페이지 공유 링크를 대체값으로 사용한다.",
                ],
            ),
            "get_web_link": _manual(
                "웹 링크 생성",
                [
                    f"대상 경로 또는 페이지: `{target}`",
                    "공유/복사 링크 UI로 웹 링크를 가져온다.",
                    "링크 생성 후 클립보드와 결과 텍스트를 함께 확인한다.",
                ],
            ),
            "get_parent_id": _manual(
                "상위 위치 조회",
                [
                    f"대상 경로: `{target}`",
                    "macOS에서는 부모 ID 대신 `전자필기장 > 섹션 그룹 > 섹션` 경로를 반환한다.",
                    "현재 선택 위치의 한 단계 상위 경로를 구조화해서 기록한다.",
                ],
            ),
            "sync_hierarchy": _manual(
                "계층 동기화",
                [
                    f"대상 경로: `{target}`",
                    "OneNote for Mac에서 동기화 상태 메뉴를 열어 현재 전자필기장/페이지를 동기화한다.",
                    "동기화 후 같은 위치를 다시 읽어 접근 가능한지 확인한다.",
                ],
            ),
            "export_pdf": _manual(
                "PDF 내보내기",
                [
                    f"대상 페이지/섹션: `{target}`",
                    f"원하는 저장 경로가 있으면 `{body}`를 우선 사용한다.",
                    "macOS 인쇄/내보내기 흐름으로 PDF를 저장한다.",
                    "저장 후 파일이 실제로 생성됐는지 확인한다.",
                ],
            ),
            "export_mhtml": _manual(
                "MHTML 내보내기",
                [
                    f"대상 페이지/섹션: `{target}`",
                    "OneNote for Mac에 MHTML 직접 내보내기가 없으면 HTML/PDF 대체 경로를 우선 검토한다.",
                    "대체 형식을 사용했다면 결과와 제한점을 함께 보고한다.",
                ],
            ),
            "export_xps": _manual(
                "XPS 내보내기",
                [
                    f"대상 페이지/섹션: `{target}`",
                    "macOS에는 XPS가 기본 형식이 아니므로 PDF 대체를 우선 사용한다.",
                    "정말 XPS가 필요하면 변환 단계를 명시하고 결과를 검증한다.",
                ],
            ),
            "get_special_locations": _manual(
                "특수 위치 조회",
                [
                    "OneNote for Mac의 기본 전자필기장 위치, 최근 위치, 백업에 해당하는 사용자 접근 경로를 정리한다.",
                    "macOS에서는 COM 특수 위치 API 대신 실제 앱/OneDrive 경로를 설명형으로 반환한다.",
                ],
            ),
            "close_notebook": _manual(
                "전자필기장 닫기",
                [
                    f"대상 전자필기장: `{target}`",
                    "OneNote for Mac UI에서 해당 전자필기장을 닫는다.",
                    "닫기 후 왼쪽 패널에서 사라졌는지 확인한다.",
                ],
            ),
            "navigate_to_url": _manual(
                "URL로 이동",
                [
                    f"대상 URL: `{target}`",
                    "macOS 기본 열기 동작으로 URL을 실행한다.",
                    "OneNote 앱 또는 브라우저에서 의도한 위치가 열리는지 확인한다.",
                ],
            ),
        }

    def _codex_onenote_templates_for_platform(
        self,
        platform_key: str,
        values: Optional[Dict[str, str]] = None,
    ) -> Dict[str, str]:
        values = values or self._codex_codegen_values()
        if platform_key == CODEX_PLATFORM_MACOS:
            return self._codex_onenote_templates_macos(values)
        return self._codex_onenote_templates_windows(values)

    def _codex_template_choice_defs(self, platform_key: str) -> List[Tuple[str, str]]:
        windows_choices = [
            ("페이지 추가", "add_page"),
            ("새 섹션 생성", "add_section"),
            ("새 섹션 그룹 생성", "add_section_group"),
            ("새 전자필기장 생성", "add_notebook"),
            ("계층 구조 조회", "list_hierarchy"),
            ("섹션 페이지 목록 읽기", "read_section_pages"),
            ("페이지 XML 읽기", "read_page_xml"),
            ("페이지 본문 추가", "append_page_body"),
            ("페이지 제목 변경", "rename_page"),
            ("전자필기장 열기", "open_notebook"),
            ("ID로 이동", "navigate_to_id"),
            ("페이지 검색", "find_pages"),
            ("링크 생성", "get_object_link"),
            ("웹 링크 생성", "get_web_link"),
            ("부모 ID 조회", "get_parent_id"),
            ("계층 동기화", "sync_hierarchy"),
            ("PDF 내보내기", "export_pdf"),
            ("MHTML 내보내기", "export_mhtml"),
            ("XPS 내보내기", "export_xps"),
            ("특수 위치 조회", "get_special_locations"),
            ("전자필기장 닫기", "close_notebook"),
            ("URL로 이동", "navigate_to_url"),
        ]
        if platform_key != CODEX_PLATFORM_MACOS:
            return windows_choices
        return [
            ("페이지 추가", "add_page"),
            ("새 섹션 생성", "add_section"),
            ("새 섹션 그룹 생성", "add_section_group"),
            ("새 전자필기장 생성", "add_notebook"),
            ("계층 구조 조회", "list_hierarchy"),
            ("섹션 페이지 목록 읽기", "read_section_pages"),
            ("페이지 내용 읽기", "read_page_xml"),
            ("페이지 본문 추가", "append_page_body"),
            ("페이지 제목 변경", "rename_page"),
            ("전자필기장 열기", "open_notebook"),
            ("경로/URL로 이동", "navigate_to_id"),
            ("페이지 검색", "find_pages"),
            ("앱 링크 생성", "get_object_link"),
            ("웹 링크 생성", "get_web_link"),
            ("상위 위치 조회", "get_parent_id"),
            ("계층 동기화", "sync_hierarchy"),
            ("PDF 내보내기", "export_pdf"),
            ("특수 위치 조회", "get_special_locations"),
            ("전자필기장 닫기", "close_notebook"),
            ("URL로 이동", "navigate_to_url"),
        ]

    def _selected_codex_template_platform(self) -> str:
        combo = getattr(self, "codex_template_platform_combo", None)
        if combo is not None:
            key = str(combo.currentData() or "").strip()
            if key:
                return key
        return _codex_active_platform_key()

    def _codex_template_platform_help_text(self, platform_key: str) -> str:
        if platform_key == CODEX_PLATFORM_MACOS:
            return (
                "macOS 스킬은 OneNote for Mac의 왼쪽 패널, 섹션/페이지 구조, "
                "접근성/UI 자동화를 기준으로 작성됩니다. Windows 스킬은 참고/혼합용으로 유지됩니다."
            )
        return (
            "Windows 스킬은 OneNote COM API와 ID 기반 검증을 전제로 합니다. "
            "macOS 스킬은 같은 화면에서 별도로 관리됩니다."
        )

    def _populate_codex_template_combo(self, platform_key: str) -> None:
        combo = getattr(self, "codex_template_combo", None)
        if combo is None:
            return
        previous_key = str(combo.currentData() or "").strip()
        choices = self._codex_template_choice_defs(platform_key)
        combo.blockSignals(True)
        combo.clear()
        selected_index = 0
        for index, (label, key) in enumerate(choices):
            combo.addItem(label, key)
            if key == previous_key:
                selected_index = index
        combo.setCurrentIndex(selected_index)
        combo.blockSignals(False)
        self._set_label_text_if_changed(
            getattr(self, "codex_template_scope_label", None),
            self._codex_template_platform_help_text(platform_key),
        )

    def _on_codex_template_platform_changed(self) -> None:
        self._populate_codex_template_combo(self._selected_codex_template_platform())
        self._update_codex_template_preview()

    def _selected_codex_template_text(self) -> str:
        combo = getattr(self, "codex_template_combo", None)
        if combo is None:
            return ""
        key = combo.currentData()
        return self._codex_onenote_templates_for_platform(
            self._selected_codex_template_platform()
        ).get(key, "")

    def _update_codex_template_preview(self) -> None:
        preview = getattr(self, "codex_template_preview", None)
        if preview is None:
            return
        self._set_plain_text_if_changed(preview, self._selected_codex_template_text())

    def _set_plain_text_if_changed(self, widget: Optional[QTextEdit], text: str) -> None:
        if widget is None:
            return
        if widget.toPlainText() != text:
            widget.setPlainText(text)

    def _set_label_text_if_changed(self, widget: Optional[QLabel], text: str) -> None:
        if widget is None:
            return
        if widget.text() != text:
            widget.setText(text)

    def _schedule_codex_codegen_previews(self, *args) -> None:
        timer = getattr(self, "_codex_codegen_preview_timer", None)
        if timer is None:
            timer = QTimer(self)
            timer.setSingleShot(True)
            timer.setInterval(120)
            timer.timeout.connect(self._update_codex_codegen_previews)
            self._codex_codegen_preview_timer = timer
        timer.start()

    def _schedule_codex_skill_call_preview(self, *args) -> None:
        timer = getattr(self, "_codex_skill_preview_timer", None)
        if timer is None:
            timer = QTimer(self)
            timer.setSingleShot(True)
            timer.setInterval(120)
            timer.timeout.connect(self._update_codex_skill_call_preview)
            self._codex_skill_preview_timer = timer
        timer.start()

    def _update_codex_codegen_previews(self) -> None:
        self._update_codex_request_preview()
        self._update_codex_template_preview()
        self._update_codex_work_order_preview()
        self._update_codex_skill_package_preview()
        self._update_codex_status_summary()

_publish_context(globals())

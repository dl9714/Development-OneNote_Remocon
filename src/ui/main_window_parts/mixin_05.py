# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin05:

    def _codex_page_reader_request_text(self) -> str:
        try:
            target = self._codex_target_from_fields()
        except Exception:
            target = {}
        path = target.get("path") or "현재 선택된 작업 위치"
        if IS_MACOS:
            return f"""작업:
아래 OneNote for Mac 작업 위치의 페이지 목록과 현재 보이는 페이지를 읽어줘.

작업 위치:
{path}

정리 방식:
- 현재 섹션의 페이지 제목 목록을 먼저 적는다.
- 현재 열려 있는 페이지가 있으면 따로 표시한다.
- 사용자가 이어서 작업할 만한 페이지 후보를 알려준다.
- 특정 페이지 내용을 읽어야 하면 어떤 제목을 골라야 하는지 묻는다.

보고:
- 페이지 수
- 페이지 제목 목록
- 현재 열린 페이지
- 다음에 읽을 만한 페이지 후보
"""
        return f"""작업:
아래 OneNote 작업 위치의 페이지 목록을 읽어줘.

작업 위치:
{path}

정리 방식:
- 페이지 제목 목록을 먼저 적는다.
- 최근 수정된 페이지가 있으면 표시한다.
- 사용자가 이어서 작업할 만한 페이지 후보를 알려준다.
- 특정 페이지 내용을 읽어야 하면 어떤 제목을 골라야 하는지 묻는다.

보고:
- 페이지 수
- 페이지 제목 목록
- 다음에 읽을 만한 페이지 후보
"""

    def _codex_json_from_text(self, text: str) -> Any:
        raw = (text or "").strip()
        if not raw:
            raise ValueError("조회 결과가 비어 있습니다.")
        try:
            return json.loads(raw)
        except Exception:
            starts = [idx for idx in (raw.find("{"), raw.find("[")) if idx >= 0]
            ends = [idx for idx in (raw.rfind("}"), raw.rfind("]")) if idx >= 0]
            if not starts or not ends:
                raise
            return json.loads(raw[min(starts): max(ends) + 1])

    def _codex_text_from_page_xml(self, xml_text: str) -> str:
        if not xml_text:
            return ""
        chunks = re.findall(r"(?is)<[^:>]*:?T[^>]*>(.*?)</[^:>]*:?T>", xml_text)
        if not chunks:
            chunks = [re.sub(r"(?is)<[^>]+>", " ", xml_text)]
        lines = []
        for chunk in chunks:
            plain = re.sub(r"(?is)<[^>]+>", " ", chunk)
            plain = html.unescape(plain)
            plain = re.sub(r"\s+", " ", plain).strip()
            if plain:
                lines.append(plain)
        return "\n".join(lines)

    def _codex_page_reader_result_summary_text(self, text: str = "") -> str:
        data = self._codex_json_from_text(text or QApplication.clipboard().text())
        if not isinstance(data, dict):
            raise ValueError("페이지 읽기 결과 형식이 올바르지 않습니다.")

        pages = data.get("pages", [])
        if not isinstance(pages, list):
            pages = []

        rows = []
        for idx, page in enumerate(pages, start=1):
            if not isinstance(page, dict):
                continue
            rows.append(
                f"| {idx} | {page.get('title', '') or '-'} | "
                f"{page.get('lastModifiedTime', '') or '-'} | `{page.get('id', '')}` |"
            )
        if not rows:
            rows.append("| - | 페이지 없음 | - | - |")

        selected = data.get("selected_page")
        selected_title = ""
        if isinstance(selected, dict):
            selected_title = str(selected.get("title", "") or "")
        selected_text = self._codex_text_from_page_xml(str(data.get("selected_page_xml", "") or ""))

        return "\n".join(
            [
                "# OneNote 페이지 읽기 결과 요약",
                "",
                f"생성 시각: {data.get('generated_at', '') or '-'}",
                f"대상 섹션: {data.get('section_path', '') or data.get('section_id', '') or '-'}",
                f"페이지 수: {data.get('page_count', len(pages))}",
                f"선택 페이지: {selected_title or '-'}",
                "",
                "## 페이지 목록",
                "",
                "| 번호 | 제목 | 수정 시각 | Page ID |",
                "| ---: | --- | --- | --- |",
                *rows,
                "",
                "## 선택 페이지 텍스트",
                "",
                selected_text or "- 선택된 페이지 XML이 없거나 텍스트를 추출하지 못했습니다.",
                "",
            ]
        )

    def _copy_codex_page_reader_result_summary_to_clipboard(self) -> None:
        try:
            QApplication.clipboard().setText(self._codex_page_reader_result_summary_text())
            try:
                self.connection_status_label.setText("OneNote 페이지 읽기 결과 요약을 클립보드에 복사했습니다.")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "페이지 읽기 결과 요약 실패", str(e))

    def _append_codex_page_reader_result_to_request_body(self) -> None:
        body_editor = getattr(self, "codex_request_body_editor", None)
        if body_editor is None:
            return
        try:
            summary = self._codex_page_reader_result_summary_text()
        except Exception as e:
            QMessageBox.warning(self, "페이지 읽기 결과 추가 실패", str(e))
            return
        current = body_editor.toPlainText().rstrip()
        body_editor.setPlainText((current + "\n\n" + summary).strip())
        self._update_codex_codegen_previews()
        try:
            self.connection_status_label.setText("페이지 읽기 결과 요약을 요청 본문에 추가했습니다.")
        except Exception:
            pass

_publish_context(globals())

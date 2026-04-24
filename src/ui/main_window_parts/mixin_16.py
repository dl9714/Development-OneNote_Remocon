# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin16:

    def _replace_codex_request_body_from_clipboard(self) -> None:
        body_editor = getattr(self, "codex_request_body_editor", None)
        if body_editor is None:
            return
        clip = QApplication.clipboard().text().strip()
        if not clip:
            QMessageBox.information(self, "클립보드 비어 있음", "요청 본문으로 교체할 클립보드 텍스트가 없습니다.")
            return
        body_editor.setPlainText(clip)
        self._update_codex_codegen_previews()
        try:
            self.connection_status_label.setText("클립보드 텍스트로 코덱스 요청 본문을 교체했습니다.")
        except Exception:
            pass

    def _codex_compact_prompt_text(self) -> str:
        try:
            target = self._codex_target_from_fields()
        except Exception:
            target = {}

        try:
            request = self._codex_request_draft_payload().get("request", {})
        except Exception:
            request = {}

        order_no = ""
        skill_name = ""
        try:
            order_no = self.codex_skill_order_input.text().strip()
            skill_name = self.codex_skill_name_input.text().strip()
        except Exception:
            pass

        return f"""아래 OneNote 작업을 바로 수행해라.

스킬:
- 주문번호: {order_no or "미지정"}
- 스킬명: {skill_name or "미지정"}

작업:
- 유형: {request.get("action", "") or "미지정"}
- 제목/이름: {request.get("title", "") or "미지정"}
- 대상 경로: {request.get("target", "") or target.get("path", "") or "미지정"}

본문:
{request.get("body", "") or "- 사용자가 제공한 내용을 OneNote에 정리한다."}

처리 기준:
- 내부 처리 방식은 코덱스 전용 지침을 따른다.
- 사용자 스킬은 글쓰기 형식과 내용 정리에만 적용한다.

완료 후 변경한 항목과 검증 결과만 간단히 보고해라.
"""

    def _copy_codex_compact_prompt_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_compact_prompt_text())
        try:
            self.connection_status_label.setText("짧은 작업 요청을 한국어로 복사했습니다.")
        except Exception:
            pass

    def _codex_review_prompt_text(self) -> str:
        return f"""아래 OneNote 작업 계획을 검토해라. 실행하지 말고 위험 요소, 빠진 확인, 잘못된 대상 위치 가능성만 점검해라.

## 현재 상태

{self._codex_status_summary_text()}

## 실행 체크리스트

{self._codex_execution_checklist_text()}

## 요청문

{self._codex_request_text()}

검토 결과는 다음 형식으로 짧게 작성해라.
- 실행 전 막아야 할 문제:
- 대상 위치 확인 필요:
- 검증 방법:
"""

    def _copy_codex_review_prompt_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_review_prompt_text())
        try:
            self.connection_status_label.setText("검토 요청을 한국어로 복사했습니다.")
        except Exception:
            pass

    def _codex_task_breakdown_prompt_text(self) -> str:
        return f"""아래 OneNote 작업을 실행 가능한 하위 작업으로 분해해라. 아직 실행하지 말고 순서, 필요한 조회, 검증 기준만 작성해라.

## 현재 요청

{self._codex_request_text()}

출력 형식:
- 목표:
- 선행 조회:
- 작업 순서:
- 검증:
- 실패 시 복구:
"""

    def _copy_codex_task_breakdown_prompt_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_task_breakdown_prompt_text())
        try:
            self.connection_status_label.setText("작업 단계 정리 요청을 한국어로 복사했습니다.")
        except Exception:
            pass

    def _codex_completion_report_template_text(self) -> str:
        try:
            request = self._codex_request_draft_payload().get("request", {})
        except Exception:
            request = {}
        return f"""# OneNote 작업 완료 보고 양식

## 요청

- 작업: {request.get("action", "") or "미지정"}
- 제목/이름: {request.get("title", "") or "미지정"}
- 대상: {request.get("target", "") or "미지정"}

## 수행한 작업

-

## 변경한 OneNote 항목

- 전자필기장:
- 섹션 그룹:
- 섹션:
- 페이지:

## 검증 결과

- 확인 방식: 코덱스 전용 지침 기준
- 확인한 값:
- 결과:

## 남은 확인 사항

-
"""

    def _copy_codex_completion_report_template_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_completion_report_template_text())
        try:
            self.connection_status_label.setText("OneNote 작업 완료 보고 양식을 한국어로 복사했습니다.")
        except Exception:
            pass

    def _codex_current_request_text_blob(self) -> str:
        parts: List[str] = []
        try:
            parts.append(self.codex_request_action_combo.currentText())
            parts.append(self.codex_request_target_input.text())
            parts.append(self.codex_request_title_input.text())
            parts.append(self.codex_request_body_editor.toPlainText())
        except Exception:
            pass
        return "\n".join(p for p in parts if p)

    def _codex_request_tokens(self) -> Set[str]:
        text = unicodedata.normalize("NFKC", self._codex_current_request_text_blob()).casefold()
        return {
            token
            for token in re.findall(r"[0-9A-Za-z가-힣_]{2,}", text)
            if token not in {"one", "note", "onenote", "com", "api"}
        }

    def _rank_codex_skill_records(self) -> List[Dict[str, Any]]:
        if not getattr(self, "_codex_skill_records", None):
            self._refresh_codex_skill_list()

        request_key = self._codex_skill_search_key(self._codex_current_request_text_blob())
        request_tokens = self._codex_request_tokens()
        ranked: List[Dict[str, Any]] = []

        for record in getattr(self, "_codex_skill_records", []):
            text = " ".join(
                [
                    record.get("order", ""),
                    record.get("name", ""),
                    record.get("filename", ""),
                    record.get("trigger", ""),
                ]
            )
            try:
                path = record.get("path", "")
                if path and os.path.exists(path):
                    text += "\n" + self._codex_skill_metadata_from_file(
                        path, record.get("name", "")
                    ).get("body", "")
            except Exception:
                pass

            skill_key = self._codex_skill_search_key(text)
            skill_tokens = {
                token.casefold()
                for token in re.findall(r"[0-9A-Za-z가-힣_]{2,}", unicodedata.normalize("NFKC", text))
            }
            hits = sorted(request_tokens & skill_tokens)
            score = len(hits) * 10
            if record.get("name") and self._codex_skill_search_key(record.get("name", "")) in request_key:
                score += 25
            if record.get("trigger") and self._codex_skill_search_key(record.get("trigger", "")) in request_key:
                score += 20
            if record.get("order") and self._codex_skill_search_key(record.get("order", "")) in request_key:
                score += 50
            if score > 0:
                ranked.append(
                    {
                        "score": score,
                        "record": record,
                        "hits": hits[:8],
                    }
                )

        ranked.sort(
            key=lambda item: (
                -int(item.get("score", 0)),
                item.get("record", {}).get("order", "") or "ZZZ",
                item.get("record", {}).get("name", ""),
            )
        )
        return ranked

    def _select_best_codex_skill_recommendation(self) -> None:
        ranked = self._rank_codex_skill_records()
        if not ranked:
            QMessageBox.information(self, "추천 스킬 없음", "현재 요청에 맞는 추천 스킬을 찾지 못했습니다.")
            return
        record = ranked[0].get("record", {})
        order_no = record.get("order", "")
        if not order_no:
            QMessageBox.information(self, "추천 스킬 선택 실패", "추천된 스킬에 주문번호가 없습니다.")
            return
        lookup = getattr(self, "codex_skill_order_lookup_input", None)
        if lookup is not None:
            lookup.setText(order_no)
        self._select_codex_skill_by_order_input()
        try:
            self.connection_status_label.setText(
                f"추천 스킬 선택: {order_no} / {record.get('name', '')}"
            )
        except Exception:
            pass

_publish_context(globals())

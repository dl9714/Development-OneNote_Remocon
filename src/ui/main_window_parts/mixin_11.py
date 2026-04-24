# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin11:

    def _codex_context_pack_text(self) -> str:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")

        try:
            page_reader_summary = self._codex_page_reader_result_summary_text()
        except Exception:
            page_reader_summary = ""

        try:
            skill_call = self._codex_skill_call_prompt_text()
        except Exception as e:
            skill_call = f"스킬 적용 요청 생성 실패: {e}"

        try:
            request_text = self._codex_request_text()
        except Exception:
            request_text = ""

        try:
            work_order_text = self._codex_work_order_text()
        except Exception:
            work_order_text = ""

        try:
            checklist_text = self._codex_execution_checklist_text()
        except Exception:
            checklist_text = ""

        try:
            compact_prompt = self._codex_compact_prompt_text()
        except Exception:
            compact_prompt = ""

        try:
            review_prompt = self._codex_review_prompt_text()
        except Exception:
            review_prompt = ""

        try:
            breakdown_prompt = self._codex_task_breakdown_prompt_text()
        except Exception:
            breakdown_prompt = ""

        try:
            completion_report = self._codex_completion_report_template_text()
        except Exception:
            completion_report = ""

        try:
            status_text = self._codex_status_summary_text()
        except Exception:
            status_text = ""

        try:
            current_skill = self._current_codex_skill_markdown()
        except Exception:
            current_skill = ""

        skill_index = ""
        try:
            self._refresh_codex_skill_list()
            with open(self._codex_skill_order_index_path(), "r", encoding="utf-8") as f:
                skill_index = f.read().strip()
        except Exception:
            skill_index = ""

        return f"""# 코덱스 작업 자료 묶음
생성 시각: {timestamp}

이 문서는 다음 코덱스 작업에 필요한 사용자 요청, 작업 위치, 사용자 스킬 자료를 묶은 자료입니다.

## 바로 실행할 지시

아래 사용자 요청과 작업 주문서를 기준으로 OneNote 작업을 수행해라.

## 내부 처리 기준

OneNote 조작 방식과 검증 기준은 코덱스 전용 지침에서 필요한 때 직접 확인한다.
사용자에게 붙여 넣는 요청에는 내부 구현 절차를 포함하지 않는다.

## 짧은 작업 요청

```text
{compact_prompt}
```

## 검토 요청

```text
{review_prompt}
```

## 작업 단계 정리 요청

```text
{breakdown_prompt}
```

## 완료 보고 양식

```markdown
{completion_report}
```

## 스킬 적용 요청

{skill_call}

## 페이지 읽기 결과 요약

```markdown
{page_reader_summary}
```

## 현재 상태

```text
{status_text}
```

## 요청문

```text
{request_text}
```

## 작업 주문서

````markdown
{work_order_text}
````

## 실행 체크리스트

```text
{checklist_text}
```

## 현재 스킬 초안

````markdown
{current_skill}
````

## 스킬 주문번호표

````markdown
{skill_index}
````
"""

    def _update_codex_context_pack_preview(self) -> None:
        preview = getattr(self, "codex_context_pack_preview", None)
        if preview is None:
            return
        self._set_plain_text_if_changed(preview, self._codex_context_pack_text())

    def _copy_codex_context_pack_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_context_pack_text())
        try:
            self.connection_status_label.setText(
                "코덱스 작업 자료 묶음을 한국어로 복사했습니다."
            )
        except Exception:
            pass

    def _save_codex_context_pack(self) -> None:
        text = self._codex_context_pack_text()
        stamp = time.strftime("%Y%m%d-%H%M%S")
        filename = f"{stamp}-codex-work-materials.md"
        try:
            path = os.path.join(self._codex_requests_dir(), filename)
            self._write_text_file_atomic(path, text)
            try:
                self.connection_status_label.setText(
                    f"코덱스 작업 자료 묶음 저장 완료: {path}"
                )
            except Exception:
                pass
            self._refresh_codex_work_order_list(path)
            QMessageBox.information(self, "작업 자료 묶음 저장 완료", path)
        except Exception as e:
            QMessageBox.warning(self, "작업 자료 묶음 저장 실패", str(e))

    def _codex_markdown_file_count(self, folder: str) -> int:
        try:
            if not os.path.isdir(folder):
                return 0
            folder_mtime = os.path.getmtime(folder)
            cache = getattr(self, "_codex_markdown_file_count_cache", {})
            cached = cache.get(folder)
            if cached and cached[0] == folder_mtime:
                return cached[1]

            count = sum(
                1
                for filename in os.listdir(folder)
                if filename.lower().endswith(".md")
                and filename not in ("README.md", "skill-order-index.md", "skill-audit.md")
            )
            cache[folder] = (folder_mtime, count)
            self._codex_markdown_file_count_cache = cache
            return count
        except Exception:
            return 0

    def _codex_status_summary_text(self) -> str:
        skill_count = self._codex_markdown_file_count(self._codex_skills_dir())
        request_count = self._codex_markdown_file_count(self._codex_requests_dir())
        target_count = len(self._load_codex_targets())

        action = ""
        title = ""
        request_target = ""
        try:
            action = self.codex_request_action_combo.currentText().strip()
            title = self.codex_request_title_input.text().strip()
            request_target = self.codex_request_target_input.text().strip()
        except Exception:
            pass

        skill_order = ""
        skill_name = ""
        try:
            skill_order = self.codex_skill_order_input.text().strip()
            skill_name = self.codex_skill_name_input.text().strip()
        except Exception:
            pass

        target = {}
        try:
            target = self._codex_target_from_fields()
        except Exception:
            target = {}

        draft_state = "있음" if os.path.exists(self._codex_request_draft_path()) else "없음"
        self._codex_last_status_snapshot = {
            "skill_count": skill_count,
            "request_count": request_count,
            "target_count": target_count,
            "draft_state": draft_state,
            "target": target,
            "action": action,
            "title": title,
            "request_target": request_target,
            "skill_order": skill_order,
            "skill_name": skill_name,
        }

        return f"""코덱스 현재 상태

- 스킬 파일: {skill_count}개
- 저장된 작업 주문서/자료 묶음: {request_count}개
- 저장된 작업 위치: {target_count}개
- 요청 초안: {draft_state}

현재 요청
- 작업: {action or "미지정"}
- 제목/이름: {title or "미지정"}
- 요청 대상: {request_target or target.get("path", "") or "미지정"}

선택 스킬
- 주문번호: {skill_order or "미지정"}
- 스킬명: {skill_name or "미지정"}

작업 위치
- 위치 이름: {target.get("name", "") or "미지정"}
- 전자필기장: {target.get("notebook", "") or "미지정"}
- 섹션 그룹: {target.get("section_group", "") or "미지정"}
- 섹션: {target.get("section", "") or "미지정"}
"""

    def _update_codex_status_summary(self) -> None:
        preview = getattr(self, "codex_status_summary_preview", None)
        summary_text = self._codex_status_summary_text()
        if preview is not None:
            self._set_plain_text_if_changed(preview, summary_text)

        try:
            snapshot = getattr(self, "_codex_last_status_snapshot", {})
            skill_count = snapshot.get("skill_count", 0)
            request_count = snapshot.get("request_count", 0)
            target_count = snapshot.get("target_count", 0)
            draft_state = (
                "저장됨" if snapshot.get("draft_state") == "있음" else "대기"
            )
            target = snapshot.get("target", {})
            action = snapshot.get("action", "")
            title = snapshot.get("title", "")
            skill_order = snapshot.get("skill_order", "")
            skill_name = snapshot.get("skill_name", "")

            hero_values = {
                "codex_metric_target_value": f"{target_count}개",
                "codex_metric_draft_value": draft_state,
                "codex_metric_skill_value": f"{skill_count}개",
                "codex_metric_order_value": f"{request_count}개",
                "codex_hero_target_value": (
                    target.get("path")
                    or target.get("name")
                    or "작업 위치 미지정"
                ),
                "codex_hero_request_value": title or action or "요청 대기 중",
                "codex_copy_target_value": (
                    target.get("path")
                    or target.get("name")
                    or "작업 위치 미지정"
                ),
                "codex_copy_skill_value": (
                    " / ".join(
                        part for part in (skill_order, skill_name) if part
                    )
                    or "선택 스킬 미지정"
                ),
                "codex_copy_request_value": title or action or "요청 대기 중",
            }
            for attr, value in hero_values.items():
                widget = getattr(self, attr, None)
                if widget is not None:
                    self._set_label_text_if_changed(widget, value)
        except Exception:
            pass

    def _copy_codex_status_summary_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_status_summary_text())
        try:
            self.connection_status_label.setText("코덱스 현재 상태 요약을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _codex_current_target_copy_text(self) -> str:
        try:
            target = self._codex_target_from_fields()
        except Exception:
            target = {}
        path = target.get("path", "") or target.get("name", "")
        parts = [
            f"작업 위치: {path or '미지정'}",
            f"전자필기장: {target.get('notebook', '') or '미지정'}",
            f"섹션 그룹: {target.get('section_group', '') or '미지정'}",
            f"섹션: {target.get('section', '') or '미지정'}",
        ]
        return "\n".join(parts)

    def _copy_codex_current_target_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_current_target_copy_text())
        try:
            self.connection_status_label.setText("현재 작업 위치를 클립보드에 복사했습니다.")
        except Exception:
            pass

_publish_context(globals())

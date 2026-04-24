# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin10:

    def _codex_skill_category_from_parts(self, *parts: Any) -> str:
        key = self._codex_skill_search_key(*parts)
        category_keywords = [
            (
                "OneNote/도구",
                (
                    "onenote",
                    "원노트",
                    "전자필기장",
                    "페이지",
                    "섹션",
                    "섹션그룹",
                    "코덱스 스킬",
                ),
            ),
            (
                "기록/정리",
                (
                    "회의록",
                    "업무일지",
                    "글쓰기 형식",
                    "아이디어 정리",
                    "학습 노트",
                    "프로젝트 회고",
                    "의사결정",
                    "문제해결 로그",
                    "버그 리포트",
                    "실험 검증",
                    "작업 체크리스트",
                    "장애 대응",
                    "회고 액션플랜",
                    "회의 후속작업",
                ),
            ),
            (
                "기획/전략",
                (
                    "기획서",
                    "요구사항",
                    "제안서",
                    "로드맵",
                    "OKR",
                    "정책 초안",
                    "콘텐츠 기획",
                    "온보딩",
                    "교육 자료",
                    "FAQ",
                    "학습 계획",
                ),
            ),
            (
                "분석/리서치",
                (
                    "리서치",
                    "데이터 분석",
                    "경쟁사",
                    "구매 비교",
                    "벤치마크",
                    "고객 피드백",
                    "고객 인터뷰",
                    "비용 최적화",
                ),
            ),
            (
                "운영/품질",
                (
                    "계약 검토",
                    "주간 보고",
                    "릴리즈 노트",
                    "운영 매뉴얼",
                    "리스크 관리",
                    "품질 점검",
                    "자동화 아이디어",
                ),
            ),
            (
                "소통/협업",
                (
                    "이메일 공지",
                    "세일즈 콜",
                    "인터뷰 질문지",
                    "의사소통 메시지",
                    "회의 아젠다",
                    "채용 후보자 평가",
                ),
            ),
        ]
        for category, keywords in category_keywords:
            for keyword in keywords:
                if self._codex_skill_search_key(keyword) in key:
                    return category
        return "기타"

    def _populate_codex_skill_list(self, records: List[Dict[str, str]]) -> None:
        skill_list = getattr(self, "codex_skill_list", None)
        if skill_list is None:
            return
        current_path = self._selected_codex_skill_path()
        skill_list.blockSignals(True)
        skill_list.clear()
        category_counts: Dict[str, int] = {}
        for skill in records:
            category = skill.get("category") or "기타"
            category_counts[category] = category_counts.get(category, 0) + 1
        current_category = ""
        for skill in records:
            category = skill.get("category") or "기타"
            if category != current_category:
                header = QListWidgetItem(
                    f"카테고리 · {category} ({category_counts.get(category, 0)})"
                )
                header.setFlags(Qt.ItemFlag.ItemIsEnabled)
                header.setForeground(QBrush(QColor("#7c879b")))
                skill_list.addItem(header)
                current_category = category
            order_no = skill.get("order", "")
            name = skill.get("name", "")
            label = f"[{order_no}] {name}" if order_no else name
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, skill.get("path", ""))
            skill_list.addItem(item)
        if current_path:
            for i in range(skill_list.count()):
                item = skill_list.item(i)
                if item.data(Qt.ItemDataRole.UserRole) == current_path:
                    skill_list.setCurrentRow(i)
                    break
        skill_list.blockSignals(False)

    def _schedule_filter_codex_skill_list(self, *args) -> None:
        timer = getattr(self, "_codex_skill_filter_timer", None)
        if timer is None:
            timer = QTimer(self)
            timer.setSingleShot(True)
            timer.setInterval(120)
            timer.timeout.connect(self._filter_codex_skill_list)
            self._codex_skill_filter_timer = timer
        timer.start()

    def _filter_codex_skill_list(self) -> None:
        records = getattr(self, "_codex_skill_records", [])
        search_input = getattr(self, "codex_skill_search_input", None)
        query = search_input.text().strip() if search_input is not None else ""
        key = self._codex_skill_search_key(query)
        if not key:
            self._populate_codex_skill_list(records)
            return
        filtered = [
            skill
            for skill in records
            if key
            in self._codex_skill_search_key(
                skill.get("category", ""),
                skill.get("order", ""),
                skill.get("name", ""),
                skill.get("filename", ""),
                skill.get("trigger", ""),
            )
        ]
        self._populate_codex_skill_list(filtered)

    def _selected_codex_skill_path(self) -> str:
        skill_list = getattr(self, "codex_skill_list", None)
        item = skill_list.currentItem() if skill_list is not None else None
        if item is None:
            return ""
        return item.data(Qt.ItemDataRole.UserRole) or ""

    def _current_codex_skill_markdown(self) -> str:
        return self._codex_skill_markdown(
            self.codex_skill_name_input.text(),
            self.codex_skill_trigger_input.text(),
            self.codex_skill_body_editor.toPlainText(),
            self.codex_skill_order_input.text(),
        )

    def _codex_skill_call_prompt_text(self) -> str:
        def _line_text(attr: str) -> str:
            widget = getattr(self, attr, None)
            if widget is None:
                return ""
            try:
                return widget.text().strip()
            except Exception:
                return ""

        def _plain_text(attr: str) -> str:
            widget = getattr(self, attr, None)
            if widget is None:
                return ""
            try:
                return widget.toPlainText().strip()
            except Exception:
                return ""

        order_no = _line_text("codex_skill_order_input") or self._codex_next_skill_order()
        name = _line_text("codex_skill_name_input") or "선택 스킬"
        trigger = _line_text("codex_skill_trigger_input")
        target = ""
        try:
            target = self._codex_target_from_fields().get("path", "")
        except Exception:
            target = ""
        request_title = _line_text("codex_request_title_input")
        request_body = _plain_text("codex_request_body_editor")

        return f"""아래 사용자 스킬을 적용해서 OneNote 작업을 수행해라.

스킬:
- 주문번호: {order_no}
- 이름: {name}
- 적용 조건: {trigger or "현재 사용자가 요청한 OneNote 작업"}

작업 위치:
{target or "현재 앱에서 선택한 OneNote 작업 위치를 먼저 확인한다."}

사용자 요청:
{request_title or name}

추가 내용:
{request_body or "- 현재 선택된 스킬의 Instructions를 우선 적용한다."}

처리 기준:
- 사용자 스킬 파일에서는 `## Instructions`만 작업에 맞게 적용한다.
- OneNote 내부 조작 방식은 코덱스 전용 지침에서 필요한 때 직접 확인한다.
"""

    def _update_codex_skill_call_preview(self) -> None:
        preview = getattr(self, "codex_skill_call_preview", None)
        if preview is None:
            self._update_codex_work_order_preview()
            self._update_codex_status_summary()
            return
        self._set_plain_text_if_changed(preview, self._codex_skill_call_prompt_text())
        self._update_codex_work_order_preview()
        self._update_codex_status_summary()

    def _copy_codex_skill_call_prompt_to_clipboard(self) -> None:
        text = self._codex_skill_call_prompt_text()
        QApplication.clipboard().setText(text)
        try:
            self.connection_status_label.setText("스킬 적용 요청을 한국어로 복사했습니다.")
        except Exception:
            pass

    def _open_selected_codex_skill_file(self) -> None:
        path = self._selected_codex_skill_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 스킬을 선택하세요.")
            return
        try:
            open_path_in_system(path)
        except Exception as e:
            QMessageBox.warning(self, "스킬 파일 열기 실패", str(e))

    def _open_codex_skills_folder(self) -> None:
        try:
            os.makedirs(self._codex_skills_dir(), exist_ok=True)
            open_path_in_system(self._codex_skills_dir())
        except Exception as e:
            QMessageBox.warning(self, "스킬 폴더 열기 실패", str(e))

    def _open_codex_requests_folder(self) -> None:
        try:
            os.makedirs(self._codex_requests_dir(), exist_ok=True)
            open_path_in_system(self._codex_requests_dir())
        except Exception as e:
            QMessageBox.warning(self, "주문서 폴더 열기 실패", str(e))

    def _codex_work_order_text(self) -> str:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        order_no = ""
        skill_name = ""
        skill_path = self._selected_codex_skill_path()
        try:
            order_no = self.codex_skill_order_input.text().strip()
            skill_name = self.codex_skill_name_input.text().strip()
        except Exception:
            pass

        try:
            target = self._codex_target_from_fields()
        except Exception:
            target = {}

        try:
            request_text = self._codex_request_text()
        except Exception:
            request_text = ""

        return f"""# 코덱스 작업 주문서

생성 시각: {timestamp}

## 사용자 스킬

- 주문번호: {order_no or "미선택"}
- 스킬명: {skill_name or "미선택"}
- 파일: `{skill_path or "미선택"}`

## 작업 위치

- 위치 이름: {target.get("name", "")}
- 작업 경로: {target.get("path", "")}
- 전자필기장: {target.get("notebook", "")}
- 섹션 그룹: {target.get("section_group", "")}
- 섹션: {target.get("section", "")}

## 요청문

```text
{request_text}
```
"""

    def _update_codex_work_order_preview(self) -> None:
        preview = getattr(self, "codex_work_order_preview", None)
        if preview is None:
            return
        self._set_plain_text_if_changed(preview, self._codex_work_order_text())

    def _copy_codex_work_order_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_work_order_text())
        try:
            self.connection_status_label.setText("코덱스 작업 주문서를 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _save_codex_work_order(self) -> None:
        text = self._codex_work_order_text()
        order_no = self.codex_skill_order_input.text().strip() or "NO-SKILL"
        name = self.codex_skill_name_input.text().strip() or "codex-work-order"
        stamp = time.strftime("%Y%m%d-%H%M%S")
        filename = f"{stamp}-{order_no}-{self._codex_skill_slug(name)}.md"
        try:
            path = os.path.join(self._codex_requests_dir(), filename)
            self._write_text_file_atomic(path, text)
            try:
                self.connection_status_label.setText(f"코덱스 작업 주문서 저장 완료: {path}")
            except Exception:
                pass
            self._refresh_codex_work_order_list(path)
            QMessageBox.information(self, "작업 주문서 저장 완료", path)
        except Exception as e:
            QMessageBox.warning(self, "작업 주문서 저장 실패", str(e))

_publish_context(globals())

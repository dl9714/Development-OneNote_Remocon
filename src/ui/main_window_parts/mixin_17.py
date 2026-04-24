# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin17:

    def _codex_request_presets(self) -> Dict[str, Dict[str, str]]:
        today = time.strftime("%Y-%m-%d")
        return {
            "quick_note": {
                "name": "빠른 메모",
                "action": "페이지 추가",
                "title": f"{today} 빠른 메모",
                "body": """핵심:
-

세부 내용:
-

다음 행동:
-
""",
            },
            "meeting": {
                "name": "회의록",
                "action": "나의 기본 글쓰기 형식으로 페이지 작성",
                "title": f"{today} 회의록",
                "body": """회의명:
참석자:

논의 요약:
-

결정 사항:
-

액션 아이템:
- 담당자 / 마감 / 확인 방법:

보류/리스크:
-
""",
            },
            "daily_log": {
                "name": "업무일지",
                "action": "나의 기본 글쓰기 형식으로 페이지 작성",
                "title": f"{today} 업무일지",
                "body": """오늘 한 일:
-

결정한 것:
-

막힌 것:
-

다음 행동:
-
""",
            },
            "project_plan": {
                "name": "작업 계획",
                "action": "페이지 추가",
                "title": "작업 계획",
                "body": """목표:
-

범위:
- 포함:
- 제외:

실행 순서:
1.
2.
3.

검증:
-

리스크:
-
""",
            },
            "cleanup": {
                "name": "OneNote 정리",
                "action": "나의 기본 글쓰기 형식으로 페이지 작성",
                "title": "OneNote 정리 계획",
                "body": """현재 위치:
-

바꿀 위치:
-

정리할 항목:
-

보존할 항목:
-

검증 방법:
- 작업 전후 구조 확인
""",
            },
            "weekly_review": {
                "name": "주간 회고",
                "action": "페이지 추가",
                "title": f"{today} 주간 회고",
                "body": """이번 주 완료:
-

이번 주 미완료:
-

배운 점:
-

다음 주 우선순위:
1.
2.
3.
""",
            },
        }

    def _apply_codex_request_preset(self) -> None:
        combo = getattr(self, "codex_request_preset_combo", None)
        if combo is None:
            return
        preset = self._codex_request_presets().get(combo.currentData())
        if not preset:
            return

        action_combo = getattr(self, "codex_request_action_combo", None)
        if action_combo is not None:
            idx = action_combo.findText(preset.get("action", ""))
            if idx >= 0:
                action_combo.setCurrentIndex(idx)

        title_input = getattr(self, "codex_request_title_input", None)
        if title_input is not None:
            title_input.setText(preset.get("title", ""))

        body_editor = getattr(self, "codex_request_body_editor", None)
        if body_editor is not None:
            body_editor.setPlainText(preset.get("body", ""))

        self._update_codex_codegen_previews()
        try:
            self.connection_status_label.setText(
                f"요청 양식 적용: {preset.get('name', '')}"
            )
        except Exception:
            pass

    def _new_codex_skill_from_current_request(self) -> None:
        required = (
            "codex_skill_order_input",
            "codex_skill_name_input",
            "codex_skill_trigger_input",
            "codex_skill_body_editor",
        )
        if any(getattr(self, attr, None) is None for attr in required):
            return

        action = ""
        title = ""
        target = ""
        body = ""
        try:
            action = self.codex_request_action_combo.currentText().strip()
            title = self.codex_request_title_input.text().strip()
            target = self.codex_request_target_input.text().strip()
            body = self.codex_request_body_editor.toPlainText().strip()
        except Exception:
            pass

        skill_name = f"{title or action or 'OneNote 작업'} 반복 스킬"
        trigger = (
            f"{action or 'OneNote 작업'}을(를) {target or '현재 작업 위치'}에서 "
            "반복해서 수행할 때"
        )
        skill_body = f"""목표:
- 아래 요청 유형을 같은 기준으로 반복 수행한다.

반복 요청:
- 작업: {action or '미지정'}
- 대상: {target or '미지정'}
- 제목/이름: {title or '미지정'}

본문 기준:
{body or '- 사용자가 제공한 본문을 유지하되 OneNote에 읽기 쉬운 구조로 작성한다.'}

처리 방식:
- 작업 위치를 먼저 확인한다.
- 사용자 요청의 형식과 출력 기준을 먼저 정한다.
- 실제 OneNote 조작 방식은 코덱스 전용 지침을 따른다.

출력:
- OneNote에 반영한 항목
- 검증 결과
- 사용자가 다음에 확인할 항목
"""

        self.codex_skill_order_input.setText(self._codex_next_skill_order())
        self.codex_skill_name_input.setText(skill_name)
        self.codex_skill_trigger_input.setText(trigger)
        self.codex_skill_body_editor.setPlainText(skill_body)
        self._update_codex_skill_call_preview()
        try:
            self._scroll_codex_to_widget("codex_skill_editor_widget")
        except Exception:
            pass
        try:
            self.connection_status_label.setText("현재 요청으로 스킬 초안을 만들었습니다.")
        except Exception:
            pass

    def _build_codex_request_group(self) -> QWidget:
        group = QWidget()
        group.setObjectName("CodexCard")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(8)

        header = QLabel("✍️ 작업 내용 작성")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        form = QGridLayout()
        form.setVerticalSpacing(6)
        form.setHorizontalSpacing(6)
        form.setColumnMinimumWidth(0, 48)
        form.setColumnStretch(1, 1)

        self.codex_request_preset_combo = WheelSafeComboBox()
        self._configure_codex_lookup_combo(self.codex_request_preset_combo)
        for key, preset in self._codex_request_presets().items():
            self.codex_request_preset_combo.addItem(preset.get("name", key), key)
        form.addWidget(QLabel("양식"), 0, 0)
        preset_row = QHBoxLayout()
        preset_row.setSpacing(6)
        preset_row.addWidget(self.codex_request_preset_combo, stretch=1)
        preset_apply_btn = QToolButton()
        preset_apply_btn.setText("적용")
        preset_apply_btn.clicked.connect(self._apply_codex_request_preset)
        preset_row.addWidget(preset_apply_btn)
        form.addLayout(preset_row, 0, 1)

        self.codex_request_action_combo = WheelSafeComboBox()
        self._configure_codex_lookup_combo(self.codex_request_action_combo)
        self.codex_request_action_combo.addItems(
            ["페이지 추가", "새 섹션 생성", "섹션 그룹 생성", "기본 형식 작성"]
        )
        self.codex_request_action_combo.currentIndexChanged.connect(self._schedule_codex_codegen_previews)
        form.addWidget(QLabel("작업"), 1, 0)
        form.addWidget(self.codex_request_action_combo, 1, 1)

        self.codex_request_target_input = QLineEdit()
        self.codex_request_target_input.setPlaceholderText("전자필기장 > 섹션그룹 > 섹션")
        self.codex_request_target_input.textChanged.connect(self._schedule_codex_codegen_previews)
        form.addWidget(QLabel("경로"), 2, 0)
        form.addWidget(self.codex_request_target_input, 2, 1)

        self.codex_request_title_input = QLineEdit()
        self.codex_request_title_input.setPlaceholderText("항목의 이름을 입력하세요...")
        self.codex_request_title_input.textChanged.connect(self._schedule_codex_codegen_previews)
        form.addWidget(QLabel("제목"), 3, 0)
        form.addWidget(self.codex_request_title_input, 3, 1)

        layout.addLayout(form)

        self.codex_request_body_editor = QTextEdit()
        self.codex_request_body_editor.setPlaceholderText("구체적인 작업 지시 내용을 입력하세요. (Shift+Enter로 줄바꿈)")
        self.codex_request_body_editor.setMinimumHeight(130)
        self.codex_request_body_editor.textChanged.connect(self._schedule_codex_codegen_previews)
        layout.addWidget(self.codex_request_body_editor)

        actions = QHBoxLayout()
        actions.setSpacing(8)

        copy_btn = QToolButton()
        copy_btn.setText("🚀 작업요청 복사")
        copy_btn.setProperty("variant", "primary")
        copy_btn.setMinimumHeight(38)
        copy_btn.setMinimumWidth(0)
        copy_btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        copy_btn.clicked.connect(self._copy_codex_request_to_clipboard)

        draft_btn = QToolButton()
        draft_btn.setText("💾 초안")
        draft_btn.setMinimumHeight(38)
        draft_btn.setFixedWidth(66)
        draft_btn.clicked.connect(self._save_codex_request_draft)

        load_btn = QToolButton()
        load_btn.setText("📂 불러오기")
        load_btn.setMinimumHeight(38)
        load_btn.setFixedWidth(86)
        load_btn.clicked.connect(self._load_codex_request_draft)

        actions.addWidget(copy_btn)
        actions.addWidget(draft_btn)
        actions.addWidget(load_btn)
        layout.addLayout(actions)

        tool_grid = QGridLayout()
        tool_grid.setHorizontalSpacing(6)
        tool_grid.setVerticalSpacing(6)
        tool_specs = [
            ("📎 붙인내용 추가", self._append_clipboard_to_codex_request_body),
            ("♻️ 붙인내용 교체", self._replace_codex_request_body_from_clipboard),
            ("✔️ 실행순서 복사", self._copy_codex_execution_checklist_to_clipboard),
            ("🧩 짧은요청 복사", self._copy_codex_compact_prompt_to_clipboard),
            ("🛡️ 검토요청 복사", self._copy_codex_review_prompt_to_clipboard),
            ("🪜 단계정리 복사", self._copy_codex_task_breakdown_prompt_to_clipboard),
            ("📣 보고양식 복사", self._copy_codex_completion_report_template_to_clipboard),
            ("🎯 맞는스킬 선택", self._select_recommended_codex_skill),
            ("🧠 스킬로 만들기", self._new_codex_skill_from_current_request),
        ]
        for index, (text, cb) in enumerate(tool_specs):
            btn = QToolButton()
            btn.setText(text)
            btn.setMinimumHeight(28)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            btn.clicked.connect(cb)
            tool_grid.addWidget(btn, index // 2, index % 2)
        layout.addLayout(tool_grid)

        layout.addWidget(QLabel("생성된 요청문 미리보기"))
        self.codex_request_preview = QTextEdit()
        self.codex_request_preview.setReadOnly(True)
        self.codex_request_preview.setMinimumHeight(130)
        layout.addWidget(self.codex_request_preview)
        self._update_codex_request_preview()

        return group

_publish_context(globals())

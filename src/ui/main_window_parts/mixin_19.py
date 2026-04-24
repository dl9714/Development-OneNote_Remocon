# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin19:


    def _build_codex_internal_instructions_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(8)

        header = QLabel("코덱스 전용 지침")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        help_label = QLabel(
            (
                "OneNote for Mac 접근성/UI 자동화, 대상 경로 판정, 안전 실행 순서, 검증 기준처럼 "
                "사용자가 매번 볼 필요 없는 실행 전제를 관리합니다."
                if IS_MACOS
                else "OneNote COM 사용 방식, 대상 판정, 안전 실행 순서, 검증 기준처럼 사용자가 매번 볼 필요 없는 실행 전제를 관리합니다."
            )
        )
        help_label.setObjectName("CodexPageSubtitle")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        self._ensure_codex_internal_instructions_file()
        self.codex_internal_instructions_editor = QTextEdit()
        self.codex_internal_instructions_editor.setPlainText(
            self._codex_internal_instructions_text()
        )
        self.codex_internal_instructions_editor.setMinimumHeight(300)
        layout.addWidget(self.codex_internal_instructions_editor)

        actions = QHBoxLayout()
        actions.setSpacing(8)

        save_btn = QToolButton()
        save_btn.setText("저장")
        save_btn.setProperty("variant", "primary")
        save_btn.clicked.connect(self._save_codex_internal_instructions)

        reload_btn = QToolButton()
        reload_btn.setText("다시 불러오기")
        reload_btn.clicked.connect(self._reload_codex_internal_instructions)

        copy_btn = QToolButton()
        copy_btn.setText("복사")
        copy_btn.clicked.connect(self._copy_codex_internal_instructions_to_clipboard)

        open_btn = QToolButton()
        open_btn.setText("폴더 열기")
        open_btn.clicked.connect(self._open_codex_instructions_folder)

        for btn in (save_btn, reload_btn, copy_btn, open_btn):
            btn.setMinimumHeight(30)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            actions.addWidget(btn, stretch=1)

        layout.addLayout(actions)
        return card


    def _build_codex_status_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 8, 10, 10)
        layout.setSpacing(8)

        header = QLabel("📋 현재 선택값 복사")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        def _copy_row(label_text: str, value_attr: str, button_text: str, callback) -> None:
            row = QHBoxLayout()
            row.setSpacing(4)

            label = QLabel(label_text)
            label.setMinimumWidth(58)
            row.addWidget(label)

            value = QLabel("미지정")
            value.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            value.setObjectName("CodexHeroMetaValue")
            value.setWordWrap(True)
            setattr(self, value_attr, value)
            row.addWidget(value, stretch=1)

            btn = QToolButton()
            btn.setText(button_text)
            btn.setMinimumHeight(28)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            btn.clicked.connect(callback)
            row.addWidget(btn)
            layout.addLayout(row)

        _copy_row(
            "작업 위치",
            "codex_copy_target_value",
            "위치 복사",
            self._copy_codex_current_target_to_clipboard,
        )
        _copy_row(
            "선택 스킬",
            "codex_copy_skill_value",
            "스킬 복사",
            self._copy_codex_current_skill_to_clipboard,
        )
        _copy_row(
            "현재 요청",
            "codex_copy_request_value",
            "요청 복사",
            self._copy_codex_current_request_to_clipboard,
        )

        return card


    def _build_codex_quick_tools_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 8, 10, 10)
        layout.setSpacing(8)

        header = QLabel("⚡ 빠른 실행 도구")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        tools_layout = QGridLayout()
        tools_layout.setSpacing(6)

        tool_specs = [
            ("📝 주문서 복사", self._copy_codex_work_order_to_clipboard, "primary"),
            ("🚀 스킬요청 복사", self._copy_codex_skill_call_prompt_to_clipboard, "secondary"),
            ("📦 자료묶음 복사", self._copy_codex_context_pack_to_clipboard, "secondary"),
            ("📄 페이지목록 요청", self._copy_codex_page_reader_script_to_clipboard, ""),
            ("🛠️ 위치조회 요청", self._copy_codex_onenote_inventory_script_to_clipboard, ""),
        ]

        for i, (text, cb, variant) in enumerate(tool_specs):
            btn = QToolButton()
            btn.setText(text)
            btn.setMinimumHeight(28)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            if variant: btn.setProperty("variant", variant)
            btn.clicked.connect(cb)
            tools_layout.addWidget(btn, i // 2, i % 2)

        layout.addLayout(tools_layout)

        return card


    def _build_codex_context_pack_group(self) -> QWidget:
        card = QWidget()
        card.setObjectName("CodexCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 8, 10, 10)
        layout.setSpacing(8)

        header = QLabel("📦 작업 자료 묶음")
        header.setObjectName("CodexCardTitle")
        layout.addWidget(header)

        self.codex_context_pack_preview = QTextEdit()
        self.codex_context_pack_preview.setReadOnly(True)
        self.codex_context_pack_preview.setMinimumHeight(96)
        self.codex_context_pack_preview.setPlaceholderText("작업 자료 묶음 미리보기가 여기에 표시됩니다.")
        layout.addWidget(self.codex_context_pack_preview)

        actions = QHBoxLayout()
        actions.setSpacing(6)
        copy_btn = QToolButton()
        copy_btn.setText("📋 자료묶음 복사")
        copy_btn.setProperty("variant", "secondary")
        copy_btn.clicked.connect(self._copy_codex_context_pack_to_clipboard)

        save_btn = QToolButton()
        save_btn.setText("💾 자료묶음 저장")
        save_btn.clicked.connect(self._save_codex_context_pack)

        refresh_btn = QToolButton()
        refresh_btn.setText("🔄 미리보기 갱신")
        refresh_btn.clicked.connect(self._update_codex_context_pack_preview)

        for btn in (copy_btn, save_btn, refresh_btn):
            btn.setMinimumWidth(0)
            btn.setMinimumHeight(28)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            actions.addWidget(btn, stretch=1)
        layout.addLayout(actions)

        return card


    def _codex_skill_package_default_codex_skills_by_platform(
        self,
    ) -> Dict[str, List[str]]:
        return {
            CODEX_PLATFORM_WINDOWS: [
                "페이지 추가",
                "전자필기장 추가",
                "전자필기장 삭제",
                "섹션 추가",
                "섹션 그룹 추가",
                "페이지 읽기",
                "페이지 XML 읽기",
                "페이지 본문 추가",
                "페이지 제목 변경",
                "섹션 페이지 목록 읽기",
                "계층 구조 조회",
                "위치 조회",
                "전자필기장 열기",
                "ID로 이동",
                "페이지 검색",
                "링크 생성",
                "웹 링크 생성",
                "부모 ID 조회",
                "계층 동기화",
                "PDF 내보내기",
                "MHTML 내보내기",
                "XPS 내보내기",
                "특수 위치 조회",
                "전자필기장 닫기",
                "URL로 이동",
            ],
            CODEX_PLATFORM_MACOS: [
                "페이지 추가",
                "전자필기장 추가",
                "섹션 추가",
                "섹션 그룹 추가",
                "페이지 읽기",
                "페이지 본문 추가",
                "페이지 제목 변경",
                "섹션 페이지 목록 읽기",
                "계층 구조 조회",
                "위치 조회",
                "전자필기장 열기",
                "페이지 검색",
                "앱 링크 생성",
                "웹 링크 생성",
                "상위 위치 조회",
                "계층 동기화",
                "PDF 내보내기",
                "특수 위치 조회",
                "전자필기장 닫기",
                "URL로 이동",
            ],
        }

    def _codex_skill_package_default_codex_skills(
        self, platform_key: Optional[str] = None
    ) -> List[str]:
        platform_key = platform_key or _codex_active_platform_key()
        return list(
            self._codex_skill_package_default_codex_skills_by_platform().get(
                platform_key, []
            )
        )

    def _codex_skill_package_default_instructions(self) -> List[str]:
        common = [
            "사용자 요청에서 목표, 대상, 출력 형식, 금지 조건을 먼저 분리",
            "사용자 스킬은 글쓰기/정리 형식에만 적용",
            "OneNote 조작은 코덱스 전용 지침과 작업별 내부 문서를 우선",
            "명시 ID 또는 저장된 위치 후보를 우선 확인",
            "삭제, 덮어쓰기, 대량 작업은 영향 범위 확인 후 실행",
            "실패 시 단계, 대상, 추정 원인, 다음 확인 값을 보고",
        ]
        if _codex_active_platform_key() == CODEX_PLATFORM_MACOS:
            return common + [
                "OneNote for Mac 접근성/UI 자동화를 우선 사용",
                "현재 보이는 전자필기장/섹션/페이지 경로를 먼저 확인",
                "쓰기 직후 왼쪽 패널과 현재 본문에서 결과를 다시 검증",
                "ID 대신 경로/URL/현재 선택 위치를 우선 식별자로 사용",
            ]
        return common + [
            "OneNote COM API 우선, 화면 클릭 자동화는 최후 수단",
            "명시 ID, 저장 대상 ID, 위치 캐시, 제한 계층 조회 순서로 대상 확정",
            "쓰기 직후 GetHierarchy 또는 GetPageContent로 자동 검증",
            "ID 실패 시 전체 탐색보다 상위 대상부터 제한 재조회",
        ]

    def _codex_text_lines_from_widget(self, widget: Optional[QTextEdit]) -> List[str]:
        if widget is None:
            return []
        return [
            line.strip().lstrip("-").strip()
            for line in widget.toPlainText().splitlines()
            if line.strip().lstrip("-").strip()
        ]

    def _set_codex_text_lines(self, widget: Optional[QTextEdit], lines: List[str]) -> None:
        if widget is None:
            return
        widget.setPlainText("\n".join(lines or []))

    def _codex_skill_records_from_files(self) -> List[Dict[str, str]]:
        skills_dir = self._codex_skills_dir()
        os.makedirs(skills_dir, exist_ok=True)
        rows: List[Dict[str, str]] = []
        for filename in sorted(os.listdir(skills_dir)):
            if not filename.lower().endswith(".md"):
                continue
            if filename in ("README.md", "skill-order-index.md", "skill-audit.md"):
                continue
            path = os.path.join(skills_dir, filename)
            order_no = ""
            name = filename[:-3]
            trigger = ""
            try:
                meta = self._codex_skill_metadata_from_file(path, name)
                order_no = meta.get("order", "")
                name = meta.get("name", name)
                trigger = meta.get("trigger", "")
            except Exception:
                pass
            category = self._codex_skill_category_from_parts(order_no, name, filename, trigger)
            rows.append(
                {
                    "category": category,
                    "order": order_no,
                    "name": name,
                    "filename": filename,
                    "trigger": trigger,
                    "path": path,
                }
            )
        category_order = {
            name: index for index, name in enumerate(self._codex_skill_category_names())
        }
        rows.sort(
            key=lambda s: (
                category_order.get(s.get("category") or "기타", len(category_order)),
                s.get("order") or "ZZZ",
                s.get("name") or "",
            )
        )
        return rows

    def _codex_checked_list_values(self, list_widget: Optional[QListWidget]) -> List[str]:
        values: List[str] = []
        if list_widget is None:
            return values
        for i in range(list_widget.count()):
            item = list_widget.item(i)
            if item.checkState() != Qt.CheckState.Checked:
                continue
            value = str(item.data(Qt.ItemDataRole.UserRole) or item.text()).strip()
            if value and value not in values:
                values.append(value)
        return values

_publish_context(globals())

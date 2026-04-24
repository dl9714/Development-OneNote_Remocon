# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


def codex_tab_stylesheet(codex_font_stack: str) -> str:
    from src.ui.main_window_parts.codex_tab_style import (
        codex_tab_stylesheet as _codex_tab_stylesheet,
    )

    return _codex_tab_stylesheet(codex_font_stack)


class MainWindowMixin24:

    def _build_codex_tab(self, section: str) -> QWidget:
        root = QWidget()
        root.setObjectName("CodexRoot")
        codex_font_stack = _platform_ui_font_stack()
        root.setStyleSheet(codex_tab_stylesheet(codex_font_stack))
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        top_bar = QWidget()
        top_bar.setObjectName("CodexTopBar")
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(2, 2, 2, 2)
        top_layout.setSpacing(1)

        top_nav_scroll = QScrollArea()
        top_nav_scroll.setObjectName("CodexTopNavScroll")
        top_nav_scroll.setWidgetResizable(True)
        top_nav_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        top_nav_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        top_nav_scroll.setFixedHeight(30)

        top_nav_content = QWidget()
        top_nav_layout = QHBoxLayout(top_nav_content)
        top_nav_layout.setContentsMargins(0, 0, 0, 0)
        top_nav_layout.setSpacing(1)
        top_nav_scroll.setWidget(top_nav_content)
        top_layout.addWidget(top_nav_scroll, stretch=1)

        root_layout.addWidget(top_bar)

        stacked_widget = QStackedWidget()
        stacked_widget.setObjectName("CodexStackedWidget")
        stack_attr = (
            "codex_harness_stacked_widget"
            if section == "harness"
            else "codex_remocon_stacked_widget"
        )
        buttons_attr = (
            "_codex_harness_nav_buttons"
            if section == "harness"
            else "_codex_remocon_nav_buttons"
        )
        setattr(self, stack_attr, stacked_widget)
        setattr(self, buttons_attr, [])

        def switch_page(idx: int, btn: QToolButton):
            stacked_widget.setCurrentIndex(idx)
            for b in getattr(self, buttons_attr, []):
                b.setChecked(False)
            btn.setChecked(True)

        def add_nav_item(text: str, index: int) -> QToolButton:
            btn = QToolButton()
            btn.setObjectName("CodexTopNavButton")
            btn.setText(text)
            btn.setCheckable(True)
            btn.setMinimumHeight(26)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            btn.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
            btn.clicked.connect(lambda checked, i=index, b=btn: switch_page(i, b))
            top_nav_layout.addWidget(btn, stretch=1)
            getattr(self, buttons_attr).append(btn)
            return btn

        def make_scroll_page(
            eyebrow: str,
            title: str,
            subtitle: str,
            *,
            show_header: bool = True,
        ):
            page = QScrollArea()
            page.setObjectName("CodexPageScroll")
            page.setWidgetResizable(True)
            page.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            page.setViewportMargins(0, 0, 0, 0)
            content = QWidget()
            content.setObjectName("CodexPageContent")
            content.setMinimumWidth(0)
            content.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
            layout = QVBoxLayout(content)
            if show_header:
                layout.setContentsMargins(8, 8, 8, 10)
                layout.setSpacing(8)
            else:
                layout.setContentsMargins(4, 4, 4, 6)
                layout.setSpacing(4)

            if show_header:
                eyebrow_label = QLabel(eyebrow)
                eyebrow_label.setObjectName("CodexPageEyebrow")
                layout.addWidget(eyebrow_label)

                title_label = QLabel(title)
                title_label.setObjectName("CodexPageTitle")
                layout.addWidget(title_label)

                subtitle_label = QLabel(subtitle)
                subtitle_label.setObjectName("CodexPageSubtitle")
                subtitle_label.setWordWrap(True)
                layout.addWidget(subtitle_label)

            page.setWidget(content)
            return page, layout

        def make_metric(label: str, value: str):
            tile = QWidget()
            tile.setObjectName("CodexMetricTile")
            tile_layout = QVBoxLayout(tile)
            tile_layout.setContentsMargins(7, 5, 7, 5)
            tile_layout.setSpacing(2)

            label_widget = QLabel(label)
            label_widget.setObjectName("CodexMetricLabel")
            value_widget = QLabel(value)
            value_widget.setObjectName("CodexMetricValue")
            value_widget.setMinimumHeight(20)

            tile_layout.addWidget(label_widget)
            tile_layout.addWidget(value_widget)
            return tile, value_widget

        pages = []

        if section == "harness":
            page_package, package_layout = make_scroll_page(
                "SKILL PACKAGE",
                "스킬 패키지",
                (
                    "사용자 스킬과 Windows/macOS 코덱스 실행 자료를 하나의 요청 묶음으로 관리합니다."
                    if IS_MACOS
                    else "사용자 스킬과 코덱스 실행 자료를 하나의 요청 묶음으로 관리합니다."
                ),
            )
            self.codex_skill_package_widget = self._build_codex_skill_package_group()
            package_layout.addWidget(self.codex_skill_package_widget)
            package_layout.addStretch(1)
            pages.append((page_package, "스킬 패키지", "스킬 패키지"))

            page_user_skills, user_skills_layout = make_scroll_page(
                "USER SKILLS",
                "사용자 스킬",
                "글쓰기 형태와 에이전트 역할처럼 결과물의 형식과 역할을 관리합니다.",
                show_header=False,
            )
            self.codex_skill_editor_widget = self._build_codex_skill_editor_group()
            user_skills_layout.addWidget(self.codex_skill_editor_widget, stretch=1)
            user_skills_layout.addStretch(1)
            pages.append((page_user_skills, "사용자 스킬", "사용자 스킬"))

            page_codex_skills, codex_skills_layout = make_scroll_page(
                "CODEX SKILLS",
                "코덱스 스킬",
                (
                    "Windows OneNote COM 스킬과 macOS OneNote 화면/UI 스킬을 나눠 관리합니다."
                    if IS_MACOS
                    else "페이지 추가, 전자필기장 추가 같은 OneNote 실행 템플릿을 관리합니다."
                ),
            )
            self.codex_template_group_widget = self._build_codex_template_group()
            codex_skills_layout.addWidget(self.codex_template_group_widget)
            codex_skills_layout.addStretch(1)
            pages.append((page_codex_skills, "코덱스 스킬", "코덱스 스킬"))

            page_instructions, instructions_layout = make_scroll_page(
                "CODEX INSTRUCTIONS",
                "코덱스 지침",
                (
                    "OneNote for Mac 접근성/UI 우선, 경로/현재 선택 위치 우선, 안전 실행 순서, 자동 검증 기준을 관리합니다."
                    if IS_MACOS
                    else "OneNote COM API 우선, 대상 ID 우선, 안전 실행 순서, 자동 검증 기준을 관리합니다."
                ),
            )
            self.codex_internal_instructions_widget = self._build_codex_internal_instructions_group()
            instructions_layout.addWidget(self.codex_internal_instructions_widget)
            instructions_layout.addStretch(1)
            pages.append((page_instructions, "코덱스 지침", "코덱스 지침"))
        else:
            page_dashboard, dashboard_layout = make_scroll_page(
                "COMMAND HOME",
                "명령홈",
                "Codex에 보낼 작업 위치, 요청, 주문서를 한 화면에서 정리합니다.",
            )

            hero = QWidget()
            hero.setObjectName("CodexHeroBand")
            hero_layout = QVBoxLayout(hero)
            hero_layout.setContentsMargins(10, 8, 10, 8)
            hero_layout.setSpacing(6)

            hero_top = QHBoxLayout()
            hero_title_area = QVBoxLayout()
            hero_title_area.setSpacing(4)
            hero_title = QLabel("현재 작업")
            hero_title.setObjectName("CodexHeroTitle")
            hero_subtitle = QLabel("작업 위치와 요청을 바로 복사합니다.")
            hero_subtitle.setObjectName("CodexHeroSubtitle")
            hero_subtitle.setWordWrap(True)
            hero_title_area.addWidget(hero_title)
            hero_title_area.addWidget(hero_subtitle)
            hero_top.addLayout(hero_title_area, stretch=1)

            self.codex_hero_request_value = QLabel("요청 대기 중")
            self.codex_hero_request_value.setObjectName("CodexStatusPill")
            self.codex_hero_request_value.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.codex_hero_request_value.setMinimumWidth(72)
            self.codex_hero_request_value.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
            hero_top.addWidget(self.codex_hero_request_value)
            hero_layout.addLayout(hero_top)

            target_row = QHBoxLayout()
            target_row.setSpacing(4)
            target_label = QLabel("대상")
            target_label.setObjectName("CodexHeroMetaLabel")
            target_row.addWidget(target_label)
            self.codex_hero_target_value = QLabel("작업 위치 미선택")
            self.codex_hero_target_value.setObjectName("CodexHeroMetaValue")
            self.codex_hero_target_value.setWordWrap(True)
            target_row.addWidget(self.codex_hero_target_value, stretch=1)
            hero_layout.addLayout(target_row)

            metrics = QHBoxLayout()
            metrics.setSpacing(4)
            target_tile, self.codex_metric_target_value = make_metric("작업 위치", "0개")
            draft_tile, self.codex_metric_draft_value = make_metric("요청 초안", "대기")
            skill_tile, self.codex_metric_skill_value = make_metric("스킬 수", "0개")
            order_tile, self.codex_metric_order_value = make_metric("작업 주문서", "0개")
            for tile in (target_tile, draft_tile, skill_tile, order_tile):
                metrics.addWidget(tile)
            hero_layout.addLayout(metrics)
            dashboard_layout.addWidget(hero)

            self.codex_status_summary_widget = self._build_codex_status_group()
            dashboard_layout.addWidget(self.codex_status_summary_widget)
            self.codex_quick_tools_widget = self._build_codex_quick_tools_group()
            dashboard_layout.addWidget(self.codex_quick_tools_widget)
            self.codex_work_order_widget = self._build_codex_work_order_group()
            dashboard_layout.addWidget(self.codex_work_order_widget)
            dashboard_layout.addStretch(1)
            pages.append((page_dashboard, "명령홈", "명령홈"))

            page_request, request_layout = make_scroll_page(
                "COMPOSITION",
                "작업요청",
                "위에서 작업 위치를 정하고 아래에서 지시 내용을 작성하세요.",
            )

            self.codex_target_group_widget = self._build_codex_target_group()
            self.codex_request_group_widget = self._build_codex_request_group()

            request_layout.addWidget(self.codex_target_group_widget)
            request_layout.addWidget(self.codex_request_group_widget)
            request_layout.addStretch(1)
            self._apply_codex_target_to_request()
            self._update_codex_request_preview()
            pages.append((page_request, "작업요청", "작업 요청"))

            page_history, history_layout = make_scroll_page(
                "HISTORY",
                "기록",
                "저장된 작업 주문서 기록을 검색하고 다시 불러옵니다.",
                show_header=False,
            )

            self.codex_work_order_history_widget = self._build_codex_work_order_history_group()
            history_layout.addWidget(self.codex_work_order_history_widget, stretch=1)
            history_layout.addStretch(1)
            pages.append((page_history, "기록", "작업 기록"))

        nav_buttons = []
        for index, (page, text, tooltip) in enumerate(pages):
            stacked_widget.addWidget(page)
            btn = add_nav_item(text, index)
            btn.setToolTip(tooltip)
            nav_buttons.append(btn)

        self.codex_splitter = None
        root_layout.addWidget(stacked_widget, stretch=1)

        self._update_codex_work_order_preview()
        self._update_codex_context_pack_preview()
        self._update_codex_status_summary()
        if nav_buttons:
            switch_page(0, nav_buttons[0])

        return root

_publish_context(globals())

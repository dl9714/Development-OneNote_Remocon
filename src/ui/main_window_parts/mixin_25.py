# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin25:

    def _show_onenote_harness_help(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("원노트 하네스 도움말")
        dialog.resize(780, 680)
        help_font_stack = _platform_ui_font_stack()
        help_html_font_stack = _platform_ui_font_stack(include_generic=True)
        dialog.setStyleSheet(
            """
            QDialog {
                background-color: #111316;
                color: #E2E2E6;
                font-family: __FONT_STACK__;
            }
            QTextEdit {
                background-color: #0C0E11;
                border: 1px solid #2F3338;
                border-radius: 8px;
                padding: 0px;
            }
            QPushButton {
                background-color: #A6D854;
                color: #223600;
                border: none;
                border-radius: 6px;
                padding: 8px 18px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #C1F56D;
            }
            QPushButton:pressed {
                background-color: #95C743;
            }
            """
            .replace("__FONT_STACK__", help_font_stack)
        )

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        help_view = QTextEdit(dialog)
        help_view.setReadOnly(True)
        help_view.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        help_codex_skill_desc = (
            "Codex가 실제로 수행할 OneNote for Mac 작업입니다. Windows 스킬과 분리해서 관리합니다."
            if IS_MACOS
            else "Codex가 실제로 수행할 OneNote 작업입니다."
        )
        help_codex_skill_tags = (
            """
                        <span class="tag">macOS 스킬</span>
                        <span class="tag">왼쪽 패널 기준</span>
                        <span class="tag">섹션/페이지 UI</span>
            """
            if IS_MACOS
            else """
                        <span class="tag">페이지 추가</span>
                        <span class="tag">전자필기장 추가</span>
                        <span class="tag">전자필기장 삭제</span>
            """
        )
        help_instruction_primary = (
            "OneNote for Mac 접근성/UI 우선"
            if IS_MACOS
            else "OneNote COM API 우선"
        )
        help_instruction_secondary = (
            "경로/현재 선택 위치 우선" if IS_MACOS else "대상 ID 우선"
        )
        help_view.setHtml(
            """
            <html>
            <head>
            <style>
                body {
                    margin: 0;
                    padding: 0;
                    background: #0C0E11;
                    color: #E2E2E6;
                    font-family: __FONT_STACK__;
                    font-size: 13px;
                    line-height: 1.55;
                }
                .wrap {
                    padding: 22px;
                }
                .hero {
                    background: #1A1C1F;
                    border: 1px solid #2F3338;
                    border-radius: 8px;
                    padding: 20px;
                    margin-bottom: 16px;
                }
                .eyebrow {
                    color: #66D9CC;
                    font-size: 11px;
                    font-weight: 800;
                    margin-bottom: 6px;
                }
                h1 {
                    color: #E2E2E6;
                    font-size: 26px;
                    margin: 0 0 8px 0;
                }
                .lead {
                    color: #C3C9B3;
                    margin: 0;
                }
                .section {
                    background: #15171A;
                    border-left: 4px solid #A6D854;
                    border-radius: 8px;
                    padding: 16px;
                    margin: 12px 0;
                }
                h2 {
                    color: #C1F56D;
                    font-size: 18px;
                    margin: 0 0 8px 0;
                }
                p {
                    margin: 0 0 10px 0;
                    color: #C3C9B3;
                }
                .tags {
                    margin-top: 8px;
                }
                .tag {
                    display: inline-block;
                    background: #282A2D;
                    color: #E2E2E6;
                    border: 1px solid #3B4046;
                    border-radius: 6px;
                    padding: 5px 8px;
                    margin: 3px 4px 3px 0;
                    font-weight: 700;
                }
                .tag-accent {
                    background: #223600;
                    color: #C1F56D;
                    border-color: #A6D854;
                }
                .flow {
                    margin: 0;
                    padding-left: 20px;
                    color: #C3C9B3;
                }
                .flow li {
                    margin: 7px 0;
                }
                .note {
                    background: #0C0E11;
                    border: 1px solid #2F3338;
                    border-radius: 8px;
                    padding: 12px;
                    color: #8D937F;
                    margin-top: 12px;
                }
                code {
                    color: #66D9CC;
                    background: #111316;
                    padding: 2px 5px;
                    border-radius: 4px;
                }
            </style>
            </head>
            <body>
            <div class="wrap">
                <div class="hero">
                    <div class="eyebrow">ONENOTE HARNESS</div>
                    <h1>원노트 하네스</h1>
                    <p class="lead">
                        사용자 스킬, 코덱스 스킬, 실행 지침을 하나의 패키지로 묶어
                        OneNote 작업 요청을 빠르게 만들고 안전하게 검증하는 작업 공간입니다.
                    </p>
                </div>

                <div class="section">
                    <h2>스킬 패키지</h2>
                    <p>
                        사용자 스킬들을 조합한 템플릿입니다. 자주 쓰는 글쓰기 방식,
                        에이전트 역할, 실행 작업을 하나의 구성으로 묶어 Codex 요청에 바로 붙일 수 있습니다.
                    </p>
                    <div class="tags">
                        <span class="tag tag-accent">템플릿</span>
                        <span class="tag">여러 사용자 스킬</span>
                        <span class="tag">여러 코덱스 스킬</span>
                        <span class="tag">지침 묶음</span>
                    </div>
                </div>

                <div class="section">
                    <h2>사용자 스킬</h2>
                    <p>결과물의 형태와 에이전트 역할을 정합니다.</p>
                    <div class="tags">
                        <span class="tag">글쓰기 형태</span>
                        <span class="tag">에이전트 역할</span>
                    </div>
                </div>

                <div class="section">
                    <h2>코덱스 스킬</h2>
                    <p>__CODEX_SKILL_DESC__</p>
                    <div class="tags">
__CODEX_SKILL_TAGS__
                    </div>
                </div>

                <div class="section">
                    <h2>코덱스 지침</h2>
                    <p>Codex가 OneNote 작업을 안전하게 실행하기 위한 내부 실행 기준입니다.</p>
                    <div class="tags">
                        <span class="tag tag-accent">__PRIMARY_INSTRUCTION__</span>
                        <span class="tag">__SECONDARY_INSTRUCTION__</span>
                        <span class="tag">작업별 안전 실행 순서</span>
                        <span class="tag">완료 후 자동 검증</span>
                        <span class="tag">실패 시 단계와 원인 보고</span>
                    </div>
                </div>

                <div class="section">
                    <h2>사용 흐름</h2>
                    <ol class="flow">
                        <li>사용할 사용자 스킬을 고릅니다.</li>
                        <li>실행할 코덱스 스킬을 고릅니다.</li>
                        <li>대상 위치, 제목, 본문 같은 요청 내용을 입력합니다.</li>
                        <li>Codex는 코덱스 지침을 기준으로 실행하고 검증합니다.</li>
                    </ol>
                    <div class="note">
                        저장된 스킬 패키지는 <code>docs/codex/skill-packages</code>에 JSON으로 보관됩니다.
                    </div>
                </div>
            </div>
            </body>
            </html>
            """
            .replace("__FONT_STACK__", help_html_font_stack)
            .replace("__CODEX_SKILL_DESC__", help_codex_skill_desc)
            .replace("__CODEX_SKILL_TAGS__", help_codex_skill_tags)
            .replace("__PRIMARY_INSTRUCTION__", help_instruction_primary)
            .replace("__SECONDARY_INSTRUCTION__", help_instruction_secondary)
        )
        layout.addWidget(help_view)

        close_btn = QPushButton("닫기", dialog)
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignRight)

        dialog.exec()

    def _show_app_info(self) -> None:
        app_path = os.path.abspath(QApplication.applicationFilePath() or sys.executable)
        if IS_MACOS and ".app/" in app_path:
            app_path = app_path.split(".app/", 1)[0] + ".app"
        settings_path = _get_settings_file_path()
        QMessageBox.information(
            self,
            "앱 정보",
            (
                "OneNote Remocon\n\n"
                f"버전: {APP_VERSION}\n"
                f"빌드: {APP_BUILD_VERSION}\n"
                f"플랫폼: {'macOS' if IS_MACOS else 'Windows' if IS_WINDOWS else sys.platform}\n"
                f"실행 경로: {app_path}\n"
                f"설정 JSON: {settings_path}\n"
                f"설정 모드: {_settings_path_mode_label()}"
            ),
        )

_publish_context(globals())

# -*- coding: utf-8 -*-
from __future__ import annotations


def codex_tab_stylesheet(codex_font_stack: str) -> str:
    return """
            QWidget#CodexRoot {
                background-color: #111316;
                color: #E2E2E6;
                font-family: __FONT_STACK__;
            }
            QWidget#CodexSidePanel {
                background-color: #0C0E11;
            }
            QWidget#CodexTopBar {
                background-color: #0C0E11;
            }
            QScrollArea#CodexTopNavScroll {
                border: none;
                background-color: #0C0E11;
            }
            QLabel#CodexBrand {
                color: #E2E2E6;
                font-size: 14pt;
                font-weight: 800;
                padding: 0px;
            }
            QLabel#CodexSidebarCaption {
                color: #8D937F;
                font-size: 8pt;
                font-weight: bold;
            }
            QLabel#CodexStatusPill {
                background-color: #1E2023;
                color: #C1F56D;
                border-radius: 6px;
                padding: 3px 5px;
                font-size: 8pt;
                font-weight: bold;
            }
            QToolButton#CodexTopNavButton {
                background-color: #1E2023;
                color: #E2E2E6;
                border: 1px solid #2F3338;
                border-bottom: 2px solid transparent;
                border-radius: 4px;
                padding: 3px 3px;
                font-size: 8pt;
                font-weight: bold;
            }
            QToolButton#CodexTopNavButton:hover {
                background-color: #282A2D;
                color: #E2E2E6;
            }
            QToolButton#CodexTopNavButton:checked {
                background-color: #A6D854;
                color: #223600;
                border-color: #A6D854;
                border-bottom: 2px solid #C1F56D;
            }
            QToolButton#CodexSideNavButton {
                background-color: transparent;
                color: #C3C9B3;
                border: none;
                border-left: 3px solid transparent;
                border-radius: 0px;
                padding: 8px 8px 8px 12px;
                text-align: left;
                font-size: 10pt;
                font-weight: bold;
            }
            QToolButton#CodexSideNavButton:hover {
                background-color: #1A1C1F;
                color: #E2E2E6;
            }
            QToolButton#CodexSideNavButton:checked {
                background-color: #1A1C1F;
                color: #C1F56D;
                border-left: 3px solid #A6D854;
            }
            QStackedWidget#CodexStackedWidget,
            QWidget#CodexPageContent {
                background-color: #111316;
            }
            QScrollArea#CodexPageScroll {
                border: none;
                background-color: #111316;
            }
            QLabel#CodexPageEyebrow {
                color: #66D9CC;
                font-size: 8pt;
                font-weight: bold;
            }
            QLabel#CodexPageTitle {
                color: #E2E2E6;
                font-size: 17pt;
                font-weight: 800;
            }
            QLabel#CodexPageSubtitle {
                color: #8D937F;
                font-size: 9pt;
            }
            QWidget#CodexHeroBand {
                background-color: #1A1C1F;
                border-radius: 6px;
            }
            QLabel#CodexHeroTitle {
                color: #E2E2E6;
                font-size: 14pt;
                font-weight: 800;
            }
            QLabel#CodexHeroSubtitle,
            QLabel#CodexHeroMetaLabel {
                color: #8D937F;
                font-size: 9pt;
            }
            QLabel#CodexHeroMetaValue {
                color: #C3C9B3;
                font-size: 9pt;
                font-weight: bold;
            }
            QWidget#CodexMetricTile {
                background-color: #0C0E11;
                border-radius: 6px;
            }
            QLabel#CodexMetricLabel {
                color: #8D937F;
                font-size: 8pt;
                font-weight: bold;
            }
            QLabel#CodexMetricValue {
                color: #C1F56D;
                font-size: 14pt;
                font-weight: 800;
            }
            QLabel#CodexClusterLabel {
                color: #66D9CC;
                font-size: 8pt;
                font-weight: bold;
                padding-top: 4px;
            }
            QGroupBox {
                background-color: #1A1C1F;
                border: none;
                border-left: 4px solid #A6D854;
                border-radius: 6px;
                margin-top: 18px;
                padding: 16px 10px 10px 12px;
                font-weight: bold;
                color: #C1F56D;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 12px;
                padding: 0px 4px;
                color: #C1F56D;
                background-color: transparent;
            }
            QLabel {
                color: #C3C9B3;
            }
            QToolButton {
                background-color: #282A2D;
                color: #E2E2E6;
                border: none;
                border-radius: 6px;
                padding: 4px 6px;
                font-weight: bold;
            }
            QToolButton:hover {
                background-color: #333538;
            }
            QToolButton:pressed {
                background-color: #1E2023;
            }
            QToolButton[variant="primary"] {
                background-color: #A6D854;
                color: #223600;
            }
            QToolButton[variant="primary"]:hover {
                background-color: #C1F56D;
            }
            QLineEdit, QTextEdit, QListWidget, QComboBox {
                background-color: #0C0E11;
                color: #E2E2E6;
                border: none;
                border-radius: 6px;
                padding: 4px;
                selection-background-color: #1EA296;
                selection-color: #0C0E11;
            }
            QLineEdit:focus, QTextEdit:focus, QListWidget:focus, QComboBox:focus {
                border: 1px solid rgba(166, 216, 84, 102);
            }
            QListWidget::item {
                padding: 5px 6px;
                margin: 1px 0px;
            }
            QListWidget::item:selected {
                background-color: #282A2D;
                color: #C1F56D;
            }
            QSplitter::handle {
                background-color: #111316;
            }
            QSplitter::handle:hover {
                background-color: #1E2023;
            }
    """.replace("__FONT_STACK__", codex_font_stack)

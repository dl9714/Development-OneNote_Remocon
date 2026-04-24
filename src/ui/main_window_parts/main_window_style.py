# -*- coding: utf-8 -*-
from __future__ import annotations


def main_window_stylesheet(
    *,
    COLOR_BACKGROUND: str,
    COLOR_PRIMARY_TEXT: str,
    COLOR_SECONDARY_TEXT: str,
    COLOR_GROUPBOX_BG: str,
    COLOR_ACCENT: str,
    COLOR_SECONDARY_BUTTON: str,
    COLOR_SECONDARY_BUTTON_HOVER: str,
    COLOR_SECONDARY_BUTTON_PRESSED: str,
    COLOR_LIST_BG: str,
    COLOR_LIST_SELECTED: str,
    COLOR_STATUS_BAR: str,
    app_font_stack: str,
    base_font_pt: str,
    status_font_pt: str,
    side_label_font_pt: str,
) -> str:
    return f"""
            QWidget {{
                background-color: {COLOR_BACKGROUND};
                color: {COLOR_PRIMARY_TEXT};
                font-family: {app_font_stack};
                font-size: {base_font_pt};
            }}
            QGroupBox {{
                background-color: {COLOR_GROUPBOX_BG};
                border: 1px solid {COLOR_BACKGROUND};
                border-radius: 8px;
                margin-top: 10px;
                font-weight: bold;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 10px;
                left: 10px;
            }}
            QLabel {{
                color: {COLOR_SECONDARY_TEXT};
                font-weight: normal;
            }}
            QListWidget {{
                background-color: {COLOR_LIST_BG};
                border: 1px solid {COLOR_GROUPBOX_BG};
                border-radius: 4px;
            }}
            QListWidget::item:selected {{
                background-color: {COLOR_LIST_SELECTED};
                color: white;
            }}
            QTreeWidget {{
                background-color: {COLOR_LIST_BG};
                border: 1px solid {COLOR_GROUPBOX_BG};
                border-radius: 6px;
            }}
            QToolButton {{
                background-color: {COLOR_SECONDARY_BUTTON};
                color: {COLOR_PRIMARY_TEXT};
                border: none;
                border-radius: 4px;
                padding: 4px 6px;
            }}
            QToolButton:hover {{
                background-color: {COLOR_SECONDARY_BUTTON_HOVER};
            }}
            QToolButton:pressed {{
                background-color: {COLOR_SECONDARY_BUTTON_PRESSED};
            }}
            QPushButton {{
                background-color: {COLOR_SECONDARY_BUTTON};
                color: {COLOR_PRIMARY_TEXT};
                border: none;
                border-radius: 4px;
                padding: 6px 12px;
            }}
            QPushButton:hover {{
                background-color: {COLOR_SECONDARY_BUTTON_HOVER};
            }}
            QPushButton:pressed {{
                background-color: {COLOR_SECONDARY_BUTTON_PRESSED};
            }}
            QPushButton:disabled {{
                background-color: #404040;
                color: #808080;
            }}
            QMenuBar {{
                background-color: {COLOR_GROUPBOX_BG};
                color: {COLOR_PRIMARY_TEXT};
            }}
            QMenuBar::item:selected {{
                background-color: {COLOR_SECONDARY_BUTTON_HOVER};
            }}
            QMenu {{
                background-color: {COLOR_GROUPBOX_BG};
                border: 1px solid {COLOR_SECONDARY_BUTTON};
            }}
            QMenu::item:selected {{
                background-color: {COLOR_LIST_SELECTED};
            }}
            #StatusBarLabel {{
                background-color: {COLOR_STATUS_BAR};
                color: {COLOR_PRIMARY_TEXT};
                padding: 5px 12px;
                font-size: {status_font_pt};
                border-top: 1px solid #444444;
            }}
            QLineEdit {{
                background-color: {COLOR_LIST_BG};
                border: 1px solid {COLOR_SECONDARY_BUTTON};
                border-radius: 4px;
                padding: 4px 8px;
            }}
            QLineEdit:focus {{
                border: 1px solid {COLOR_LIST_SELECTED};
            }}
            QTextEdit {{
                background-color: {COLOR_LIST_BG};
                border: 1px solid {COLOR_GROUPBOX_BG};
                border-radius: 8px;
                color: {COLOR_PRIMARY_TEXT};
                padding: 10px;
            }}
            QTabWidget::pane {{
                border: 1px solid {COLOR_GROUPBOX_BG};
                border-radius: 6px;
            }}
            QTabBar::tab {{
                background-color: {COLOR_SECONDARY_BUTTON};
                color: {COLOR_PRIMARY_TEXT};
                padding: 7px 14px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                margin-right: 2px;
            }}
            QTabBar::tab:selected {{
                background-color: {COLOR_ACCENT};
                color: #111;
                font-weight: bold;
            }}
            QScrollArea {{
                border: none;
                background-color: transparent;
            }}
            QScrollArea > QWidget > QWidget {{
                background-color: {COLOR_BACKGROUND};
            }}
            #CodexSidePanel {{
                background-color: #242424;
                border-right: 1px solid #444444;
            }}
            #CodexSideSectionLabel {{
                color: {COLOR_SECONDARY_TEXT};
                font-size: {side_label_font_pt};
                padding: 8px 8px 2px 8px;
            }}
            #CodexSideNavButton {{
                background-color: transparent;
                color: {COLOR_PRIMARY_TEXT};
                border-radius: 16px;
                padding: 7px 12px;
                text-align: left;
            }}
            #CodexSideNavButton:hover {{
                background-color: {COLOR_SECONDARY_BUTTON_HOVER};
            }}
            #CodexSideNavButton:pressed {{
                background-color: {COLOR_SECONDARY_BUTTON_PRESSED};
            }}
        """

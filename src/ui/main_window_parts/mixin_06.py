# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


class MainWindowWindowsTemplatesAMixin:

    def _codex_onenote_templates_windows_a(
        self, values, title, body, target, section_group_id, section_id
    ) -> Dict[str, str]:
        from src.ui.main_window_parts.mixin_06_windows_a import (
            MainWindowWindowsTemplatesAMixin as _WindowsTemplatesAMixin,
        )

        return _WindowsTemplatesAMixin._codex_onenote_templates_windows_a(
            self, values, title, body, target, section_group_id, section_id
        )


class MainWindowWindowsTemplatesBMixin:

    def _codex_onenote_templates_windows_b(
        self, values, title, body, target, section_group_id, section_id
    ) -> Dict[str, str]:
        from src.ui.main_window_parts.mixin_06_windows_b import (
            MainWindowWindowsTemplatesBMixin as _WindowsTemplatesBMixin,
        )

        return _WindowsTemplatesBMixin._codex_onenote_templates_windows_b(
            self, values, title, body, target, section_group_id, section_id
        )


class MainWindowMixin06:

    def _codex_onenote_templates_windows(
        self, values: Optional[Dict[str, str]] = None
    ) -> Dict[str, str]:
        values = values or self._codex_codegen_values()
        title = self._ps_single_quoted(values["title"])
        body = self._ps_single_quoted(values["body"])
        target = self._ps_single_quoted(values["target"])
        section_group_id = self._ps_single_quoted(values["section_group_id"])
        section_id = self._ps_single_quoted(values["section_id"])
        templates: Dict[str, str] = {}
        templates.update(
            self._codex_onenote_templates_windows_a(
                values, title, body, target, section_group_id, section_id
            )
        )
        templates.update(
            self._codex_onenote_templates_windows_b(
                values, title, body, target, section_group_id, section_id
            )
        )
        return templates

_publish_context(globals())

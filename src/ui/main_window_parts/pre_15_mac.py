# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


class OpenAllNotebooksWorkerMacMixin:

    def _run_macos_open_all(self, result, win) -> None:
        state = self._collect_macos_open_all_sources(result, win)
        if state is None:
            return
        state = self._prepare_macos_open_all_records(result, state)
        if state is None:
            return
        self._launch_macos_open_all_records(result, win, state)

_publish_context(globals())

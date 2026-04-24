# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


class MainWindowInitStateMixin:

    def _restore_initial_splitter_states(self) -> None:
        # --- [START] 스플리터 상태 복원 로직 (수정됨) ---
        # 저장된 스플리터 상태 불러오기
        splitter_states = self.settings.get("splitter_states")
        restored = False
        codex_restored = False
        if isinstance(splitter_states, dict):
            try:
                main_state_b64 = splitter_states.get("main")
                if main_state_b64:
                    main_state = QByteArray.fromBase64(main_state_b64.encode("ascii"))
                    if not main_state.isEmpty():
                        self.main_splitter.restoreState(main_state)

                left_state_b64 = splitter_states.get("left")
                if left_state_b64:
                    left_state = QByteArray.fromBase64(left_state_b64.encode("ascii"))
                    if not left_state.isEmpty():
                        self.left_splitter.restoreState(left_state)

                codex_state_b64 = splitter_states.get("codex")
                if codex_state_b64 and getattr(self, "codex_splitter", None) is not None:
                    codex_state = QByteArray.fromBase64(codex_state_b64.encode("ascii"))
                    if not codex_state.isEmpty():
                        codex_restored = self.codex_splitter.restoreState(codex_state)

                restored = True
            except Exception as e:
                print(f"[WARN] 스플리터 상태 복원 실패: {e}")
                restored = False
                codex_restored = False

        # 복원에 실패했거나 저장된 상태가 없으면 기본값으로 설정
        if not restored:
            self.left_splitter.setSizes([150, 250])
            self.main_splitter.setSizes([400, 560])
        if getattr(self, "codex_splitter", None) is not None and not codex_restored:
            self.codex_splitter.setSizes([208, 920])
        # --- [END] 스플리터 상태 복원 로직 ---

_publish_context(globals())

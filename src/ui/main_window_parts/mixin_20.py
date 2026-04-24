# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin20:

    def _make_codex_checkable_item(
        self, label: str, value: str, checked: bool = False
    ) -> QListWidgetItem:
        item = QListWidgetItem(label)
        item.setData(Qt.ItemDataRole.UserRole, value)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        return item

    def _codex_selected_package_user_skills(self) -> List[str]:
        return self._codex_checked_list_values(
            getattr(self, "codex_skill_package_user_skill_list", None)
        )

    def _codex_skill_package_platform_list_widget(
        self, platform_key: str
    ) -> Optional[QListWidget]:
        if platform_key == CODEX_PLATFORM_MACOS:
            return getattr(self, "codex_skill_package_macos_skill_list", None)
        return getattr(self, "codex_skill_package_windows_skill_list", None)

    def _populate_codex_skill_package_user_skill_choices(
        self, checked_values: Optional[List[str]] = None
    ) -> None:
        skill_list = getattr(self, "codex_skill_package_user_skill_list", None)
        if skill_list is None:
            return

        checked: Set[str] = set()
        if checked_values is None:
            checked.update(self._codex_selected_package_user_skills())
        else:
            checked.update(
                str(value).strip() for value in checked_values if str(value).strip()
            )

        skill_list.blockSignals(True)
        try:
            skill_list.clear()
            seen: Set[str] = set()
            for skill in self._codex_skill_records_from_files():
                order_no = str(skill.get("order") or "").strip()
                name = str(skill.get("name") or "").strip()
                filename = str(skill.get("filename") or "").strip()
                value = order_no or name
                if not value:
                    continue
                label = f"[{order_no}] {name}" if order_no else name
                aliases = {
                    value,
                    order_no,
                    name,
                    filename[:-3] if filename.endswith(".md") else filename,
                }
                checked_aliases = {alias for alias in aliases if alias}
                item = self._make_codex_checkable_item(
                    label,
                    value,
                    bool(checked.intersection(checked_aliases)),
                )
                skill_list.addItem(item)
                seen.update(checked_aliases)

            for value in sorted(checked - seen):
                item = self._make_codex_checkable_item(value, value, True)
                skill_list.addItem(item)
        except Exception:
            pass
        finally:
            skill_list.blockSignals(False)

    def _set_codex_package_user_skills(self, skills: List[str]) -> None:
        self._populate_codex_skill_package_user_skill_choices(
            [str(item) for item in skills or [] if str(item).strip()]
        )

    def _codex_selected_package_codex_skills_for_platform(
        self, platform_key: str
    ) -> List[str]:
        return self._codex_checked_list_values(
            self._codex_skill_package_platform_list_widget(platform_key)
        )

    def _codex_selected_package_extra_skills(self) -> List[str]:
        return self._codex_text_lines_from_widget(
            getattr(self, "codex_skill_package_extra_skills_editor", None)
        )

    def _set_codex_package_platform_skills(
        self, platform_key: str, skills: List[str]
    ) -> None:
        selected = {
            _canonical_codex_platform_skill(platform_key, str(skill))
            for skill in skills or []
            if str(skill).strip()
        }
        skill_list = self._codex_skill_package_platform_list_widget(platform_key)
        if skill_list is None:
            return
        skill_list.blockSignals(True)
        for i in range(skill_list.count()):
            item = skill_list.item(i)
            value = _canonical_codex_platform_skill(
                platform_key,
                str(item.data(Qt.ItemDataRole.UserRole) or item.text()),
            )
            label = _canonical_codex_platform_skill(platform_key, item.text())
            item.setCheckState(
                Qt.CheckState.Checked
                if value in selected or label in selected
                else Qt.CheckState.Unchecked
            )
        skill_list.blockSignals(False)

    def _codex_skill_package_templates(self) -> Dict[str, Dict[str, Any]]:
        default_instructions = self._codex_skill_package_default_instructions()
        active_platform = _codex_active_platform_key()

        def package(
            name: str,
            description: str,
            user_skills: List[str],
            active_skills: List[str],
            extra_instructions: Optional[List[str]] = None,
        ) -> Dict[str, Any]:
            windows_skills = (
                list(active_skills) if active_platform == CODEX_PLATFORM_WINDOWS else []
            )
            macos_skills = (
                list(active_skills) if active_platform == CODEX_PLATFORM_MACOS else []
            )
            return {
                "version": 2,
                "name": name,
                "description": description,
                "user_skills": user_skills,
                "codex_skills_windows": windows_skills,
                "codex_skills_macos": macos_skills,
                "codex_skills_extra": [],
                "codex_skills": list(active_skills),
                "instructions": default_instructions + list(extra_instructions or []),
            }

        return {
            "quick_note": package(
                "기본 메모 패키지",
                "사용자 글쓰기 형태에 맞춰 OneNote 페이지를 빠르게 추가하는 기본 패키지입니다.",
                ["SK-001"],
                ["페이지 추가"],
            ),
            "work_log": package(
                "업무 기록 패키지",
                "업무 메모를 정리하고 필요한 경우 기존 페이지를 읽어 맥락을 이어가는 패키지입니다.",
                ["SK-001"],
                ["페이지 추가", "페이지 읽기", "페이지 본문 추가"],
                ["결과는 업무 기록 형식으로 정리", "다음 행동은 체크리스트로 분리"],
            ),
            "notebook_admin": package(
                "전자필기장 관리 패키지",
                "전자필기장 추가, 위치 확인, 열기/닫기 같은 관리 작업을 안전하게 실행하는 패키지입니다.",
                [],
                (
                    [
                        "전자필기장 추가",
                        "전자필기장 삭제",
                        "전자필기장 열기",
                        "전자필기장 닫기",
                        "계층 구조 조회",
                        "계층 동기화",
                        "특수 위치 조회",
                        "위치 조회",
                    ]
                    if active_platform == CODEX_PLATFORM_WINDOWS
                    else [
                        "전자필기장 추가",
                        "전자필기장 열기",
                        "전자필기장 닫기",
                        "계층 구조 조회",
                        "계층 동기화",
                        "특수 위치 조회",
                        "위치 조회",
                    ]
                ),
                ["작업 전후 계층 구조를 비교"],
            ),
            "meeting_note": package(
                "회의 정리 패키지",
                "회의 내용을 OneNote 페이지로 만들고 결정 사항과 후속 작업을 분리하는 패키지입니다.",
                ["SK-001"],
                ["페이지 추가", "페이지 읽기", "섹션 페이지 목록 읽기"],
                ["결정 사항과 할 일을 분리", "담당자와 기한이 있으면 본문에 유지"],
            ),
            "page_maintenance": package(
                "페이지 유지보수 패키지",
                "기존 페이지를 읽고 본문 추가, 제목 변경, 이동까지 이어서 처리하는 패키지입니다.",
                [],
                (
                    [
                        "페이지 XML 읽기",
                        "페이지 본문 추가",
                        "페이지 제목 변경",
                        "링크 생성",
                        "웹 링크 생성",
                        "부모 ID 조회",
                        "ID로 이동",
                    ]
                    if active_platform == CODEX_PLATFORM_WINDOWS
                    else [
                        "페이지 읽기",
                        "페이지 본문 추가",
                        "페이지 제목 변경",
                        "앱 링크 생성",
                        "웹 링크 생성",
                        "상위 위치 조회",
                        "URL로 이동",
                    ]
                ),
                (
                    ["대상 Page ID를 먼저 확인", "수정 전후 Page XML을 비교"]
                    if active_platform == CODEX_PLATFORM_WINDOWS
                    else ["현재 선택 페이지와 제목을 먼저 확인", "수정 후 화면에서 다시 검증"]
                ),
            ),
            "search_export": package(
                "검색과 내보내기 패키지",
                "OneNote에서 페이지를 찾고 링크를 만들거나 PDF로 내보내는 조회 중심 패키지입니다.",
                [],
                (
                    [
                        "페이지 검색",
                        "섹션 페이지 목록 읽기",
                        "링크 생성",
                        "웹 링크 생성",
                        "PDF 내보내기",
                        "MHTML 내보내기",
                    ]
                    if active_platform == CODEX_PLATFORM_WINDOWS
                    else [
                        "페이지 검색",
                        "섹션 페이지 목록 읽기",
                        "앱 링크 생성",
                        "웹 링크 생성",
                        "PDF 내보내기",
                    ]
                ),
                ["검색 결과는 최대 50개까지만 보고", "내보내기 파일 경로를 보고에 남김"],
            ),
            "sharing_export": package(
                "공유와 내보내기 패키지",
                "OneNote 항목의 앱 링크/웹 링크를 만들고 내보내기를 정리하는 패키지입니다.",
                [],
                (
                    ["링크 생성", "웹 링크 생성", "PDF 내보내기", "MHTML 내보내기", "XPS 내보내기"]
                    if active_platform == CODEX_PLATFORM_WINDOWS
                    else ["앱 링크 생성", "웹 링크 생성", "PDF 내보내기"]
                ),
                ["내보내기 대상과 저장 경로를 보고", "생성된 링크는 클립보드 복사 여부를 확인"],
            ),
        }

    def _apply_codex_skill_package_template(self) -> None:
        combo = getattr(self, "codex_skill_package_template_combo", None)
        key = str(combo.currentData() or "") if combo is not None else ""
        package = self._codex_skill_package_templates().get(key)
        if not package:
            return
        self._set_codex_skill_package_editor(dict(package))
        try:
            self.connection_status_label.setText(
                f"스킬 패키지 템플릿 적용: {package.get('name', '')}"
            )
        except Exception:
            pass

    def _default_codex_skill_package(self) -> Dict[str, Any]:
        active_platform = _codex_active_platform_key()
        return {
            "version": 2,
            "name": "기본 원노트 하네스",
            "description": (
                "현재 플랫폼 기준 OneNote 작업 흐름으로 새 페이지를 만들고 사용자 스킬 형식에 맞춰 기록할 때 쓰는 기본 패키지입니다."
            ),
            "user_skills": ["SK-001"],
            "codex_skills_windows": ["페이지 추가"]
            if active_platform == CODEX_PLATFORM_WINDOWS
            else [],
            "codex_skills_macos": ["페이지 추가"]
            if active_platform == CODEX_PLATFORM_MACOS
            else [],
            "codex_skills_extra": [],
            "codex_skills": ["페이지 추가"],
            "instructions": self._codex_skill_package_default_instructions(),
        }

    def _current_codex_skill_package(self) -> Dict[str, Any]:
        name_input = getattr(self, "codex_skill_package_name_input", None)
        desc_editor = getattr(self, "codex_skill_package_desc_editor", None)
        name = name_input.text().strip() if name_input is not None else ""
        desc = desc_editor.toPlainText().strip() if desc_editor is not None else ""
        windows_skills = self._codex_selected_package_codex_skills_for_platform(
            CODEX_PLATFORM_WINDOWS
        )
        macos_skills = self._codex_selected_package_codex_skills_for_platform(
            CODEX_PLATFORM_MACOS
        )
        extra_skills = self._codex_selected_package_extra_skills()
        combined: List[str] = []
        for skill in windows_skills + macos_skills + extra_skills:
            if skill and skill not in combined:
                combined.append(skill)
        return {
            "version": 2,
            "name": name or "새 스킬 패키지",
            "description": desc,
            "user_skills": self._codex_selected_package_user_skills(),
            "codex_skills_windows": windows_skills,
            "codex_skills_macos": macos_skills,
            "codex_skills_extra": extra_skills,
            "codex_skills": combined,
            "instructions": self._codex_text_lines_from_widget(
                getattr(self, "codex_skill_package_instructions_editor", None)
            ),
            "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        }

_publish_context(globals())

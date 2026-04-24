# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin21:

    def _set_codex_skill_package_editor(self, package: Dict[str, Any]) -> None:
        package = package or self._default_codex_skill_package()
        name_input = getattr(self, "codex_skill_package_name_input", None)
        if name_input is not None:
            name_input.setText(str(package.get("name") or "새 스킬 패키지"))
        desc_editor = getattr(self, "codex_skill_package_desc_editor", None)
        if desc_editor is not None:
            desc_editor.setPlainText(str(package.get("description") or ""))
        self._set_codex_package_user_skills(
            [str(item) for item in package.get("user_skills", [])]
        )
        legacy_skills = [str(item) for item in package.get("codex_skills", [])]
        windows_skills = package.get("codex_skills_windows")
        macos_skills = package.get("codex_skills_macos")
        extra_skills = package.get("codex_skills_extra")
        if windows_skills is None and macos_skills is None and extra_skills is None:
            windows_skills = legacy_skills
            macos_skills = []
            extra_skills = []
        self._set_codex_package_platform_skills(
            CODEX_PLATFORM_WINDOWS,
            [str(item) for item in windows_skills or []],
        )
        self._set_codex_package_platform_skills(
            CODEX_PLATFORM_MACOS,
            [str(item) for item in macos_skills or []],
        )
        self._set_codex_text_lines(
            getattr(self, "codex_skill_package_extra_skills_editor", None),
            [str(item) for item in extra_skills or []],
        )
        self._set_codex_text_lines(
            getattr(self, "codex_skill_package_instructions_editor", None),
            [str(item) for item in package.get("instructions", [])],
        )
        self._update_codex_skill_package_preview()

    def _codex_skill_package_prompt_text(
        self, package: Optional[Dict[str, Any]] = None
    ) -> str:
        package = package or self._current_codex_skill_package()
        name = str(package.get("name") or "새 스킬 패키지")
        description = str(package.get("description") or "").strip()
        user_skills = [str(item) for item in package.get("user_skills", []) if str(item).strip()]
        windows_skills = [
            str(item)
            for item in package.get("codex_skills_windows", [])
            if str(item).strip()
        ]
        macos_skills = [
            str(item)
            for item in package.get("codex_skills_macos", [])
            if str(item).strip()
        ]
        extra_skills = [
            str(item)
            for item in package.get("codex_skills_extra", [])
            if str(item).strip()
        ]
        instructions = [str(item) for item in package.get("instructions", []) if str(item).strip()]

        def bullet(items: List[str], fallback: str) -> str:
            return "\n".join(f"- {item}" for item in items) if items else f"- {fallback}"

        return f"""# 스킬 패키지: {name}

설명:
{description or "-"}

## 사용자 스킬

{bullet(user_skills, "사용자 스킬 미지정")}

## Windows 코덱스 스킬

{bullet(windows_skills, "Windows 스킬 미지정")}

## macOS 코덱스 스킬

{bullet(macos_skills, "macOS 스킬 미지정")}

## 추가/혼합 스킬

{bullet(extra_skills, "추가 스킬 미지정")}

## 코덱스 지침

{bullet(instructions, "코덱스 지침 미지정")}

## 사용 방식

이 스킬 패키지를 적용해서 현재 OneNote 작업 요청을 처리한다.
사용자 스킬은 결과물의 글쓰기 형태와 에이전트 역할에만 적용한다.
현재 실행 플랫폼은 `{_codex_platform_display_name(_codex_active_platform_key())}` 기준으로 우선 적용한다.
다른 플랫폼 스킬 섹션은 참고/혼합용으로 유지한다.
코덱스 스킬과 코덱스 지침은 OneNote 작업 실행 방식과 검증 기준으로 적용한다.
"""

    def _schedule_codex_skill_package_preview(self, *args) -> None:
        timer = getattr(self, "_codex_skill_package_preview_timer", None)
        if timer is None:
            timer = QTimer(self)
            timer.setSingleShot(True)
            timer.setInterval(120)
            timer.timeout.connect(self._update_codex_skill_package_preview)
            self._codex_skill_package_preview_timer = timer
        timer.start()

    def _update_codex_skill_package_preview(self) -> None:
        preview = getattr(self, "codex_skill_package_preview", None)
        if preview is None:
            return
        self._set_plain_text_if_changed(
            preview,
            self._codex_skill_package_prompt_text(),
        )

    def _selected_codex_skill_package_path(self) -> str:
        package_list = getattr(self, "codex_skill_package_list", None)
        item = package_list.currentItem() if package_list is not None else None
        return item.data(Qt.ItemDataRole.UserRole) if item is not None else ""

    def _refresh_codex_skill_package_list(self, selected_name: str = "") -> None:
        package_list = getattr(self, "codex_skill_package_list", None)
        if package_list is None:
            return
        current_name = selected_name or ""
        try:
            current_item = package_list.currentItem()
            if not current_name and current_item is not None:
                current_name = current_item.text()
        except Exception:
            pass

        package_list.blockSignals(True)
        package_list.clear()
        packages_dir = self._codex_skill_packages_dir()
        try:
            os.makedirs(packages_dir, exist_ok=True)
            rows: List[Dict[str, str]] = []
            for filename in sorted(os.listdir(packages_dir)):
                if not filename.lower().endswith(".json"):
                    continue
                path = os.path.join(packages_dir, filename)
                name = filename[:-5]
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    if isinstance(data, dict):
                        name = str(data.get("name") or name)
                except Exception:
                    pass
                rows.append({"name": name, "path": path})
            rows.sort(key=lambda row: _name_sort_key(row.get("name", "")))
            for row in rows:
                item = QListWidgetItem(row["name"])
                item.setData(Qt.ItemDataRole.UserRole, row["path"])
                package_list.addItem(item)
                if current_name and row["name"] == current_name:
                    package_list.setCurrentItem(item)
        finally:
            package_list.blockSignals(False)

    def _new_codex_skill_package(self) -> None:
        self._set_codex_skill_package_editor(self._default_codex_skill_package())
        try:
            self.connection_status_label.setText("새 스킬 패키지 초안을 만들었습니다.")
        except Exception:
            pass

    def _save_codex_skill_package(self) -> None:
        package = self._current_codex_skill_package()
        name = str(package.get("name") or "새 스킬 패키지").strip()
        packages_dir = self._codex_skill_packages_dir()
        path = os.path.join(packages_dir, self._codex_skill_slug(name) + ".json")
        try:
            self._write_json_file_atomic(path, package)
            self._refresh_codex_skill_package_list(name)
            try:
                self.connection_status_label.setText(f"스킬 패키지 저장 완료: {path}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "스킬 패키지 저장 실패", str(e))

    def _load_selected_codex_skill_package(
        self, item: Optional[QListWidgetItem] = None
    ) -> None:
        if item is not None and not isinstance(item, QListWidgetItem):
            item = None
        if item is None:
            package_list = getattr(self, "codex_skill_package_list", None)
            item = package_list.currentItem() if package_list is not None else None
        if item is None:
            QMessageBox.information(self, "안내", "먼저 스킬 패키지를 선택하세요.")
            return
        path = item.data(Qt.ItemDataRole.UserRole)
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                package = json.load(f)
            if not isinstance(package, dict):
                raise ValueError("스킬 패키지 JSON 형식이 올바르지 않습니다.")
            self._set_codex_skill_package_editor(package)
            try:
                self.connection_status_label.setText(f"스킬 패키지 불러옴: {path}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "스킬 패키지 불러오기 실패", str(e))

    def _delete_selected_codex_skill_package(self) -> None:
        path = self._selected_codex_skill_package_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 삭제할 스킬 패키지를 선택하세요.")
            return
        answer = QMessageBox.question(
            self,
            "스킬 패키지 삭제",
            f"선택한 스킬 패키지를 삭제합니다.\n\n{path}\n\n계속하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            os.remove(path)
            self._refresh_codex_skill_package_list()
            self._new_codex_skill_package()
            try:
                self.connection_status_label.setText("스킬 패키지를 삭제했습니다.")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "스킬 패키지 삭제 실패", str(e))

    def _copy_codex_skill_package_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_skill_package_prompt_text())
        try:
            self.connection_status_label.setText("스킬 패키지 호출문을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _open_codex_skill_packages_folder(self) -> None:
        try:
            os.makedirs(self._codex_skill_packages_dir(), exist_ok=True)
            open_path_in_system(self._codex_skill_packages_dir())
        except Exception as e:
            QMessageBox.warning(self, "스킬 패키지 폴더 열기 실패", str(e))

_publish_context(globals())

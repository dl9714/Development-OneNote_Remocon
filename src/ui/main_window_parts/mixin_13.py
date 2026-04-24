# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin13:

    def _select_recommended_codex_skill(self) -> None:
        order_no = self._recommended_codex_skill_order()
        if not getattr(self, "_codex_skill_records", None):
            self._refresh_codex_skill_list()
        target_key = self._codex_skill_search_key(order_no)
        records = getattr(self, "_codex_skill_records", [])
        if not any(
            self._codex_skill_search_key(record.get("order", "")) == target_key
            for record in records
        ):
            QMessageBox.information(
                self,
                "스킬 없음",
                f"현재 작업에 맞는 기본 스킬을 찾지 못했습니다: {order_no}",
            )
            return
        lookup = getattr(self, "codex_skill_order_lookup_input", None)
        if lookup is not None:
            lookup.setText(order_no)
        self._select_codex_skill_by_order_input()
        try:
            self.connection_status_label.setText(f"현재 작업에 맞는 스킬을 선택했습니다: {order_no}")
        except Exception:
            pass

    def _duplicate_selected_codex_skill(self) -> None:
        path = self._selected_codex_skill_path()
        if path and os.path.exists(path):
            try:
                meta = self._codex_skill_metadata_from_file(path, os.path.basename(path)[:-3])
            except Exception as e:
                QMessageBox.warning(self, "스킬 복제 실패", str(e))
                return
        else:
            meta = {
                "name": self.codex_skill_name_input.text().strip() or "새 코덱스 스킬",
                "trigger": self.codex_skill_trigger_input.text().strip(),
                "body": self.codex_skill_body_editor.toPlainText().strip(),
            }

        new_order = self._codex_next_skill_order()
        base_name = meta.get("name", "새 코덱스 스킬")
        new_name = f"{base_name} 복사본"
        new_body = meta.get("body", "")
        new_trigger = meta.get("trigger", "")
        skills_dir = self._codex_skills_dir()
        try:
            os.makedirs(skills_dir, exist_ok=True)
            path = os.path.join(skills_dir, self._codex_skill_slug(new_name) + ".md")
            if os.path.exists(path):
                path = os.path.join(skills_dir, self._codex_skill_slug(f"{new_name}-{new_order}") + ".md")
            self._write_text_file_atomic(
                path,
                self._codex_skill_markdown(new_name, new_trigger, new_body, new_order),
            )
            self._refresh_codex_skill_list()
            self.codex_skill_order_lookup_input.setText(new_order)
            self._select_codex_skill_by_order_input()
            try:
                self.connection_status_label.setText(f"스킬 복제 완료: {new_order}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "스킬 복제 실패", str(e))

    def _delete_selected_codex_skill(self) -> None:
        path = self._selected_codex_skill_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 삭제할 스킬을 선택하세요.")
            return
        filename = os.path.basename(path)
        if filename in ("README.md", "skill-order-index.md"):
            QMessageBox.warning(self, "삭제 불가", "기본 안내 파일은 삭제할 수 없습니다.")
            return
        answer = QMessageBox.question(
            self,
            "스킬 삭제",
            f"선택한 스킬 파일을 삭제합니다.\n\n{path}\n\n계속하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            os.remove(path)
            self._refresh_codex_skill_list()
            self._new_codex_skill_draft()
            try:
                self.connection_status_label.setText(f"스킬 삭제 완료: {filename}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "스킬 삭제 실패", str(e))

    def _new_codex_skill_from_clipboard(self) -> None:
        text = QApplication.clipboard().text().strip()
        if not text:
            QMessageBox.information(self, "안내", "클립보드에 스킬로 만들 내용이 없습니다.")
            return
        meta = self._codex_skill_metadata_from_text(text, "클립보드 스킬")
        order_no = meta.get("order") or self._codex_next_skill_order()
        name = meta.get("name") or "클립보드 스킬"
        if name.startswith("[") and "]" in name:
            name = name.split("]", 1)[1].strip() or "클립보드 스킬"
        self.codex_skill_order_input.setText(order_no)
        self.codex_skill_name_input.setText(name)
        self.codex_skill_trigger_input.setText(meta.get("trigger", "클립보드에서 만든 스킬"))
        self.codex_skill_body_editor.setPlainText(meta.get("body", text))
        self._update_codex_skill_call_preview()
        try:
            self.connection_status_label.setText("클립보드 내용으로 스킬 초안을 만들었습니다.")
        except Exception:
            pass

    def _codex_skill_markdown(
        self, name: str, trigger: str, body: str, order_no: str = ""
    ) -> str:
        name = (name or "새 코덱스 스킬").strip()
        trigger = (trigger or "").strip()
        body = (body or "").strip()
        order_no = (order_no or self._codex_next_skill_order()).strip()
        return f"""# {name}

## 주문번호

{order_no}

## Trigger

{trigger or "- 사용자가 이 스킬을 호출할 자연어 조건을 적는다."}

## Instructions

{body or "여기에 스킬 실행 절차, 입력 형식, 출력 기준, 검증 기준을 적는다."}
"""

    def _new_codex_skill_draft(self) -> None:
        self.codex_skill_order_input.setText(self._codex_next_skill_order())
        self.codex_skill_name_input.setText("새 코덱스 스킬")
        self.codex_skill_trigger_input.setText("이 스킬을 쓸 상황을 적는다")
        self.codex_skill_body_editor.setPlainText(
            """목표:
- 이 스킬이 해결할 작업을 명확히 적는다.

형식:
- 제목:
- 입력:
- 처리 방식:
- 출력:

검증:
- 완료 후 확인할 기준을 적는다.
"""
        )
        self._update_codex_skill_call_preview()

    def _create_default_codex_skill_set(self) -> None:
        skills_dir = self._codex_skills_dir()
        try:
            os.makedirs(skills_dir, exist_ok=True)
            existing_names: Set[str] = set()
            for filename in os.listdir(skills_dir):
                if not filename.lower().endswith(".md"):
                    continue
                if filename in ("README.md", "skill-order-index.md", "skill-audit.md"):
                    continue
                path = os.path.join(skills_dir, filename)
                try:
                    meta = self._codex_skill_metadata_from_file(path, filename[:-3])
                    existing_names.add(self._codex_skill_search_key(meta.get("name", "")))
                except Exception:
                    existing_names.add(self._codex_skill_search_key(filename[:-3]))

            created: List[str] = []
            for template in self._codex_skill_templates().values():
                name = template.get("name", "새 코덱스 스킬")
                if self._codex_skill_search_key(name) in existing_names:
                    continue
                order_no = self._codex_next_skill_order()
                path = os.path.join(skills_dir, self._codex_skill_slug(name) + ".md")
                if os.path.exists(path):
                    path = os.path.join(
                        skills_dir, self._codex_skill_slug(f"{name}-{order_no}") + ".md"
                    )
                self._write_text_file_atomic(
                    path,
                    self._codex_skill_markdown(
                        name,
                        template.get("trigger", ""),
                        template.get("body", ""),
                        order_no,
                    ),
                )
                created.append(f"{order_no} {name}")
                existing_names.add(self._codex_skill_search_key(name))

            self._refresh_codex_skill_list()
            self._update_codex_status_summary()
            if created:
                QMessageBox.information(
                    self,
                    "기본 스킬 세트 생성 완료",
                    "생성한 스킬:\n\n" + "\n".join(created),
                )
            else:
                QMessageBox.information(
                    self,
                    "기본 스킬 세트",
                    "이미 기본 스킬 양식이 모두 준비되어 있습니다.",
                )
        except Exception as e:
            QMessageBox.warning(self, "기본 스킬 세트 생성 실패", str(e))

    def _refresh_codex_skill_list(self) -> None:
        skill_list = getattr(self, "codex_skill_list", None)
        if skill_list is None:
            return
        skills_dir = self._codex_skills_dir()
        try:
            skill_index_rows = self._codex_skill_records_from_files()
            self._codex_skill_records = skill_index_rows
            try:
                self._codex_skill_records_dir_mtime = os.path.getmtime(skills_dir)
            except Exception:
                self._codex_skill_records_dir_mtime = None
            self._filter_codex_skill_list()
            self._populate_codex_skill_package_user_skill_choices()
            self._write_codex_skill_order_index(skill_index_rows)
        except Exception as e:
            try:
                self.connection_status_label.setText(f"코덱스 스킬 목록 로드 실패: {e}")
            except Exception:
                pass

    def _copy_codex_skill_order_index_to_clipboard(self) -> None:
        try:
            self._refresh_codex_skill_list()
            with open(self._codex_skill_order_index_path(), "r", encoding="utf-8") as f:
                text = f.read()
            QApplication.clipboard().setText(text)
            try:
                self.connection_status_label.setText("스킬 주문번호표를 클립보드에 복사했습니다.")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "주문번호표 복사 실패", str(e))

    def _save_codex_skill_draft(self) -> None:
        order_no = self.codex_skill_order_input.text().strip()
        name = self.codex_skill_name_input.text().strip()
        trigger = self.codex_skill_trigger_input.text().strip()
        body = self.codex_skill_body_editor.toPlainText().strip()
        if not name:
            name = "새 코덱스 스킬"
        if not order_no:
            order_no = self._codex_next_skill_order()
            self.codex_skill_order_input.setText(order_no)
        skills_dir = self._codex_skills_dir()
        try:
            path = os.path.join(skills_dir, self._codex_skill_slug(name) + ".md")
            self._write_text_file_atomic(
                path,
                self._codex_skill_markdown(name, trigger, body, order_no),
            )
            self._refresh_codex_skill_list()
            self._update_codex_skill_call_preview()
            try:
                self.connection_status_label.setText(
                    f"코덱스 스킬 저장 완료: {order_no} / {path}"
                )
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "스킬 저장 실패", str(e))

    def _load_selected_codex_skill(self, item: Optional[QListWidgetItem] = None) -> bool:
        if item is not None and not isinstance(item, QListWidgetItem):
            item = None
        if item is None:
            skill_list = getattr(self, "codex_skill_list", None)
            item = skill_list.currentItem() if skill_list is not None else None
        if item is None:
            return False
        path = item.data(Qt.ItemDataRole.UserRole)
        if not path:
            return False
        try:
            meta = self._codex_skill_metadata_from_file(path, item.text())
            self.codex_skill_order_input.setText(meta.get("order", ""))
            self.codex_skill_name_input.setText(meta.get("name", ""))
            self.codex_skill_trigger_input.setText(meta.get("trigger", ""))
            self.codex_skill_body_editor.setPlainText(meta.get("body", ""))
            self._update_codex_skill_call_preview()
            try:
                self.connection_status_label.setText(f"코덱스 스킬 불러옴: {path}")
            except Exception:
                pass
            return True
        except Exception as e:
            QMessageBox.warning(self, "스킬 불러오기 실패", str(e))
        return False

    def _open_selected_codex_skill_in_editor(self, item: Optional[QListWidgetItem] = None) -> None:
        if not self._load_selected_codex_skill(item):
            return
        try:
            self._scroll_codex_to_widget("codex_skill_editor_area")
        except Exception:
            pass

    def _copy_codex_skill_prompt_to_clipboard(self) -> None:
        text = self._current_codex_skill_markdown()
        QApplication.clipboard().setText(text)
        try:
            self.connection_status_label.setText("코덱스 스킬 초안을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _codex_targets_path(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "onenote-targets.json")

    def _default_codex_targets(self) -> List[Dict[str, str]]:
        return [
            {
                "name": "임시 메모 - 미정리",
                "path": "생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리",
                "notebook": "생산성도구-임시 메모",
                "section_group": "A 미정리-생성 메모",
                "section": "미정리",
                "section_group_id": "{2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}",
                "section_id": "{175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}",
            }
        ]

    def _load_codex_targets(self) -> List[Dict[str, str]]:
        path = self._codex_targets_path()
        try:
            if os.path.exists(path):
                file_mtime = os.path.getmtime(path)
                cached = getattr(self, "_codex_targets_cache", None)
                if cached and cached[0] == file_mtime:
                    return [dict(t) for t in cached[1]]

                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                targets = data.get("targets") if isinstance(data, dict) else data
                if isinstance(targets, list) and targets:
                    normalized = [dict(t) for t in targets if isinstance(t, dict)]
                    self._codex_targets_cache = (file_mtime, normalized)
                    return [dict(t) for t in normalized]
        except Exception as e:
            try:
                self.connection_status_label.setText(f"코덱스 작업 위치 로드 실패: {e}")
            except Exception:
                pass
        defaults = self._default_codex_targets()
        self._codex_targets_cache = (None, defaults)
        return [dict(t) for t in defaults]

_publish_context(globals())

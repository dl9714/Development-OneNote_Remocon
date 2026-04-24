# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin12:

    def _codex_current_skill_copy_text(self) -> str:
        order_no = ""
        skill_name = ""
        try:
            order_no = self.codex_skill_order_input.text().strip()
            skill_name = self.codex_skill_name_input.text().strip()
        except Exception:
            pass
        path = self._selected_codex_skill_path()
        return "\n".join(
            [
                f"주문번호: {order_no or '미지정'}",
                f"스킬명: {skill_name or '미지정'}",
                f"파일: {path or '미선택'}",
            ]
        )

    def _copy_codex_current_skill_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_current_skill_copy_text())
        try:
            self.connection_status_label.setText("현재 선택 스킬을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _codex_current_request_copy_text(self) -> str:
        try:
            return self._codex_visible_request_text()
        except Exception:
            return ""

    def _copy_codex_current_request_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_current_request_copy_text())
        try:
            self.connection_status_label.setText("현재 요청 요약을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _selected_codex_work_order_path(self) -> str:
        order_list = getattr(self, "codex_work_order_list", None)
        item = order_list.currentItem() if order_list is not None else None
        if item is None:
            return ""
        return item.data(Qt.ItemDataRole.UserRole) or ""

    def _selected_codex_work_order_text(self) -> str:
        path = self._selected_codex_work_order_path()
        if not path:
            return ""
        with open(path, "r", encoding="utf-8") as f:
            return f.read()

    def _schedule_refresh_codex_work_order_list(self, *args) -> None:
        timer = getattr(self, "_codex_work_order_list_timer", None)
        if timer is None:
            timer = QTimer(self)
            timer.setSingleShot(True)
            timer.setInterval(160)
            timer.timeout.connect(self._refresh_codex_work_order_list)
            self._codex_work_order_list_timer = timer
        timer.start()

    def _codex_work_order_search_text(self, path: str, filename: str) -> str:
        try:
            mtime = os.path.getmtime(path)
        except Exception:
            return filename

        cache = getattr(self, "_codex_work_order_search_text_cache", {})
        cached = cache.get(path)
        if cached and cached[0] == mtime:
            return cached[1]

        text = filename
        try:
            with open(path, "r", encoding="utf-8") as f:
                text += "\n" + f.read(8000)
        except Exception:
            pass

        cache[path] = (mtime, text)
        if len(cache) > 256:
            for old_path in list(cache)[: len(cache) - 256]:
                cache.pop(old_path, None)
        self._codex_work_order_search_text_cache = cache
        return text

    def _refresh_codex_work_order_list(self, selected_path: str = "") -> None:
        order_list = getattr(self, "codex_work_order_list", None)
        if order_list is None:
            return

        current_path = selected_path or self._selected_codex_work_order_path()
        query_input = getattr(self, "codex_work_order_search_input", None)
        query = query_input.text().strip() if query_input is not None else ""
        query_key = self._codex_skill_search_key(query)
        order_list.blockSignals(True)
        order_list.clear()
        try:
            os.makedirs(self._codex_requests_dir(), exist_ok=True)
            records = []
            for filename in os.listdir(self._codex_requests_dir()):
                if not filename.lower().endswith(".md"):
                    continue
                path = os.path.join(self._codex_requests_dir(), filename)
                try:
                    mtime = os.path.getmtime(path)
                except Exception:
                    mtime = 0.0
                if query_key:
                    haystack = self._codex_work_order_search_text(path, filename)
                    if query_key not in self._codex_skill_search_key(haystack):
                        continue
                records.append((mtime, filename, path))
            records.sort(key=lambda row: (-row[0], row[1]))
            for mtime, filename, path in records:
                stamp = time.strftime("%Y-%m-%d %H:%M", time.localtime(mtime)) if mtime else ""
                label = f"{stamp}  {filename}" if stamp else filename
                item = QListWidgetItem(label)
                item.setData(Qt.ItemDataRole.UserRole, path)
                order_list.addItem(item)
        except Exception as e:
            try:
                self.connection_status_label.setText(f"작업 주문서 기록 로드 실패: {e}")
            except Exception:
                pass
        finally:
            order_list.blockSignals(False)

        if current_path:
            for i in range(order_list.count()):
                item = order_list.item(i)
                if item.data(Qt.ItemDataRole.UserRole) == current_path:
                    order_list.setCurrentRow(i)
                    self._on_codex_work_order_selected(item)
                    return
        if order_list.count() > 0:
            order_list.setCurrentRow(0)
            self._on_codex_work_order_selected(order_list.item(0))
        else:
            preview = getattr(self, "codex_work_order_history_preview", None)
            if preview is not None:
                preview.setPlainText("저장된 작업 주문서가 없습니다.")

    def _on_codex_work_order_selected(
        self, item: Optional[QListWidgetItem] = None
    ) -> None:
        preview = getattr(self, "codex_work_order_history_preview", None)
        if preview is None:
            return
        if item is None:
            order_list = getattr(self, "codex_work_order_list", None)
            item = order_list.currentItem() if order_list is not None else None
        path = item.data(Qt.ItemDataRole.UserRole) if item is not None else ""
        if not path:
            preview.setPlainText("선택된 작업 주문서가 없습니다.")
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                preview.setPlainText(f.read())
        except Exception as e:
            preview.setPlainText(f"작업 주문서를 읽지 못했습니다.\n\n{e}")

    def _copy_selected_codex_work_order_to_clipboard(self) -> None:
        path = self._selected_codex_work_order_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 작업 주문서를 선택하세요.")
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                QApplication.clipboard().setText(f.read())
            try:
                self.connection_status_label.setText("선택 작업 주문서를 클립보드에 복사했습니다.")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "작업 주문서 복사 실패", str(e))

    def _open_selected_codex_work_order_file(self) -> None:
        path = self._selected_codex_work_order_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 작업 주문서를 선택하세요.")
            return
        try:
            open_path_in_system(path)
        except Exception as e:
            QMessageBox.warning(self, "작업 주문서 열기 실패", str(e))

    def _delete_selected_codex_work_order(self) -> None:
        path = self._selected_codex_work_order_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 삭제할 작업 주문서를 선택하세요.")
            return
        answer = QMessageBox.question(
            self,
            "작업 주문서 삭제",
            f"선택한 작업 주문서를 삭제합니다.\n\n{path}\n\n계속하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            os.remove(path)
            self._refresh_codex_work_order_list()
            try:
                self.connection_status_label.setText("작업 주문서를 삭제했습니다.")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "작업 주문서 삭제 실패", str(e))

    def _codex_extract_fenced_section(self, text: str, heading: str) -> str:
        pattern = rf"(?ms)^##\s+{re.escape(heading)}\s*$\n+```[^\n]*\n(.*?)\n```\s*"
        match = re.search(pattern, text or "")
        if match:
            return match.group(1).strip()
        return self._codex_skill_section(text, heading)

    def _load_selected_codex_work_order_into_request(self) -> None:
        try:
            text = self._selected_codex_work_order_text()
        except Exception as e:
            QMessageBox.warning(self, "주문서 불러오기 실패", str(e))
            return
        if not text:
            QMessageBox.information(self, "안내", "먼저 작업 주문서를 선택하세요.")
            return

        request_text = self._codex_extract_fenced_section(text, "요청문")
        target_path = ""
        skill_order = ""
        title = ""

        target_match = re.search(r"(?m)^-\s*작업 경로:\s*(.+?)\s*$", text)
        if target_match:
            target_path = target_match.group(1).strip()
        order_match = re.search(r"(?m)^-\s*주문번호:\s*(.+?)\s*$", text)
        if order_match:
            skill_order = order_match.group(1).strip()
        title_match = re.search(r"(?m)^제목/이름:\s*\n(.+?)\s*(?:\n\n|\Z)", request_text)
        if title_match:
            title = title_match.group(1).strip()

        target_input = getattr(self, "codex_request_target_input", None)
        if target_input is not None and target_path:
            target_input.setText(target_path)

        title_input = getattr(self, "codex_request_title_input", None)
        if title_input is not None:
            title_input.setText(title or os.path.basename(self._selected_codex_work_order_path()))

        body_editor = getattr(self, "codex_request_body_editor", None)
        if body_editor is not None:
            body_editor.setPlainText(request_text or text)

        if skill_order:
            lookup = getattr(self, "codex_skill_order_lookup_input", None)
            if lookup is not None:
                lookup.setText(skill_order)
                self._select_codex_skill_by_order_input()

        self._update_codex_codegen_previews()
        try:
            self._scroll_codex_to_widget("codex_request_group_widget")
        except Exception:
            pass
        try:
            self.connection_status_label.setText("선택 주문서를 현재 요청으로 불러왔습니다.")
        except Exception:
            pass

    def _copy_selected_codex_work_order_followup_prompt(self) -> None:
        try:
            text = self._selected_codex_work_order_text()
        except Exception as e:
            QMessageBox.warning(self, "후속 요청 생성 실패", str(e))
            return
        if not text:
            QMessageBox.information(self, "안내", "먼저 작업 주문서를 선택하세요.")
            return

        prompt = f"""아래 이전 OneNote 작업 주문서를 이어서 처리해라.

먼저 기존 주문서의 대상 위치, 스킬, 요청문을 요약하고, 이어서 필요한 다음 행동만 제안해라.
실행이 필요하면 코덱스 전용 OneNote 조작 지침을 따르고, 완료 후 변경 항목과 검증 결과만 간단히 보고해라.

## 이전 작업 주문서

````markdown
{text}
````
"""
        QApplication.clipboard().setText(prompt)
        try:
            self.connection_status_label.setText("선택 주문서 후속 요청을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _copy_selected_codex_skill_path_to_clipboard(self) -> None:
        path = self._selected_codex_skill_path()
        if not path:
            QMessageBox.information(self, "안내", "먼저 스킬을 선택하세요.")
            return
        QApplication.clipboard().setText(path)
        try:
            self.connection_status_label.setText("선택 스킬 파일 경로를 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _select_codex_skill_by_order_input(self) -> None:
        lookup = getattr(self, "codex_skill_order_lookup_input", None)
        order_no = lookup.text().strip() if lookup is not None else ""
        if not order_no:
            order_no = self.codex_skill_order_input.text().strip()
        if not order_no:
            return

        target_key = self._codex_skill_search_key(order_no)
        records = getattr(self, "_codex_skill_records", [])
        match = None
        for record in records:
            if self._codex_skill_search_key(record.get("order", "")) == target_key:
                match = record
                break
        if match is None:
            QMessageBox.information(self, "스킬 없음", f"주문번호를 찾지 못했습니다: {order_no}")
            return

        self.codex_skill_search_input.blockSignals(True)
        self.codex_skill_search_input.setText("")
        self.codex_skill_search_input.blockSignals(False)
        self._populate_codex_skill_list(records)

        path = match.get("path", "")
        for i in range(self.codex_skill_list.count()):
            item = self.codex_skill_list.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == path:
                self.codex_skill_list.setCurrentRow(i)
                self._load_selected_codex_skill(item)
                return

    def _recommended_codex_skill_order(self) -> str:
        ranked = self._rank_codex_skill_records()
        if ranked:
            order_no = ranked[0].get("record", {}).get("order", "")
            if order_no:
                return order_no
        return "SK-001"

_publish_context(globals())

# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin09:

    def _codex_skill_templates(self) -> Dict[str, Dict[str, str]]:
        return {
            "writing": {
                "name": "나의 기본 글쓰기 형식",
                "trigger": "사용자가 정리된 글쓰기 형식으로 OneNote 페이지 작성을 요청할 때",
                "body": """목표:
- 사용자의 거친 메모를 읽기 쉬운 글로 정리한다.

형식:
- 제목: 핵심 주제를 짧게 쓴다.
- 한 줄 요약: 결론을 먼저 쓴다.
- 본문: 배경, 핵심 내용, 다음 행동 순서로 정리한다.
- 체크포인트: 다시 확인할 항목을 bullet로 남긴다.

작성 기준:
- 사용자의 원래 표현을 최대한 보존하되 문장 흐름만 정리한다.
- 불확실한 내용은 단정하지 말고 확인 필요로 표시한다.
- OneNote에는 제목과 본문을 분리해서 작성한다.

검증:
- 제목과 본문이 위 형식을 충족하는지 확인한다.
""",
            },
            "daily_log": {
                "name": "업무일지 정리",
                "trigger": "오늘 작업 내용, 업무 기록, 진행 상황을 OneNote에 남길 때",
                "body": """목표:
- 하루 동안 한 작업과 다음 행동을 빠르게 되돌아볼 수 있게 정리한다.

형식:
- 날짜:
- 오늘 한 일:
- 결정한 것:
- 막힌 것:
- 다음 행동:

작성 기준:
- 작업 결과와 미완료 항목을 분리한다.
- 다음 행동은 바로 실행 가능한 동사형으로 쓴다.
- 중요한 파일/프로젝트/스킬 주문번호가 있으면 함께 적는다.

검증:
- 페이지 제목에 날짜 또는 작업명이 들어갔는지 확인한다.
""",
            },
            "meeting": {
                "name": "회의록 정리",
                "trigger": "회의 내용, 통화 내용, 논의 결과를 OneNote에 정리할 때",
                "body": """목표:
- 회의에서 결정된 내용과 후속 작업을 빠르게 찾을 수 있게 정리한다.

형식:
- 회의명:
- 참석자:
- 논의 요약:
- 결정 사항:
- 액션 아이템:
- 보류/리스크:

작성 기준:
- 결정 사항과 의견을 섞지 않는다.
- 액션 아이템에는 담당자, 마감, 확인 방법을 포함한다.
- 근거가 부족한 내용은 보류/리스크로 분리한다.

검증:
- 액션 아이템이 별도 목록으로 남아 있는지 확인한다.
""",
            },
            "idea": {
                "name": "아이디어 정리",
                "trigger": "아이디어, 기획, 개선안을 OneNote에서 발전시킬 때",
                "body": """목표:
- 떠오른 아이디어를 실행 가능한 형태로 바꾼다.

형식:
- 아이디어:
- 왜 필요한가:
- 사용자/상황:
- 가능한 구현:
- 예상 문제:
- 다음 실험:

작성 기준:
- 막연한 표현을 구체적인 실험이나 작업 단위로 바꾼다.
- 구현 아이디어와 검증 아이디어를 분리한다.
- 당장 할 수 없는 것은 후보로만 남긴다.

검증:
- 다음 실험이 1개 이상 있는지 확인한다.
""",
            },
            "study": {
                "name": "학습 노트 정리",
                "trigger": "강의, 문서, 공부 내용을 OneNote 학습 노트로 만들 때",
                "body": """목표:
- 학습 내용을 다시 보기 쉬운 구조로 압축한다.

형식:
- 주제:
- 핵심 개념:
- 예시:
- 헷갈린 점:
- 적용할 곳:
- 복습 질문:

작성 기준:
- 정의, 예시, 적용을 분리한다.
- 이해가 불확실한 부분은 헷갈린 점에 남긴다.
- 복습 질문은 나중에 바로 테스트할 수 있게 작성한다.

검증:
- 핵심 개념과 복습 질문이 둘 다 있는지 확인한다.
""",
            },
            "checklist": {
                "name": "작업 체크리스트",
                "trigger": "반복 작업, 배포 전 점검, 정리 절차를 체크리스트로 만들 때",
                "body": """목표:
- 반복 작업을 빠뜨리지 않도록 체크리스트로 만든다.

형식:
- 목적:
- 사전 확인:
- 실행 순서:
- 검증:
- 실패 시 복구:

작성 기준:
- 각 항목은 체크 가능한 문장으로 쓴다.
- 실행 순서와 검증 순서를 분리한다.
- 위험한 작업은 복구 방법을 같이 남긴다.

검증:
- 실행 전/후 확인 항목이 모두 있는지 확인한다.
""",
            },
            "onenote_cleanup": {
                "name": "OneNote 정리 작업",
                "trigger": "OneNote의 메모, 섹션, 섹션 그룹을 정리하거나 재배치할 때",
                "body": """목표:
- OneNote 구조를 안전하게 정리하고 변경 내역을 검증한다.

형식:
- 현재 위치:
- 바꿀 위치:
- 변경 작업:
- 보존할 항목:
- 검증 방법:

작업 기준:
- 삭제보다 이동/이름 변경을 우선한다.
- 전자필기장/섹션/페이지 계층을 먼저 조회한 뒤 작업한다.
- 보존할 항목과 바꿀 위치를 작업 전에 분리해서 적는다.

검증:
- 작업 전후 계층 구조를 비교한다.
""",
            },
        }

    def _apply_codex_skill_template(self) -> None:
        combo = getattr(self, "codex_skill_template_combo", None)
        if combo is None:
            return
        key = combo.currentData()
        template = self._codex_skill_templates().get(key)
        if not template:
            return

        self.codex_skill_order_input.setText(self._codex_next_skill_order())
        self.codex_skill_name_input.setText(template.get("name", "새 코덱스 스킬"))
        self.codex_skill_trigger_input.setText(template.get("trigger", ""))
        self.codex_skill_body_editor.setPlainText(template.get("body", ""))
        self._update_codex_skill_call_preview()
        try:
            self.connection_status_label.setText(
                f"스킬 양식 적용: {template.get('name', '')}"
            )
        except Exception:
            pass

    def _codex_skill_section(self, text: str, heading: str) -> str:
        pattern = rf"(?ms)^##\s+{re.escape(heading)}\s*$\n(.*?)(?=^##\s+|\Z)"
        match = re.search(pattern, text or "")
        return match.group(1).strip() if match else ""

    def _codex_skill_metadata_from_text(
        self, text: str, fallback_name: str = ""
    ) -> Dict[str, str]:
        lines = (text or "").splitlines()
        name = fallback_name
        for line in lines:
            if line.strip().startswith("# "):
                name = line.strip()[2:].strip()
                break

        order_no = self._codex_skill_section(text, "주문번호")
        if not order_no:
            match = re.search(r"(?im)^\s*(?:주문번호|Order)\s*[:：]\s*(.+?)\s*$", text or "")
            order_no = match.group(1).strip() if match else ""

        trigger = self._codex_skill_section(text, "Trigger")
        body = self._codex_skill_section(text, "Instructions")
        if not body:
            body = text or ""

        return {
            "name": name or "새 코덱스 스킬",
            "order": order_no,
            "trigger": trigger,
            "body": body,
        }

    def _codex_skill_metadata_from_file(
        self, path: str, fallback_name: str = ""
    ) -> Dict[str, str]:
        stat = os.stat(path)
        signature = (stat.st_mtime_ns, stat.st_size)
        cache = getattr(self, "_codex_skill_metadata_cache", {})
        cached = cache.get(path)
        if cached and cached[0] == signature:
            return dict(cached[1])

        with open(path, "r", encoding="utf-8") as f:
            meta = self._codex_skill_metadata_from_text(f.read(), fallback_name)

        cache[path] = (signature, dict(meta))
        if len(cache) > 256:
            for old_path in list(cache)[: len(cache) - 256]:
                cache.pop(old_path, None)
        self._codex_skill_metadata_cache = cache
        return dict(meta)

    def _codex_next_skill_order(self) -> str:
        used: Set[int] = set()
        skills_dir = self._codex_skills_dir()
        try:
            records = getattr(self, "_codex_skill_records", [])
            records_mtime = getattr(self, "_codex_skill_records_dir_mtime", None)
            if records and records_mtime == os.path.getmtime(skills_dir):
                for record in records:
                    match = re.search(r"(\d+)", record.get("order", ""))
                    if match:
                        used.add(int(match.group(1)))
                n = 1
                while n in used:
                    n += 1
                return f"SK-{n:03d}"
        except Exception:
            used.clear()

        try:
            for filename in os.listdir(skills_dir):
                if not filename.lower().endswith(".md"):
                    continue
                if filename in ("README.md", "skill-order-index.md", "skill-audit.md"):
                    continue
                path = os.path.join(skills_dir, filename)
                meta = self._codex_skill_metadata_from_file(path, filename[:-3])
                match = re.search(r"(\d+)", meta.get("order", ""))
                if match:
                    used.add(int(match.group(1)))
        except Exception:
            pass

        n = 1
        while n in used:
            n += 1
        return f"SK-{n:03d}"

    def _write_codex_skill_order_index(self, skills: List[Dict[str, str]]) -> None:
        path = self._codex_skill_order_index_path()
        rows = []
        for skill in sorted(skills, key=lambda s: (s.get("order") or "ZZZ", s.get("name") or "")):
            order_no = skill.get("order") or "미지정"
            name = skill.get("name") or "이름 없음"
            filename = skill.get("filename") or ""
            trigger = (skill.get("trigger") or "").replace("\n", " ").strip()
            rows.append(f"| {order_no} | {name} | `{filename}` | {trigger} |")

        text = "\n".join(
            [
                "# 사용자 스킬 주문번호표",
                "",
                "사용자가 주문번호로 스킬을 지시하면 이 표에서 해당 Markdown 파일을 찾아 따른다. OneNote 조작 방식은 `docs/codex/instructions`의 코덱스 전용 지침에서 관리한다.",
                "",
                "| 주문번호 | 스킬 이름 | 파일 | 호출 조건 |",
                "| --- | --- | --- | --- |",
                *rows,
                "",
            ]
        )
        self._write_text_file_atomic(path, text)

    def _codex_skill_search_key(self, *parts: Any) -> str:
        text = " ".join("" if p is None else str(p) for p in parts)
        return re.sub(r"\s+", "", unicodedata.normalize("NFKC", text)).casefold()

    def _codex_skill_category_names(self) -> List[str]:
        return [
            "OneNote/도구",
            "기록/정리",
            "기획/전략",
            "분석/리서치",
            "운영/품질",
            "소통/협업",
            "기타",
        ]

_publish_context(globals())

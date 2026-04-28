# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin08:

    def _copy_codex_template_to_clipboard(self) -> None:
        text = self._selected_codex_template_text()
        if not text:
            return
        QApplication.clipboard().setText(text)
        try:
            self.connection_status_label.setText("코덱스 OneNote 작업 양식을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _codex_skills_dir(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "skills")

    def _codex_skill_packages_dir(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "skill-packages")

    def _codex_instructions_dir(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "instructions")

    def _codex_internal_instructions_legacy_path(self) -> str:
        return os.path.join(self._codex_instructions_dir(), "onenote-com-internal.md")

    def _codex_internal_instructions_path(
        self, platform_key: Optional[str] = None
    ) -> str:
        platform_key = platform_key or _codex_active_platform_key()
        filename = (
            "onenote-macos-internal.md"
            if platform_key == CODEX_PLATFORM_MACOS
            else "onenote-windows-internal.md"
        )
        return os.path.join(self._codex_instructions_dir(), filename)

    def _codex_internal_reference_text(self) -> str:
        return (
            "OneNote 조작 방식, 대상 판정, 플랫폼별 스크립트/접근성 패턴, 안전 실행 순서, 검증 기준은 "
            "`docs/codex/instructions/`와 `docs/codex/onenote-targets.json`에서 "
            "필요할 때 조회한다. 사용자 요청문이나 작업 주문서에는 내부 지침 전문을 붙이지 않는다."
        )

    def _codex_requests_dir(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "requests")

    def _codex_request_draft_path(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "request-draft.json")

    def _codex_skill_slug(self, name: str) -> str:
        raw = (name or "").strip()
        raw = re.sub(r'[<>:"/\\\\|?*]+', "-", raw)
        raw = re.sub(r"\s+", "-", raw)
        raw = raw.strip(".-")
        return raw or "codex-skill"

    def _codex_skill_order_index_path(self) -> str:
        return os.path.join(self._codex_skills_dir(), "skill-order-index.md")

    def _write_text_file_atomic(self, path: str, text: str) -> bool:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        try:
            with open(path, "r", encoding="utf-8") as f:
                if f.read() == text:
                    return False
        except FileNotFoundError:
            pass
        except Exception:
            pass

        tmp_path = f"{path}.{uuid.uuid4().hex}.tmp"
        try:
            with open(tmp_path, "w", encoding="utf-8", newline="") as f:
                f.write(text)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp_path, path)
        finally:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass
        return True

    def _write_json_file_atomic(self, path: str, payload: Any) -> bool:
        text = json.dumps(payload, ensure_ascii=False, indent=2)
        return self._write_text_file_atomic(path, text)

    def _codex_builtin_internal_instructions_text_windows(self) -> str:
        return """# 코덱스 전용 OneNote 조작 지침

이 문서는 사용자 스킬이 아니다. OneNote 작업을 수행하는 Codex가 항상 전제로 삼는 내부 실행 지침이다.

## 적용 순서

1. 사용자 요청에서 목표, 대상, 출력 형식, 금지 조건, 주의 조건을 분리한다.
2. 사용자 스킬은 글쓰기 형식과 정리 방식만 적용하고, OneNote 조작 방식은 이 폴더의 내부 지침에서 고른다.
3. 대상은 명시된 ID, 저장된 대상 ID, 위치 캐시, 제한된 계층 조회 순서로 확정한다.
4. 변경 전에는 작업 종류, 대상 경로, 대상 ID, 검증 방법을 먼저 확정한다.
5. 완료 후에는 플랫폼에 맞는 조회 결과로 검증한다. Windows는 COM 결과, macOS는 접근성/UI 조회 결과를 우선 사용한다.
6. 최종 보고에는 변경 항목, 대상, 검증 결과, 남은 확인 사항만 짧게 남긴다.

## 사용자 스킬과의 경계

- 사용자 스킬에서는 `## Instructions`만 작업에 맞게 적용한다.
- 사용자 요청문, 작업 주문서, 스킬 호출문에는 이 문서 전문이나 플랫폼 내부 템플릿을 붙이지 않는다.
- 글쓰기 형식, 정리 방식, 이름 규칙처럼 요청마다 달라지는 기준은 사용자 스킬에 둔다.
- 플랫폼별 호출 순서, 대상 ID/경로 우선순위, 본문 수정 방식, 검증 절차는 코덱스 전용 지침에 둔다.
- 사용자가 내부 구현 설명을 요청하지 않았다면 COM/접근성 세부 호출을 길게 설명하지 않는다.

## 대상 판정 원칙

- 사용자가 대상 ID를 직접 준 경우 그 ID를 최우선으로 사용한다.
- 사용자가 전자필기장, 섹션 그룹, 섹션 이름을 준 경우 상위 경로까지 함께 맞는지 확인한다.
- `docs/codex/onenote-targets.json`의 고정 대상이 있으면 위치 캐시보다 먼저 확인한다.
- `docs/codex/onenote-location-cache.json`의 ID는 빠른 경로로만 사용하고, 실패하면 계층 조회로 재확인한다.
- 이름만으로 대상을 확정하지 않는다. 같은 이름이 여러 개면 사용자에게 확인하거나 더 좁은 상위 경로를 사용한다.
- 새 ID를 찾거나 자주 쓸 대상이 생기면 검증 후 대상 캐시에 저장한다.

## 안전 실행 원칙

- Windows에서는 화면 클릭 자동화보다 OneNote COM API를 우선 사용한다.
- macOS에서는 OneNote 접근성 트리와 UI 자동화를 우선 사용한다.
- Windows PowerShell에서는 `New-Object -ComObject OneNote.Application`으로 연결한다.
- 구조 탐색은 기본적으로 `GetHierarchy('', hsSections, ref xml)`까지만 사용한다.
- `hsPages` 전체 조회는 페이지가 많으면 느리므로 페이지 복제, 페이지 목록 조회, 최종 페이지 수 검증처럼 필요한 경우에만 쓴다.
- XML은 문자열 치환보다 XML 파서로 수정한다.
- `OpenHierarchy` 또는 `CreateNewPage` 직후에는 짧게 대기하고 재조회해서 새 ID가 실제 사용 가능한지 확인한다.
- 삭제, 덮어쓰기, 대량 이동, 대량 복제는 대상과 영향 범위를 다시 확인한 뒤 실행한다.
- OneNote 호출이 실패하면 같은 쓰기 작업을 무한 반복하지 않는다. 원인 확인 후 한 번만 보정 재시도한다.

## 작업별 내부 문서

- 페이지 추가: `원노트-페이지-추가.md`
- 섹션 생성: `원노트-섹션-생성.md`
- 섹션 그룹 생성: `원노트-섹션그룹-생성.md`
- 전자필기장 생성/열기: `원노트-전자필기장-생성.md`
- 전자필기장 복제: `원노트-전자필기장-복제.md`
- 대상 ID 찾기: `원노트-대상ID-찾기.md`
- 작업 템플릿 기준: `onenote-com-templates.md`

## 빠른 라우팅

- 페이지 추가는 대상 `Section ID`로 `CreateNewPage`, `GetPageContent`, `UpdatePageContent`를 사용한다.
- 섹션 생성은 대상 `SectionGroup ID`로 `OpenHierarchy("섹션명.one", sectionGroupId, ref newId, cftSection)`를 사용한다.
- 섹션 그룹 생성은 전자필기장 또는 섹션 그룹 ID로 `OpenHierarchy("그룹명", parentId, ref newId, cftFolder)`를 사용한다.
- 새 전자필기장은 OneNote가 열 수 있는 로컬/동기화 경로를 확인한 뒤 `OpenHierarchy(notebookPath, "", ref newId, cftNotebook)`를 사용한다.
- 전자필기장 복제는 대상 이름을 `코덱스-{원본 전자필기장명}`으로 잡고, 활성 섹션 그룹/섹션/페이지 수가 원본과 같은지 확인한다.

## 검증 기준

- 페이지 추가/수정: `GetPageContent(pageId)`에서 제목과 본문 일부를 확인한다.
- 섹션 생성: `GetHierarchy(sectionGroupId, hsSections, ref xml)`에서 새 섹션 이름을 확인한다.
- 섹션 그룹 생성: `GetHierarchy(parentId, hsSections, ref xml)`에서 새 그룹 이름을 확인한다.
- 전자필기장 생성/열기: `GetHierarchy('', hsNotebooks, ref xml)` 또는 `hsSections`에서 열린 전자필기장을 확인한다.
- 전자필기장 복제: 내부 휴지통을 제외한 활성 섹션 그룹 수, 섹션 수, 페이지 수가 원본과 대상에서 일치하는지 확인한다.
- 캐시 갱신: 저장한 ID로 다시 조회했을 때 같은 이름과 상위 경로가 나오는지 확인한다.

## 실패 대응

- 실패하면 먼저 대상 ID 만료, 잘못된 상위 경로, OneNote 동기화 지연, 권한 문제를 구분한다.
- ID 호출 실패 시 전체 탐색으로 바로 확장하지 말고 상위 대상부터 제한적으로 재조회한다.
- XML 수정 실패 시 원본 XML 일부와 namespace 처리 여부를 확인한다.
- 쓰기 성공 여부가 애매하면 추가 쓰기를 하지 말고 조회 검증부터 수행한다.

## 보고 기준

- 성공하면 만든/수정한 OneNote 항목, 대상 경로 또는 ID, 검증 결과만 간단히 보고한다.
- 실패하면 대상 경로, 대상 ID, 실패한 단계, 추정 원인, 다음 확인 값을 짧게 보고한다.
- 사용자가 화면 확인을 요청한 경우에만 마지막에 `NavigateTo(...)`를 호출했다고 언급한다.
"""

    def _codex_builtin_internal_instructions_text_macos(self) -> str:
        return """# 코덱스 전용 OneNote 조작 지침 (macOS)

이 문서는 사용자 스킬이 아니다. OneNote for Mac 작업을 수행하는 Codex가 항상 전제로 삼는 내부 실행 지침이다.

## 적용 순서

1. 사용자 요청에서 목표, 대상 경로, 출력 형식, 금지 조건, 주의 조건을 먼저 분리한다.
2. 사용자 스킬은 글쓰기 형식과 정리 방식만 적용하고, OneNote 조작 방식은 이 문서와 작업별 템플릿에서 고른다.
3. 대상은 명시 경로, 저장된 위치 캐시, 현재 열린 전자필기장/섹션, 제한된 UI 조회 순서로 확정한다.
4. 변경 전에는 현재 보이는 전자필기장/섹션/페이지가 맞는지 먼저 확인한다.
5. 완료 후에는 접근성/UI 조회 결과와 화면상 위치를 기준으로 검증한다.
6. 최종 보고에는 변경 항목, 대상 경로, 검증 결과, 남은 확인 사항만 짧게 남긴다.

## 사용자 스킬과의 경계

- 사용자 스킬에서는 `## Instructions`만 작업에 맞게 적용한다.
- 사용자 요청문, 작업 주문서, 스킬 호출문에는 이 문서 전문이나 내부 템플릿 전문을 붙이지 않는다.
- 글쓰기 형식, 정리 방식, 이름 규칙처럼 요청마다 달라지는 기준은 사용자 스킬에 둔다.
- OneNote for Mac의 접근성/UI 자동화 순서, 대상 경로 판정, 본문 수정 방식, 검증 절차는 코덱스 전용 지침에 둔다.
- 사용자가 내부 구현 설명을 요청하지 않았다면 접근성 트리나 AppleScript 세부 호출을 길게 설명하지 않는다.

## 대상 판정 원칙

- 사용자가 경로 또는 전자필기장/섹션/페이지 이름을 직접 준 경우 상위 경로까지 함께 확인한다.
- `docs/codex/onenote-targets.json`의 고정 대상이 있으면 위치 캐시보다 먼저 확인한다.
- `docs/codex/onenote-location-cache.json`은 빠른 경로 후보로만 사용하고, 실패하면 실제 OneNote 화면에서 다시 확인한다.
- macOS에서는 COM ID를 가정하지 않는다. 경로 문자열, 현재 열린 전자필기장 이름, 현재 선택된 섹션/페이지를 우선 식별자로 사용한다.
- 이름만으로 대상을 확정하지 않는다. 같은 이름이 여러 개면 상위 전자필기장/섹션 그룹까지 맞는지 확인한다.

## 안전 실행 원칙

- OneNote for Mac에서는 접근성 트리와 UI 자동화를 우선 사용한다.
- 가능한 경우 메뉴/단축키/표준 버튼을 먼저 사용하고, 좌표 클릭은 최후 수단으로만 사용한다.
- 쓰기 작업 전에는 현재 선택 위치를 다시 읽어 대상이 맞는지 확인한다.
- 페이지 생성, 섹션 생성, 제목 변경, 본문 추가 후에는 왼쪽 패널과 현재 본문에서 결과를 다시 확인한다.
- 삭제, 덮어쓰기, 대량 이동, 대량 복제는 영향 범위를 다시 확인한 뒤 실행한다.
- 삭제보다 이동/재분류를 우선한다. OneNote가 자동으로 만든 빈 페이지도 삭제 전 사용자 확인을 받는다.
- 재분류 중 같은 제목이 대상 섹션에 이미 있더라도 원본에 남은 실제 글은 먼저 대상 섹션으로 이동한다. 이후 동일 제목 중복은 삭제하지 말고 중복 후보로 기록해 사용자 확인을 기다린다.
- 대량 이동/정리는 페이지마다 별도 `osascript`를 띄우지 말고 가능한 한 하나의 AppleScript 루프로 묶는다.
- OneNote 동기화가 느리면 같은 쓰기 작업을 연속 반복하지 말고, 짧게 기다린 뒤 다시 읽어 검증한다.

## 대량 이동/정리 최적화

- 페이지 이동은 상단 메뉴 `전자 필기장 > 페이지 > 페이지 이동 위치...`를 우선 사용한다.
- 같은 대상 섹션으로 여러 페이지를 옮길 때는 첫 페이지만 전체 이동 시트에서 대상을 고르고, 이후 페이지는 `"{대상 섹션}"에 페이지 다시 이동` 메뉴를 재사용한다.
- 이동 시트의 대상 목록은 항목이 많으면 매우 느리다. 매번 전체 outline을 스캔하지 말고 대상 섹션의 접근성 row index 또는 보이는 row band를 캐시한다.
- 캐시한 row가 맞지 않거나 이동 시트 구조가 바뀌면 시트를 취소하고 현재 화면을 다시 읽은 뒤 한 번만 재스캔한다.
- 원본 섹션을 선택한 직후 페이지 목록이 늦게 갱신될 수 있다. 대상 페이지 제목이 실제 목록에 나타날 때까지 짧은 wait/retry를 둔다.
- 페이지 row 제목은 `row`의 `entire contents` 중 제목 텍스트 child에서 읽고, 섹션 row는 첫 UI element의 `AXDescription` 후보를 우선 확인한다.
- 최종 검수는 원본 섹션 카운트와 대상 섹션 내 동일 제목 개수를 함께 출력한다. 원본 0개와 중복 후보를 분리해서 보고하면 삭제 확인 없이도 재분류 상태를 설명할 수 있다.
- 긴 배치는 `FULL`, `AGAIN`, `MISS_PAGE`, `MISS_DEST`처럼 결과 로그를 남기고, 마지막에 원본 섹션과 대상 섹션을 다시 스캔한다.

## 페이지 제목/본문 입력 안정화

- 페이지 제목을 바꿀 때는 캔버스 좌표 클릭보다 상단 메뉴 `전자 필기장 > 페이지 > 페이지 이름 바꾸기`를 우선 사용한다.
- 제목 영역처럼 보이는 곳을 좌표 클릭하면 제목이 아니라 본문 노트 컨테이너가 생길 수 있다. 제목 수정 전후에는 왼쪽 페이지 목록 제목을 반드시 확인한다.
- macOS 키보드 입력이 한글 제목 전체를 누락하면 숫자/ASCII 접두어를 먼저 입력하고, 나머지 한글은 클립보드 붙여넣기로 이어 붙인다.
- 본문 붙여넣기는 제목 확정 후 새 줄 또는 본문 컨테이너에 포커스를 둔 상태에서 수행한다. 본문 첫 줄이 페이지 목록 제목으로 바뀌면 `페이지 이름 바꾸기`로 즉시 복구한다.

## 작업별 내부 문서

- 페이지 추가/본문 추가/제목 변경: `onenote-com-templates.md`의 macOS 템플릿
- 섹션 생성/섹션 그룹 생성: `onenote-com-templates.md`의 macOS 템플릿
- 위치 판정/계층 읽기: `원노트-대상ID-찾기.md`와 저장된 위치 캐시
- Windows 참고 문서: `onenote-com-playbook.md`, `onenote-windows-internal.md`

## 빠른 라우팅

- 페이지 추가는 현재 섹션을 열고 새 페이지를 만든 뒤 제목과 본문을 입력한다.
- 섹션 생성은 대상 전자필기장 또는 섹션 그룹을 연 상태에서 새 섹션 UI를 사용한다.
- 섹션 그룹 생성은 대상 전자필기장의 왼쪽 패널에서 그룹 생성 UI를 사용한다.
- 전자필기장 열기는 `onenote:` 링크, 웹 링크, 파일 경로 중 가능한 값을 우선 사용한다.
- 링크 생성은 앱 링크/공유 링크를 우선 사용하고, 직접 ID 링크를 요구하지 않는다.
- 페이지 재분류는 대상 섹션을 먼저 만들고, 원본 섹션별 페이지 목록을 스냅샷으로 잡은 뒤 대상별로 묶어서 이동한다.
- 섹션 자체를 옮겨야 할 때도 Windows COM ID를 가정하지 말고 현재 열린 전자필기장/섹션 그룹 경로와 화면상 선택 상태를 기준으로 처리한다.

## 검증 기준

- 페이지 추가/수정: 새 제목과 본문 일부가 현재 페이지와 목록에 모두 보이는지 확인한다.
- 섹션 생성: 왼쪽 섹션 목록에 새 섹션 이름이 나타나는지 확인한다.
- 섹션 그룹 생성: 전자필기장 패널에서 새 그룹이 나타나는지 확인한다.
- 전자필기장 열기/닫기: 왼쪽 패널에서 해당 전자필기장이 나타나거나 사라졌는지 확인한다.
- 대량 이동/재분류: 원본 섹션에 남은 페이지와 대상 섹션에 들어간 페이지를 다시 읽어 누락 항목을 확인한다.
- 내보내기: 저장 경로에 파일이 실제로 생성됐는지 확인한다.
- 위치 캐시 갱신: 저장한 경로로 다시 읽었을 때 같은 전자필기장/섹션이 보이는지 확인한다.

## 실패 대응

- 실패하면 먼저 잘못된 대상 경로, 전자필기장 미오픈, 동기화 지연, 접근성 권한 문제를 구분한다.
- 대상 찾기 실패 시 전체 창을 무한 탐색하지 말고 상위 전자필기장부터 제한적으로 다시 연다.
- 이동 시트가 열린 채 멈추면 새 작업을 시작하지 말고 시트 상태를 먼저 확인한 뒤 취소 또는 선택 완료 중 하나로 정리한다.
- UI 구조가 예상과 다르면 현재 화면 구성을 다시 읽고, 가능한 다른 표준 경로를 한 번만 시도한다.
- 쓰기 성공 여부가 애매하면 추가 쓰기를 하지 말고 조회/화면 검증부터 수행한다.

## 보고 기준

- 성공하면 만든/수정한 OneNote 항목, 대상 경로, 검증 결과만 간단히 보고한다.
- 실패하면 대상 경로, 실패한 단계, 추정 원인, 다음 확인 값을 짧게 보고한다.
"""

    def _codex_builtin_internal_instructions_text(self) -> str:
        if _codex_active_platform_key() == CODEX_PLATFORM_MACOS:
            return self._codex_builtin_internal_instructions_text_macos()
        return self._codex_builtin_internal_instructions_text_windows()

    def _codex_internal_instructions_text(self) -> str:
        path = self._codex_internal_instructions_path()
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    text = f.read().strip()
                if text:
                    return text
        except Exception:
            pass
        return self._codex_builtin_internal_instructions_text()

    def _ensure_codex_internal_instructions_file(self) -> str:
        path = self._codex_internal_instructions_path()
        if not os.path.exists(path):
            legacy_path = self._codex_internal_instructions_legacy_path()
            if (
                _codex_active_platform_key() == CODEX_PLATFORM_WINDOWS
                and os.path.exists(legacy_path)
            ):
                try:
                    with open(legacy_path, "r", encoding="utf-8") as f:
                        legacy_text = f.read().strip()
                    if legacy_text:
                        self._write_text_file_atomic(path, legacy_text + "\n")
                        return path
                except Exception:
                    pass
            self._write_text_file_atomic(
                path,
                self._codex_builtin_internal_instructions_text() + "\n",
            )
        return path

    def _save_codex_internal_instructions(self) -> None:
        editor = getattr(self, "codex_internal_instructions_editor", None)
        if editor is None:
            return
        try:
            path = self._codex_internal_instructions_path()
            text = editor.toPlainText().strip()
            if not text:
                text = self._codex_builtin_internal_instructions_text()
                editor.setPlainText(text)
            self._write_text_file_atomic(path, text + "\n")
            self._update_codex_codegen_previews()
            try:
                self.connection_status_label.setText(f"코덱스 전용 지침 저장 완료: {path}")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "코덱스 전용 지침 저장 실패", str(e))

    def _reload_codex_internal_instructions(self) -> None:
        editor = getattr(self, "codex_internal_instructions_editor", None)
        if editor is None:
            return
        try:
            self._ensure_codex_internal_instructions_file()
            editor.setPlainText(self._codex_internal_instructions_text())
            self._update_codex_codegen_previews()
            try:
                self.connection_status_label.setText("코덱스 전용 지침을 다시 불러왔습니다.")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "코덱스 전용 지침 불러오기 실패", str(e))

    def _copy_codex_internal_instructions_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_internal_instructions_text())
        try:
            self.connection_status_label.setText("코덱스 전용 지침을 클립보드에 복사했습니다.")
        except Exception:
            pass

    def _open_codex_instructions_folder(self) -> None:
        try:
            os.makedirs(self._codex_instructions_dir(), exist_ok=True)
            open_path_in_system(self._codex_instructions_dir())
        except Exception as e:
            QMessageBox.warning(self, "코덱스 전용 지침 폴더 열기 실패", str(e))

_publish_context(globals())

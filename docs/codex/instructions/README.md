# Codex Internal Instructions

이 폴더는 사용자 스킬이 아니라 Codex가 OneNote 작업을 수행할 때 항상 따르는 내부 지침을 저장한다.

사용자 스킬은 `docs/codex/skills`에서 관리한다. 글쓰기 형식, 회의록 정리, 업무일지처럼 사용자 요청마다 달라지는 처리 방식만 사용자 스킬에 둔다.

이제 내부 지침은 Windows와 macOS를 분리해서 관리한다. Windows 자료는 유지하고, macOS 흐름은 OneNote for Mac 화면 구조와 접근성/UI 자동화 기준으로 따로 확장한다.

## 사용 순서

1. 현재 플랫폼에 맞는 기본 라우팅 문서를 먼저 본다.
2. 실제 실행은 작업 종류에 맞는 `원노트-*.md` 또는 앱 내부 템플릿을 따른다.
3. 운영 중 헷갈리는 기준은 플랫폼 문서와 `onenote-com-playbook.md`를 함께 참고한다.
4. 사용자 스킬에는 OneNote 실행 엔진, COM/접근성 호출 방식, 내부 PowerShell 템플릿을 넣지 않는다.
5. 삭제, 덮어쓰기, 대량 작업은 대상과 영향 범위를 다시 확인한 뒤 실행한다.

## 플랫폼 문서

- `onenote-windows-internal.md`: Windows OneNote COM 기준 기본 내부 지침
- `onenote-macos-internal.md`: OneNote for Mac 화면/UI 기준 기본 내부 지침
- `onenote-com-internal.md`: 기존 Windows 중심 문서. 하위 호환용으로 유지

## 공통 참고 문서

- `onenote-com-playbook.md`: Windows COM 운영 노하우와 작업 라우팅 참고서
- `onenote-com-templates.md`: 앱 내부 작업 템플릿 기준
- `원노트-페이지-추가.md`: 페이지 생성 내부 절차
- `원노트-섹션-생성.md`: 섹션 생성 내부 절차
- `원노트-섹션그룹-생성.md`: 섹션 그룹 생성 내부 절차
- `원노트-전자필기장-생성.md`: 전자필기장 생성/열기 내부 절차
- `원노트-전자필기장-복제.md`: 전자필기장 복제 내부 절차
- `원노트-대상ID-찾기.md`: 대상 캐시/위치 판정 참고 절차

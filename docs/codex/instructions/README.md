# Codex-Only Instructions

이 폴더는 사용자 스킬이 아니라 Codex가 OneNote 작업을 수행할 때 항상 따르는 내부 지침을 저장한다.

사용자 스킬은 `docs/codex/skills`에서 관리한다. 글쓰기 형식, 회의록 정리, 업무일지처럼 사용자 요청마다 달라지는 처리 방식만 사용자 스킬에 둔다.

OneNote COM 호출 방식, 대상 ID 우선순위, 작업별 PowerShell 패턴, 완료 검증 기준은 이 폴더에서 관리한다.

## 사용 순서

1. `onenote-com-internal.md`를 기본 라우팅 문서로 본다.
2. 실제 실행은 작업 종류에 맞는 `원노트-*.md` 문서를 따른다.
3. 운영 중 헷갈리는 기준은 `onenote-com-playbook.md`에서 확인한다.
4. 사용자 스킬에는 OneNote COM 호출 방식이나 PowerShell 템플릿을 넣지 않는다.

## 문서

- `onenote-com-internal.md`: 코덱스 탭의 전용 지침 편집기에서 읽고 쓰는 기본 지침
- `onenote-com-playbook.md`: OneNote COM 운영 노하우와 작업 라우팅
- `onenote-com-templates.md`: 앱 내부 템플릿 관리 기준
- `원노트-페이지-추가.md`: 페이지 생성 내부 절차
- `원노트-섹션-생성.md`: 섹션 생성 내부 절차
- `원노트-섹션그룹-생성.md`: 섹션 그룹 생성 내부 절차
- `원노트-전자필기장-생성.md`: 전자필기장 생성/열기 내부 절차
- `원노트-전자필기장-복제.md`: 전자필기장 복제 내부 절차
- `원노트-대상ID-찾기.md`: 대상 ID 캐시 갱신 내부 절차

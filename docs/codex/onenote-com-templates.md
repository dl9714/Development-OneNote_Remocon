# OneNote COM Templates

Codex 탭의 템플릿 선택기와 같은 기준으로 관리하는 작업 템플릿 목록이다.

## 기본 템플릿

- 페이지 추가: `Section ID`로 `CreateNewPage`, `GetPageContent`, `UpdatePageContent`, `NavigateTo`를 호출한다.
- 새 섹션 생성: `SectionGroup ID`로 `OpenHierarchy("섹션명.one", sectionGroupId, ref newId, cftSection)`를 호출한다.
- 새 섹션 그룹 생성: 부모 섹션그룹 또는 전자필기장 ID로 `OpenHierarchy("그룹명", parentId, ref newId, cftFolder)`를 호출한다.
- 새 전자필기장 생성: 저장 경로를 확정한 뒤 `OpenHierarchy(notebookPath, "", ref newId, cftNotebook)`를 호출한다.

## 스킬 설계 기준

스킬은 아래 정보를 한 묶음으로 저장해야 한다.

- 자연어 지시문
- 대상 OneNote 경로
- 캐시된 OneNote ID
- PowerShell COM 템플릿
- 완료 검증 방법
- 실패 시 fallback 절차

## 대상 캐시 연동

기본 대상은 `docs/codex/onenote-targets.json`에서 읽는다. 앱의 `코덱스` 탭에서 대상 경로와 ID를 저장하면 요청 생성기, 템플릿, 이후 스킬 설계가 같은 값을 공유한다.

새 대상 ID를 찾은 뒤에는 JSON에 저장하고, 다음 요청부터는 저장된 `Section ID` 또는 `SectionGroup ID`를 우선 사용한다.

## 우선 만들 스킬

- 나의 기본 글쓰기 형식
- 임시 메모 페이지 추가
- 섹션 생성
- 섹션 그룹 생성
- 전자필기장 생성
- 업무일지 생성
- 회의록 생성

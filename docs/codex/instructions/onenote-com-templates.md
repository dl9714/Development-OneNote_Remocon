# OneNote COM Templates

Codex 탭의 템플릿 선택기와 같은 기준으로 관리하는 내부 작업 템플릿 목록이다.

## 기본 템플릿

- 페이지 추가: `Section ID`로 `CreateNewPage`, `GetPageContent`, `UpdatePageContent`, `NavigateTo`를 호출한다.
- 새 섹션 생성: `SectionGroup ID`로 `OpenHierarchy("섹션명.one", sectionGroupId, ref newId, cftSection)`를 호출한다.
- 새 섹션 그룹 생성: 부모 섹션그룹 또는 전자필기장 ID로 `OpenHierarchy("그룹명", parentId, ref newId, cftFolder)`를 호출한다.
- 새 전자필기장 생성: 저장 경로를 확정한 뒤 `OpenHierarchy(notebookPath, "", ref newId, cftNotebook)`를 호출한다.
- 전자필기장 복제: 대상 이름을 `코덱스-{원본 전자필기장명}`으로 정하고 활성 섹션 그룹/섹션/페이지 구조를 재작성한다.

## 내부 템플릿 관리 기준

OneNote 조작 템플릿은 사용자 스킬이 아니라 코덱스 전용 지침으로 관리한다. 템플릿에는 아래 정보를 한 묶음으로 저장한다.

- 대상 OneNote 경로
- 캐시된 OneNote ID
- PowerShell COM 템플릿 또는 템플릿을 만들기 위한 기준
- 완료 검증 방법
- 실패 시 fallback 절차

## 대상 캐시 연동

기본 대상은 `docs/codex/onenote-targets.json`에서 읽는다. OneNote 조회 결과가 있으면 `docs/codex/onenote-location-cache.json`도 참고한다. 앱의 `코덱스` 탭에서 대상 경로와 ID를 저장하면 요청 생성기와 내부 템플릿이 같은 값을 공유한다.

새 대상 ID를 찾은 뒤에는 JSON에 저장하고, 다음 요청부터는 저장된 `Section ID` 또는 `SectionGroup ID`를 우선 사용한다.

## 사용자 스킬과의 경계

- 사용자 스킬에는 글쓰기 형식, 업무일지, 회의록처럼 요청마다 달라지는 처리 기준만 둔다.
- 페이지/섹션/전자필기장 생성 방식은 이 폴더의 내부 지침에서 관리한다.
- 전자필기장 복제처럼 긴 실행 절차도 사용자 스킬이 아니라 이 폴더의 내부 지침에서 관리한다.

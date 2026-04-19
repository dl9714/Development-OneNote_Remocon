# 코덱스 전용 OneNote 조작 지침

이 문서는 사용자 스킬이 아니다. OneNote 작업을 수행하는 Codex가 항상 전제로 삼는 내부 실행 지침이다.

## 적용 순서

1. 사용자 요청과 선택된 사용자 스킬에서 목표, 대상, 작성 형식만 추출한다.
2. OneNote 조작 방식은 이 폴더의 작업별 내부 지침에서 고른다.
3. 저장된 대상 ID가 있으면 전체 탐색보다 ID 직접 호출을 우선한다.
4. 완료 후 화면 캡처가 아니라 `GetHierarchy` 또는 `GetPageContent` 결과로 검증한다.
5. 최종 보고에는 변경 항목과 검증 결과만 짧게 남긴다.

## 사용자 스킬과의 경계

- 사용자 스킬에서는 `## Instructions`만 작업에 맞게 적용한다.
- 사용자 요청문, 작업 주문서, 스킬 호출문에는 이 문서 전문이나 PowerShell 템플릿을 붙이지 않는다.
- 글쓰기 형식, 정리 방식, 이름 규칙처럼 요청마다 달라지는 기준은 사용자 스킬에 둔다.
- COM 호출 순서, 대상 ID 우선순위, XML 수정 방식, 검증 절차는 코덱스 전용 지침에 둔다.
- 사용자가 내부 구현 설명을 요청하지 않았다면 COM 세부 호출을 길게 설명하지 않는다.

## 공통 실행 원칙

- OneNote 조작은 화면 클릭 자동화보다 Windows OneNote COM API를 우선 사용한다.
- PowerShell에서는 `New-Object -ComObject OneNote.Application`으로 연결한다.
- 구조 탐색은 기본적으로 `GetHierarchy('', hsSections, ref xml)`까지만 사용한다.
- `hsPages` 전체 조회는 페이지가 많으면 느리므로 페이지 복제, 페이지 목록 조회, 최종 페이지 수 검증처럼 필요한 경우에만 쓴다.
- `docs/codex/onenote-targets.json`과 `docs/codex/onenote-location-cache.json`에 대상 ID가 있으면 먼저 사용한다.
- 이름이 중복될 수 있으므로 전자필기장, 섹션 그룹, 섹션을 상위 경로까지 함께 제한해 찾는다.
- 새 ID를 찾거나 자주 쓸 대상이 생기면 대상 캐시에 저장한다.
- XML은 문자열 치환보다 XML 파서로 수정한다.
- `OpenHierarchy` 또는 `CreateNewPage` 직후에는 필요하면 잠시 대기하고 `GetHierarchy` 또는 `GetPageContent`로 새 ID가 실제 사용 가능한지 확인한다.

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

## 보고 기준

- 성공하면 만든/수정한 OneNote 항목과 검증 결과만 간단히 보고한다.
- 실패하면 대상 경로, 대상 ID, 실패한 단계, 검증 결과를 짧게 보고한다.
- 사용자가 화면 확인을 요청한 경우에만 마지막에 `NavigateTo(...)`를 호출했다고 언급한다.

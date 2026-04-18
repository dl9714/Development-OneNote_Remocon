# 코덱스 전용 OneNote 조작 지침

이 문서는 사용자 스킬이 아니다. OneNote 작업을 수행하는 Codex가 항상 전제로 삼는 내부 실행 지침이다.

## 기본 원칙

- OneNote 조작은 화면 클릭 자동화보다 Windows OneNote COM API를 우선 사용한다.
- PowerShell에서는 `New-Object -ComObject OneNote.Application`으로 연결한다.
- 구조 탐색은 기본적으로 `GetHierarchy('', hsSections, ref xml)`까지만 사용한다. `hsPages` 전체 조회는 페이지가 많으면 느리므로 필요한 경우에만 제한적으로 사용한다.
- `docs/codex/onenote-targets.json`에 대상 ID가 있으면 전체 계층 탐색보다 ID 직접 호출을 우선한다.
- 완료 검증은 화면 캡처보다 `GetHierarchy` 또는 `GetPageContent` 결과로 확인한다.

## 대상 ID 확인

- 전자필기장, 섹션 그룹, 섹션 이름이 중복될 수 있으므로 상위 경로까지 함께 제한해 찾는다.
- 새 대상 ID를 찾으면 `docs/codex/onenote-targets.json`에 저장한다.
- ID가 비었거나 실패할 때만 `GetHierarchy('', hsSections, ref xml)`로 다시 찾아 갱신한다.

## 페이지 추가

- 페이지 추가는 대상 `Section ID`로 `CreateNewPage`를 호출한다.
- 새 페이지 XML은 `GetPageContent`로 읽고 제목과 본문을 수정한다.
- 반영은 `UpdatePageContent`로 처리한다.
- 작성 후 `GetPageContent`로 제목과 본문 일부가 들어갔는지 확인한다.
- 사용자가 화면 확인을 원할 때만 마지막에 `NavigateTo(pageId)`를 호출한다.

## 섹션 생성

- 섹션 생성은 대상 `SectionGroup ID`를 먼저 확보한다.
- `OpenHierarchy("섹션명.one", sectionGroupId, ref newId, cftSection)`를 사용한다.
- 섹션명은 파일명으로 쓸 수 없는 문자를 정리하고 `.one` 확장자를 붙인다.
- 생성 후 `GetHierarchy(sectionGroupId, hsSections, ref xml)`로 새 섹션 이름을 확인한다.

## 섹션 그룹 생성

- 섹션 그룹의 부모는 전자필기장 또는 기존 섹션 그룹이어야 한다. 섹션 안에는 만들 수 없다.
- `OpenHierarchy("그룹명", parentId, ref newId, cftFolder)`를 사용한다.
- 생성 후 `GetHierarchy(parentId, hsSections, ref xml)`로 새 그룹 이름을 확인한다.
- 자주 쓸 그룹이면 새 ID를 대상 캐시에 저장한다.

## 전자필기장 생성/열기

- 새 전자필기장은 OneNote가 열 수 있는 로컬 또는 동기화 경로를 먼저 확정한다.
- 웹 OneDrive URL을 그대로 쓰지 말고 실제 파일 시스템 경로를 확인한다.
- `OpenHierarchy(notebookPath, "", ref newId, cftNotebook)`를 사용한다.
- 완료 후 `GetHierarchy("", hsNotebooks, ref xml)` 또는 `hsSections`로 열린 전자필기장을 확인한다.

## 보고 기준

- 사용자 요청문, 작업 주문서, 사용자 스킬 호출문에는 이 문서의 전문이나 PowerShell 템플릿을 붙이지 않는다.
- 사용자가 별도로 내부 절차를 요청하지 않았다면 코덱스가 필요할 때 이 폴더의 지침과 대상 캐시를 직접 조회한다.
- 사용자가 내부 구현 절차를 요청하지 않았다면 COM 호출 세부사항을 길게 설명하지 않는다.
- 최종 보고에는 변경한 OneNote 항목과 검증 결과만 간단히 남긴다.
- 실패하면 어느 대상 ID, 경로, 검증 단계에서 막혔는지 짧게 보고한다.

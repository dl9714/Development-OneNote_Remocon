# Codex OneNote COM Playbook

이 문서는 Codex 또는 다음 AI가 OneNote를 빠르게 조작할 때 쓰는 운영 노하우다.

## 핵심 원칙

- OneNote 조작은 화면 클릭 자동화보다 Windows OneNote COM API를 우선 사용한다.
- PowerShell에서 `New-Object -ComObject OneNote.Application`으로 OneNote에 연결한다.
- 구조 탐색은 `GetHierarchy('', hsSections, ref xml)`까지만 우선 사용한다. `hsPages`는 페이지가 많으면 느리다.
- 대상 전자필기장/섹션그룹/섹션 ID를 한 번 찾으면 이후 작업은 ID를 바로 써서 호출한다.
- 검증은 화면 캡처보다 `GetHierarchy` 또는 `GetPageContent` 결과로 확인한다.

## 빠른 대상 경로

- 전자필기장: `생산성도구-임시 메모`
- 섹션 그룹: `A 미정리-생성 메모`
- 기본 섹션: `미정리`

현재 확인된 빠른 ID:

```text
SectionGroup ID:
{2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}

미정리 Section ID:
{175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}
```

## 대상 ID 캐시

앱의 `코덱스` 탭은 자주 쓰는 OneNote 대상을 `docs/codex/onenote-targets.json`에 저장한다.

- `path`: 사람이 읽는 대상 경로
- `notebook`: 전자필기장 이름
- `section_group`: 섹션 그룹 이름
- `section`: 섹션 이름
- `section_group_id`: 섹션 생성에 바로 쓰는 ID
- `section_id`: 페이지 추가에 바로 쓰는 ID

다음 AI는 먼저 이 JSON을 확인하고, ID가 있으면 전체 계층 탐색 없이 바로 작업한다. ID가 비었거나 실패할 때만 `GetHierarchy('', hsSections, ref xml)`로 다시 찾아 갱신한다.

## 대상 ID 찾기 절차

새 전자필기장/섹션그룹/섹션을 자주 쓰게 되면 앱의 `코덱스` 탭에서 `ID 찾기 스크립트 복사`를 사용한다.

1. 대상 이름, 전자필기장, 섹션 그룹, 섹션 이름을 입력한다.
2. `ID 찾기 스크립트 복사`로 PowerShell 스크립트를 복사한다.
3. 스크립트는 `GetHierarchy('', hsSections, ref xml)`만 사용해 ID를 찾고 JSON을 출력한다.
4. 스크립트는 출력 JSON을 클립보드에도 복사한다.
5. 앱에서 `클립보드 JSON 저장`을 누르면 `docs/codex/onenote-targets.json` 캐시에 반영된다.
6. 이후 페이지 추가는 `section_id`, 섹션 생성은 `section_group_id`를 바로 사용한다.

## PowerShell 기본 연결

```powershell
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$xml = ""
$one.GetHierarchy(
    "",
    [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections,
    [ref]$xml
)
```

## 페이지 추가 패턴

```powershell
$sectionId = "{175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}"
$pageId = ""

$one.CreateNewPage(
    $sectionId,
    [ref]$pageId,
    [Microsoft.Office.Interop.OneNote.NewPageStyle]::npsBlankPageWithTitle
)

$pageXml = ""
$one.GetPageContent($pageId, [ref]$pageXml)

# pageXml 안의 one:Title / one:Outline을 수정한 뒤 반영한다.
$one.UpdatePageContent($pageXml)
$one.NavigateTo($pageId)
```

## 섹션 추가 패턴

```powershell
$sectionGroupId = "{2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}"
$newSectionId = ""

$one.OpenHierarchy(
    "코덱스가 만든 섹션.one",
    $sectionGroupId,
    [ref]$newSectionId,
    [Microsoft.Office.Interop.OneNote.CreateFileType]::cftSection
)

$one.NavigateTo($newSectionId)
```

## 다음에 템플릿화할 기본 기능

- 새 전자필기장 생성
- 새 섹션 그룹 생성
- 새 섹션 생성
- 페이지 추가
- 기본 글쓰기 형식 적용
- 사용자가 앱에서 직접 만드는 Codex 스킬 저장/호출

## 다음 AI에게 줄 지시문

```text
OneNote는 화면 클릭으로 조작하지 말고 Windows OneNote COM API를 PowerShell에서 사용해라.
먼저 New-Object -ComObject OneNote.Application으로 연결하고,
GetHierarchy('', hsSections, ref xml)로 전자필기장 구조를 XML로 읽어라.
대상 위치는 생산성도구-임시 메모 > A 미정리-생성 메모 이다.

페이지 생성은 대상 Section ID를 찾은 뒤 CreateNewPage + UpdatePageContent를 사용해라.
섹션 생성은 대상 SectionGroup ID를 찾은 뒤 OpenHierarchy("섹션명.one", sectionGroupId, ref newId, cftSection)를 사용해라.
검증은 화면 캡처보다 GetHierarchy 또는 GetPageContent로 해라.
```

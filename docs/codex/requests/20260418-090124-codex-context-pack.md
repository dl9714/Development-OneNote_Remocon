# 코덱스 컨텍스트 패키지
생성 시각: 2026-04-18 09:01:24

이 문서는 다음 코덱스 작업에 필요한 현재 OneNote 작업 위치, 요청문, 스킬 호출문, 템플릿, 주문번호표를 한 번에 묶은 자료입니다.

## 바로 실행할 지시

아래 `스킬 호출문`과 `작업 주문서`를 기준으로 OneNote 작업을 수행해라. 화면 클릭보다 OneNote COM API 조회/수정/검증을 우선한다.

## 초압축 실행 프롬프트

```text
OneNote COM API로 아래 작업을 바로 수행해라. 화면 클릭보다 COM 조회/수정/검증을 우선한다.

스킬:
- 주문번호: SK-001
- 스킬명: 나의 기본 글쓰기 형식

작업:
- 유형: 페이지 추가
- 제목/이름: ??? ??
- 대상 경로: 생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리
- SectionGroup ID: {2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}
- Section ID: {175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}

본문:
?? ?? ??

완료 후 `GetHierarchy` 또는 `GetPageContent`로 검증하고, 변경한 항목과 검증 결과만 간단히 보고해라.

```

## 검토 프롬프트

```text
아래 OneNote 작업 계획을 검토해라. 실행하지 말고 위험 요소, 빠진 검증, 잘못된 대상 ID 가능성만 점검해라.

## 현재 상태

코덱스 현재 상태

- 스킬 파일: 6개
- 저장된 작업 주문서/패키지: 1개
- 저장된 작업 위치: 1개
- 요청 초안: 없음

현재 요청
- 작업: 페이지 추가
- 제목/이름: ??? ??
- 요청 대상: 생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리

선택 스킬
- 주문번호: SK-001
- 스킬명: 나의 기본 글쓰기 형식

작업 위치
- 위치 이름: 임시 메모 - 미정리
- 전자필기장: 생산성도구-임시 메모
- 섹션 그룹: A 미정리-생성 메모
- 섹션: 미정리


## 실행 체크리스트

# 코덱스 OneNote 실행 체크리스트
생성 시각: 2026-04-18 09:01:24

## 현재 작업

- 작업: 페이지 추가
- 제목/이름: ??? ??
- 대상 경로: 생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리
- SectionGroup ID: {2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}
- Section ID: {175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}

## 실행 전

- OneNote COM 연결을 먼저 확인한다.
- 대상 전자필기장/섹션/섹션 그룹을 `GetHierarchy`로 조회한다.
- 내부 ID가 있으면 이름 탐색보다 ID 직접 사용을 우선한다.
- 삭제/이동 작업은 작업 전 구조를 기록한다.

## 실행 중

- 화면 클릭 자동화보다 OneNote COM API를 우선 사용한다.
- 페이지 추가는 `CreateNewPage`, `GetPageContent`, `UpdatePageContent` 순서로 처리한다.
- 섹션/섹션 그룹 생성은 `OpenHierarchy`를 사용한다.
- 예외가 나면 OneNote 상태와 대상 ID를 다시 조회한다.

## 실행 후 검증

- `GetHierarchy` 또는 `GetPageContent`로 결과를 확인한다.
- 생성/수정한 제목과 본문 일부가 실제 XML에 있는지 확인한다.
- 검증 결과와 다음 확인 항목을 사용자에게 짧게 보고한다.


## 요청문

OneNote는 화면 클릭으로 조작하지 말고 Windows OneNote COM API를 PowerShell에서 사용해라.
먼저 `New-Object -ComObject OneNote.Application`으로 연결하고,
`GetHierarchy('', hsSections, ref xml)`로 전자필기장/섹션그룹/섹션 구조를 확인해라.

작업:
페이지 추가

대상 경로:
생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리

제목/이름:
??? ??

본문/내용:
?? ?? ??

내부 빠른 위치 정보 (코덱스용):
- SectionGroup ID: {2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}
- Section ID: {175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}

코덱스 전용 지침:
- `docs/codex/instructions/onenote-com-internal.md`를 따른다.


검토 결과는 다음 형식으로 짧게 작성해라.
- 실행 전 막아야 할 문제:
- 대상 위치 확인 필요:
- COM API 사용 시 주의:
- 검증 방법:

```

## 작업 분해 프롬프트

```text
아래 OneNote 작업을 실행 가능한 하위 작업으로 분해해라. 아직 실행하지 말고 순서, 필요한 조회, 검증 기준만 작성해라.

## 현재 요청

OneNote는 화면 클릭으로 조작하지 말고 Windows OneNote COM API를 PowerShell에서 사용해라.
먼저 `New-Object -ComObject OneNote.Application`으로 연결하고,
`GetHierarchy('', hsSections, ref xml)`로 전자필기장/섹션그룹/섹션 구조를 확인해라.

작업:
페이지 추가

대상 경로:
생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리

제목/이름:
??? ??

본문/내용:
?? ?? ??

내부 빠른 위치 정보 (코덱스용):
- SectionGroup ID: {2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}
- Section ID: {175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}

코덱스 전용 지침:
- `docs/codex/instructions/onenote-com-internal.md`를 따른다.


## 현재 작업 위치

```json
{
  "version": 1,
  "targets": [
    {
      "name": "임시 메모 - 미정리",
      "path": "생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리",
      "notebook": "생산성도구-임시 메모",
      "section_group": "A 미정리-생성 메모",
      "section": "미정리",
      "section_group_id": "{2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}",
      "section_id": "{175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}"
    }
  ]
}
```

출력 형식:
- 목표:
- 선행 조회:
- 작업 순서:
- 검증:
- 실패 시 복구:

```

## 완료 보고 템플릿

```markdown
# OneNote 작업 완료 보고 템플릿

## 요청

- 작업: 페이지 추가
- 제목/이름: ??? ??
- 대상: 생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리

## 수행한 작업

-

## 변경한 OneNote 항목

- 전자필기장:
- 섹션 그룹:
- 섹션:
- 페이지:

## 검증 결과

- 조회 방식: GetHierarchy / GetPageContent
- 확인한 값:
- 결과:

## 남은 확인 사항

-

```

## 스킬 추천 리포트

```markdown
# 코덱스 스킬 추천 리포트

## 현재 요청 요약

페이지 추가
생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리
??? ??
?? ?? ??

## 추천 스킬

| 점수 | 주문번호 | 스킬 | 근거 키워드 | 파일 |
| ---: | --- | --- | --- | --- |
| 30 | SK-001 | 나의 기본 글쓰기 형식 | 메모, 생성, 페이지 | `나의-기본-글쓰기-형식.md` |
| 20 | SK-002 | 원노트 페이지 추가 | 추가, 페이지 | `원노트-페이지-추가.md` |
| 10 | SK-003 | 원노트 섹션 생성 | 생성 | `원노트-섹션-생성.md` |
| 10 | SK-004 | 원노트 섹션 그룹 생성 | 생성 | `원노트-섹션그룹-생성.md` |
| 10 | SK-005 | 원노트 전자필기장 생성 | 생성 | `원노트-전자필기장-생성.md` |

추천 점수는 현재 요청의 단어가 스킬 이름, 호출 조건, 본문과 겹치는 정도로 계산한다.

```

## 스킬 호출문

이 프로젝트의 코덱스 & 원노트 스킬 주문번호 `SK-001`를 조회해서 그대로 따라라.

조회 위치:
- docs/codex/skills/skill-order-index.md
- docs/codex/skills/*.md

스킬:
- 주문번호: SK-001
- 이름: 나의 기본 글쓰기 형식
- 호출 조건: OneNote에 새 페이지를 만들거나 Codex가 사용자의 메모/업무 기록을 작성할 때 사용한다.

작업 위치:
생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리

사용자 요청:
??? ??

추가 내용:
?? ?? ??

실행 기준:
- 사용자 스킬 파일에서는 `## Instructions`만 적용한다.
- OneNote는 가능하면 화면 클릭보다 COM API로 조작한다.
- 완료 검증은 가능한 한 OneNote 내부 조회 결과로 확인한다.


## 작업 위치 JSON

```json
{
  "version": 1,
  "targets": [
    {
      "name": "임시 메모 - 미정리",
      "path": "생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리",
      "notebook": "생산성도구-임시 메모",
      "section_group": "A 미정리-생성 메모",
      "section": "미정리",
      "section_group_id": "{2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}",
      "section_id": "{175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}"
    }
  ]
}
```

## 작업 위치 진단

```markdown
# 코덱스 작업 위치 진단

생성 시각: 2026-04-18 09:01:24
작업 위치 수: 1

## 문제

- 문제 없음

## 목록

| 이름 | 경로 | Section ID | SectionGroup ID |
| --- | --- | --- | --- |
| 임시 메모 - 미정리 | 생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리 | {175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0} | {2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0} |

```

## 페이지 읽기 결과 요약

```markdown

```

## 현재 상태

```text
코덱스 현재 상태

- 스킬 파일: 6개
- 저장된 작업 주문서/패키지: 1개
- 저장된 작업 위치: 1개
- 요청 초안: 없음

현재 요청
- 작업: 페이지 추가
- 제목/이름: ??? ??
- 요청 대상: 생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리

선택 스킬
- 주문번호: SK-001
- 스킬명: 나의 기본 글쓰기 형식

작업 위치
- 위치 이름: 임시 메모 - 미정리
- 전자필기장: 생산성도구-임시 메모
- 섹션 그룹: A 미정리-생성 메모
- 섹션: 미정리

```

## 요청문

```text
OneNote는 화면 클릭으로 조작하지 말고 Windows OneNote COM API를 PowerShell에서 사용해라.
먼저 `New-Object -ComObject OneNote.Application`으로 연결하고,
`GetHierarchy('', hsSections, ref xml)`로 전자필기장/섹션그룹/섹션 구조를 확인해라.

작업:
페이지 추가

대상 경로:
생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리

제목/이름:
??? ??

본문/내용:
?? ?? ??

내부 빠른 위치 정보 (코덱스용):
- SectionGroup ID: {2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}
- Section ID: {175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}

코덱스 전용 지침:
- `docs/codex/instructions/onenote-com-internal.md`를 따른다.

```

## 작업 주문서

````markdown
# 코덱스 작업 주문서

생성 시각: 2026-04-18 09:01:24

## 스킬 호출

이 프로젝트의 코덱스 & 원노트 스킬 주문번호 `SK-001`를 조회해서 그대로 따라라.

조회 위치:
- docs/codex/skills/skill-order-index.md
- docs/codex/skills/*.md

스킬:
- 주문번호: SK-001
- 이름: 나의 기본 글쓰기 형식
- 호출 조건: OneNote에 새 페이지를 만들거나 Codex가 사용자의 메모/업무 기록을 작성할 때 사용한다.

작업 위치:
생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리

사용자 요청:
??? ??

추가 내용:
?? ?? ??

실행 기준:
- 사용자 스킬 파일에서는 `## Instructions`만 적용한다.
- OneNote는 가능하면 화면 클릭보다 COM API로 조작한다.
- 완료 검증은 가능한 한 OneNote 내부 조회 결과로 확인한다.


## 선택 스킬

- 주문번호: SK-001
- 스킬명: 나의 기본 글쓰기 형식
- 파일: `미선택`

## 작업 위치

- 위치 이름: 임시 메모 - 미정리
- 작업 경로: 생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리
- 전자필기장: 생산성도구-임시 메모
- 섹션 그룹: A 미정리-생성 메모
- 섹션: 미정리

## 요청문

```text
OneNote는 화면 클릭으로 조작하지 말고 Windows OneNote COM API를 PowerShell에서 사용해라.
먼저 `New-Object -ComObject OneNote.Application`으로 연결하고,
`GetHierarchy('', hsSections, ref xml)`로 전자필기장/섹션그룹/섹션 구조를 확인해라.

작업:
페이지 추가

대상 경로:
생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리

제목/이름:
??? ??

본문/내용:
?? ?? ??

내부 빠른 위치 정보 (코덱스용):
- SectionGroup ID: {2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}
- Section ID: {175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}

코덱스 전용 지침:
- `docs/codex/instructions/onenote-com-internal.md`를 따른다.

```

## 작업 템플릿

```powershell
# OneNote COM: 페이지 추가
# 대상: 생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$sectionId = '{175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}'
$title = '??? ??'
$body = '?? ?? ??'
$pageId = ""
$one.CreateNewPage(
    $sectionId,
    [ref]$pageId,
    [Microsoft.Office.Interop.OneNote.NewPageStyle]::npsBlankPageWithTitle
)

$pageXml = ""
$one.GetPageContent($pageId, [ref]$pageXml)

[xml]$doc = $pageXml
$nsUri = $doc.DocumentElement.NamespaceURI
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("one", $nsUri)

$titleNode = $doc.SelectSingleNode("//one:Title//one:T", $ns)
if ($titleNode -ne $null) {
    $titleNode.InnerText = $title
}

$outline = $doc.CreateElement("one", "Outline", $nsUri)
$oeChildren = $doc.CreateElement("one", "OEChildren", $nsUri)
foreach ($line in ($body -split "`r?`n")) {
    $oe = $doc.CreateElement("one", "OE", $nsUri)
    $t = $doc.CreateElement("one", "T", $nsUri)
    $t.InnerText = $line
    [void]$oe.AppendChild($t)
    [void]$oeChildren.AppendChild($oe)
}
[void]$outline.AppendChild($oeChildren)
[void]$doc.DocumentElement.AppendChild($outline)

$one.UpdatePageContent($doc.OuterXml)
$verifyXml = ""
$one.GetPageContent($pageId, [ref]$verifyXml)
if ($verifyXml -notlike "*$title*") {
    throw "페이지 생성 검증 실패: 제목을 찾지 못했습니다."
}
$one.NavigateTo($pageId)

```

````

## 실행 체크리스트

```text
# 코덱스 OneNote 실행 체크리스트
생성 시각: 2026-04-18 09:01:24

## 현재 작업

- 작업: 페이지 추가
- 제목/이름: ??? ??
- 대상 경로: 생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리
- SectionGroup ID: {2716C2CA-1EA5-4697-9AE7-97380372C026}{1}{B0}
- Section ID: {175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}

## 실행 전

- OneNote COM 연결을 먼저 확인한다.
- 대상 전자필기장/섹션/섹션 그룹을 `GetHierarchy`로 조회한다.
- 내부 ID가 있으면 이름 탐색보다 ID 직접 사용을 우선한다.
- 삭제/이동 작업은 작업 전 구조를 기록한다.

## 실행 중

- 화면 클릭 자동화보다 OneNote COM API를 우선 사용한다.
- 페이지 추가는 `CreateNewPage`, `GetPageContent`, `UpdatePageContent` 순서로 처리한다.
- 섹션/섹션 그룹 생성은 `OpenHierarchy`를 사용한다.
- 예외가 나면 OneNote 상태와 대상 ID를 다시 조회한다.

## 실행 후 검증

- `GetHierarchy` 또는 `GetPageContent`로 결과를 확인한다.
- 생성/수정한 제목과 본문 일부가 실제 XML에 있는지 확인한다.
- 검증 결과와 다음 확인 항목을 사용자에게 짧게 보고한다.

```

## OneNote PowerShell 템플릿

```powershell
# OneNote COM: 페이지 추가
# 대상: 생산성도구-임시 메모 > A 미정리-생성 메모 > 미정리
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$sectionId = '{175CFD85-2C5C-0C63-3116-2598A37ACB11}{1}{B0}'
$title = '??? ??'
$body = '?? ?? ??'
$pageId = ""
$one.CreateNewPage(
    $sectionId,
    [ref]$pageId,
    [Microsoft.Office.Interop.OneNote.NewPageStyle]::npsBlankPageWithTitle
)

$pageXml = ""
$one.GetPageContent($pageId, [ref]$pageXml)

[xml]$doc = $pageXml
$nsUri = $doc.DocumentElement.NamespaceURI
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("one", $nsUri)

$titleNode = $doc.SelectSingleNode("//one:Title//one:T", $ns)
if ($titleNode -ne $null) {
    $titleNode.InnerText = $title
}

$outline = $doc.CreateElement("one", "Outline", $nsUri)
$oeChildren = $doc.CreateElement("one", "OEChildren", $nsUri)
foreach ($line in ($body -split "`r?`n")) {
    $oe = $doc.CreateElement("one", "OE", $nsUri)
    $t = $doc.CreateElement("one", "T", $nsUri)
    $t.InnerText = $line
    [void]$oe.AppendChild($t)
    [void]$oeChildren.AppendChild($oe)
}
[void]$outline.AppendChild($oeChildren)
[void]$doc.DocumentElement.AppendChild($outline)

$one.UpdatePageContent($doc.OuterXml)
$verifyXml = ""
$one.GetPageContent($pageId, [ref]$verifyXml)
if ($verifyXml -notlike "*$title*") {
    throw "페이지 생성 검증 실패: 제목을 찾지 못했습니다."
}
$one.NavigateTo($pageId)

```

## 현재 스킬 초안

````markdown
# 나의 기본 글쓰기 형식

## 주문번호

SK-001

## Trigger

OneNote에 새 페이지를 만들거나 Codex가 사용자의 메모/업무 기록을 작성할 때 사용한다.

## Instructions

목표:
- 사용자의 기본 글쓰기 형식을 OneNote 페이지에 일관되게 적용한다.

형식:
- 제목:
- 목적:
- 핵심 내용:
- 다음 행동:

검증:
- 생성된 페이지 제목과 본문을 GetPageContent로 확인한다.

OneNote 작성 기준:

- 화면 클릭보다 OneNote COM API를 우선 사용한다.
- 대상 섹션 ID가 있으면 코덱스 전용 지침의 페이지 작성 절차를 따른다.
- 생성 후 코덱스 전용 지침 기준으로 제목과 본문 반영을 확인한다.
- 가능하면 마지막에 `NavigateTo(pageId)`로 새 페이지를 화면에 띄운다.

## 내부 지침 분리 안내

- OneNote 조작 방식은 docs/codex/instructions/onenote-com-internal.md에서 관리한다.

````

## 스킬 주문번호표

````markdown
# 코덱스 & 원노트 스킬 주문번호표

사용자가 주문번호로 스킬을 지시하면 이 표에서 해당 Markdown 파일을 찾아 따른다.

| 주문번호 | 스킬 이름 | 파일 | 호출 조건 |
| --- | --- | --- | --- |
| SK-001 | 나의 기본 글쓰기 형식 | `나의-기본-글쓰기-형식.md` | OneNote에 새 페이지를 만들거나 Codex가 사용자의 메모/업무 기록을 작성할 때 사용한다. |
| SK-002 | 원노트 페이지 추가 | `원노트-페이지-추가.md` | 사용자가 OneNote 특정 섹션에 새 페이지를 만들고 제목/본문을 작성해달라고 할 때 사용한다. |
| SK-003 | 원노트 섹션 생성 | `원노트-섹션-생성.md` | 사용자가 특정 섹션 그룹 아래에 새 섹션을 만들어달라고 할 때 사용한다. |
| SK-004 | 원노트 섹션 그룹 생성 | `원노트-섹션그룹-생성.md` | 사용자가 전자필기장 또는 기존 섹션 그룹 아래에 새 섹션 그룹을 만들어달라고 할 때 사용한다. |
| SK-005 | 원노트 전자필기장 생성 | `원노트-전자필기장-생성.md` | 사용자가 새 OneNote 전자필기장을 만들거나 OneDrive/로컬 경로의 전자필기장을 열어달라고 할 때 사용한다. |
| SK-006 | 원노트 대상 ID 찾기 | `원노트-대상ID-찾기.md` | 사용자가 새 OneNote 대상 경로를 자주 쓰게 되었거나, 섹션/섹션그룹 ID를 캐시에 추가해야 할 때 사용한다. |
````

## 스킬 진단

````markdown
# 코덱스 & 원노트 스킬 진단

생성 시각: 2026-04-18 09:01:24
스킬 수: 6

## 문제

- 문제 없음

## 주문번호 목록

| 주문번호 | 스킬 이름 | 파일 |
| --- | --- | --- |
| SK-001 | 나의 기본 글쓰기 형식 | `나의-기본-글쓰기-형식.md` |
| SK-002 | 원노트 페이지 추가 | `원노트-페이지-추가.md` |
| SK-003 | 원노트 섹션 생성 | `원노트-섹션-생성.md` |
| SK-004 | 원노트 섹션 그룹 생성 | `원노트-섹션그룹-생성.md` |
| SK-005 | 원노트 전자필기장 생성 | `원노트-전자필기장-생성.md` |
| SK-006 | 원노트 대상 ID 찾기 | `원노트-대상ID-찾기.md` |

````

# 원노트 대상 ID 찾기

## 내부 용도

사용자가 새 OneNote 대상 경로를 자주 쓰게 되었거나, 섹션/섹션그룹 ID를 캐시에 추가해야 할 때 사용한다.

## 절차

1. 앱의 `코덱스` 탭에서 대상 이름, 전자필기장, 섹션 그룹, 섹션 이름을 입력한다.
2. `ID 찾기 스크립트 복사`를 눌러 PowerShell 스크립트를 가져온다.
3. 스크립트는 OneNote COM API의 `GetHierarchy('', hsSections, ref xml)`만 사용한다.
4. 스크립트는 출력 JSON을 클립보드에도 복사한다.
5. 앱의 `클립보드 JSON 저장`으로 대상 캐시에 반영한다.
6. `section_group_id`와 `section_id`가 저장됐는지 확인한다.
7. 다음 작업부터는 전체 계층 탐색보다 저장된 ID를 먼저 사용한다.

## PowerShell Pattern

```powershell
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$xml = ""
$one.GetHierarchy(
    "",
    [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections,
    [ref]$xml
)

[xml]$doc = $xml
$nsUri = $doc.DocumentElement.NamespaceURI
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("one", $nsUri)

# one:Notebook / one:SectionGroup / one:Section을 이름으로 찾아 ID를 추출한다.
$json = $result | ConvertTo-Json -Depth 4
$json | Set-Clipboard
$json
```

## 주의 기준

- ID 찾기는 `hsSections`까지만 조회한다. `hsPages`는 페이지가 많으면 느리다.
- 이름이 중복될 수 있으므로 전자필기장 이름과 섹션 그룹 이름을 같이 제한한다.
- 출력 JSON을 캐시에 저장한 뒤에는 실제 작업에서 ID 직접 호출을 우선한다.

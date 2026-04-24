# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


class MainWindowWindowsTemplatesAMixin:

    def _codex_onenote_templates_windows_a(
        self, values, title, body, target, section_group_id, section_id
    ) -> Dict[str, str]:
        return {
            "add_page": f"""# OneNote COM: 페이지 추가
# 대상: {values["target"]}
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$sectionId = {section_id}
$title = {title}
$body = {body}
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
if ($titleNode -ne $null) {{
    $titleNode.InnerText = $title
}}

$outline = $doc.CreateElement("one", "Outline", $nsUri)
$oeChildren = $doc.CreateElement("one", "OEChildren", $nsUri)
foreach ($line in ($body -split "`r?`n")) {{
    $oe = $doc.CreateElement("one", "OE", $nsUri)
    $t = $doc.CreateElement("one", "T", $nsUri)
    $t.InnerText = $line
    [void]$oe.AppendChild($t)
    [void]$oeChildren.AppendChild($oe)
}}
[void]$outline.AppendChild($oeChildren)
[void]$doc.DocumentElement.AppendChild($outline)

$one.UpdatePageContent($doc.OuterXml)
$verifyXml = ""
$one.GetPageContent($pageId, [ref]$verifyXml)
if ($verifyXml -notlike "*$title*") {{
    throw "페이지 생성 검증 실패: 제목을 찾지 못했습니다."
}}
$one.NavigateTo($pageId)
""",
            "add_section": f"""# OneNote COM: 새 섹션 생성
# 대상: {values["target"]}
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$sectionGroupId = {section_group_id}
$sectionName = {title}
$newSectionId = ""
$safeName = $sectionName -replace '[<>:"/\\\\|?*]', '-'
if (-not $safeName.EndsWith(".one")) {{
    $safeName = "$safeName.one"
}}

$one.OpenHierarchy(
    $safeName,
    $sectionGroupId,
    [ref]$newSectionId,
    [Microsoft.Office.Interop.OneNote.CreateFileType]::cftSection
)

$verifyXml = ""
$one.GetHierarchy($sectionGroupId, [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections, [ref]$verifyXml)
if ($verifyXml -notlike "*$sectionName*") {{
    throw "섹션 생성 검증 실패: 새 섹션 이름을 찾지 못했습니다."
}}
$one.NavigateTo($newSectionId)
""",
            "add_section_group": f"""# OneNote COM: 새 섹션 그룹 생성
# 대상: {values["target"]}
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$parentSectionGroupId = {section_group_id}
$sectionGroupName = {title}
$newSectionGroupId = ""

$one.OpenHierarchy(
    $sectionGroupName,
    $parentSectionGroupId,
    [ref]$newSectionGroupId,
    [Microsoft.Office.Interop.OneNote.CreateFileType]::cftFolder
)

$verifyXml = ""
$one.GetHierarchy($parentSectionGroupId, [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections, [ref]$verifyXml)
if ($verifyXml -notlike "*$sectionGroupName*") {{
    throw "섹션 그룹 생성 검증 실패: 새 섹션 그룹 이름을 찾지 못했습니다."
}}
$one.NavigateTo($newSectionGroupId)
""",
            "add_notebook": f"""# OneNote COM: 새 전자필기장 생성
# 주의: 새 전자필기장은 저장 위치 경로가 필요하다. OneDrive 동기화 폴더 또는 로컬 경로를 먼저 확정한다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$notebookPath = {target}
$newNotebookId = ""
if ([string]::IsNullOrWhiteSpace($notebookPath)) {{
    throw "새 전자필기장을 만들 파일 시스템 경로가 필요합니다."
}}

$one.OpenHierarchy(
    $notebookPath,
    "",
    [ref]$newNotebookId,
    [Microsoft.Office.Interop.OneNote.CreateFileType]::cftNotebook
)

$one.NavigateTo($newNotebookId)
""",
            "list_hierarchy": f"""# OneNote COM: 계층 구조 조회
# 열린 전자필기장, 섹션 그룹, 섹션 계층을 JSON으로 정리하고 클립보드에 복사합니다.
# 페이지 목록은 출력이 커질 수 있으므로 '섹션 페이지 목록 읽기' 스킬로 대상 섹션만 따로 조회합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$hierarchyXml = ""
$one.GetHierarchy(
    "",
    [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections,
    [ref]$hierarchyXml
)

if ([string]::IsNullOrWhiteSpace($hierarchyXml)) {{
    throw "계층 구조 조회 실패: OneNote hierarchy XML이 비어 있습니다."
}}

[xml]$doc = $hierarchyXml
$sections = @($doc.SelectNodes("//*[local-name()='Section']")) | ForEach-Object {{
    [ordered]@{{
        name = $_.GetAttribute("name")
        id = $_.GetAttribute("ID")
        path = $_.GetAttribute("path")
    }}
}}

$result = [ordered]@{{
    generated_at = (Get-Date).ToString("s")
    notebook_count = @($doc.SelectNodes("//*[local-name()='Notebook']")).Count
    section_group_count = @($doc.SelectNodes("//*[local-name()='SectionGroup']")).Count
    section_count = @($doc.SelectNodes("//*[local-name()='Section']")).Count
    sections = $sections
}}

$json = $result | ConvertTo-Json -Depth 10
$json | Set-Clipboard
$json
""",
            "read_section_pages": f"""# OneNote COM: 섹션 페이지 목록 읽기
# 대상 Section ID의 페이지 목록을 읽고, 제목 필터가 맞으면 대표 페이지 XML도 함께 가져옵니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$sectionId = {section_id}
$PageTitleContains = {title}
if ([string]::IsNullOrWhiteSpace($sectionId)) {{
    throw "섹션 페이지 목록 읽기 실패: Section ID가 비어 있습니다."
}}

$hierarchyXml = ""
$one.GetHierarchy(
    $sectionId,
    [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages,
    [ref]$hierarchyXml
)

[xml]$doc = $hierarchyXml
$nsUri = $doc.DocumentElement.NamespaceURI
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("one", $nsUri)

$pages = @($doc.SelectNodes("//one:Page", $ns)) | ForEach-Object {{
    [ordered]@{{
        title = $_.GetAttribute("name")
        id = $_.GetAttribute("ID")
        lastModifiedTime = $_.GetAttribute("lastModifiedTime")
    }}
}}

$selectedPage = $null
if (-not [string]::IsNullOrWhiteSpace($PageTitleContains)) {{
    $selectedPage = $pages |
        Where-Object {{ $_.title -like "*$PageTitleContains*" }} |
        Select-Object -First 1
}}

$pageContent = ""
if ($null -ne $selectedPage) {{
    $one.GetPageContent($selectedPage.id, [ref]$pageContent)
}}

$result = [ordered]@{{
    generated_at = (Get-Date).ToString("s")
    section_id = $sectionId
    page_count = @($pages).Count
    pages = $pages
    selected_page = $selectedPage
    selected_page_xml = $pageContent
}}

$json = $result | ConvertTo-Json -Depth 8
$json | Set-Clipboard
$json
""",
            "read_page_xml": f"""# OneNote COM: 페이지 XML 읽기
# 대상 입력값에는 Page ID를 넣습니다. 결과 XML은 클립보드에도 복사됩니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$pageId = {target}
if ([string]::IsNullOrWhiteSpace($pageId)) {{
    throw "페이지 XML 읽기 실패: Page ID가 비어 있습니다."
}}

$pageXml = ""
$one.GetPageContent($pageId, [ref]$pageXml)
if ([string]::IsNullOrWhiteSpace($pageXml)) {{
    throw "페이지 XML 읽기 검증 실패: 결과 XML이 비어 있습니다."
}}

$pageXml | Set-Clipboard
$pageXml
""",
            "append_page_body": f"""# OneNote COM: 페이지 본문 추가
# 대상 입력값에는 Page ID를 넣고, 본문 입력값을 페이지 끝에 새 Outline으로 추가합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$pageId = {target}
$body = {body}
if ([string]::IsNullOrWhiteSpace($pageId)) {{
    throw "페이지 본문 추가 실패: Page ID가 비어 있습니다."
}}
if ([string]::IsNullOrWhiteSpace($body)) {{
    throw "페이지 본문 추가 실패: 추가할 본문이 비어 있습니다."
}}

$pageXml = ""
$one.GetPageContent($pageId, [ref]$pageXml)
$beforeLength = $pageXml.Length

[xml]$doc = $pageXml
$nsUri = $doc.DocumentElement.NamespaceURI
$outline = $doc.CreateElement("one", "Outline", $nsUri)
$oeChildren = $doc.CreateElement("one", "OEChildren", $nsUri)
foreach ($line in ($body -split "`r?`n")) {{
    if ([string]::IsNullOrWhiteSpace($line)) {{ continue }}
    $oe = $doc.CreateElement("one", "OE", $nsUri)
    $t = $doc.CreateElement("one", "T", $nsUri)
    $t.InnerText = $line
    [void]$oe.AppendChild($t)
    [void]$oeChildren.AppendChild($oe)
}}
[void]$outline.AppendChild($oeChildren)
[void]$doc.DocumentElement.AppendChild($outline)

$one.UpdatePageContent($doc.OuterXml)
$verifyXml = ""
$one.GetPageContent($pageId, [ref]$verifyXml)
if ($verifyXml.Length -le $beforeLength) {{
    throw "페이지 본문 추가 검증 실패: 페이지 XML 길이가 증가하지 않았습니다."
}}
$one.NavigateTo($pageId)
""",
            "rename_page": f"""# OneNote COM: 페이지 제목 변경
# 대상 입력값에는 Page ID를 넣고, 제목 입력값을 새 페이지 제목으로 사용합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$pageId = {target}
$newTitle = {title}
if ([string]::IsNullOrWhiteSpace($pageId)) {{
    throw "페이지 제목 변경 실패: Page ID가 비어 있습니다."
}}
if ([string]::IsNullOrWhiteSpace($newTitle)) {{
    throw "페이지 제목 변경 실패: 새 제목이 비어 있습니다."
}}

$pageXml = ""
$one.GetPageContent($pageId, [ref]$pageXml)
[xml]$doc = $pageXml
$nsUri = $doc.DocumentElement.NamespaceURI
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("one", $nsUri)
$titleNode = $doc.SelectSingleNode("//one:Title//one:T", $ns)
if ($titleNode -eq $null) {{
    throw "페이지 제목 변경 실패: Title 노드를 찾지 못했습니다."
}}
$titleNode.InnerText = $newTitle
$one.UpdatePageContent($doc.OuterXml)

$verifyXml = ""
$one.GetPageContent($pageId, [ref]$verifyXml)
if ($verifyXml -notlike "*$newTitle*") {{
    throw "페이지 제목 변경 검증 실패: 새 제목을 찾지 못했습니다."
}}
$one.NavigateTo($pageId)
""",
            "open_notebook": f"""# OneNote COM: 전자필기장 열기
# 대상 입력값에는 기존 전자필기장 폴더/파일 경로를 넣습니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$notebookPath = {target}
$notebookId = ""
if ([string]::IsNullOrWhiteSpace($notebookPath)) {{
    throw "전자필기장 열기 실패: 전자필기장 경로가 비어 있습니다."
}}

$one.OpenHierarchy(
    $notebookPath,
    "",
    [ref]$notebookId,
    [Microsoft.Office.Interop.OneNote.CreateFileType]::cftNone
)

if ([string]::IsNullOrWhiteSpace($notebookId)) {{
    throw "전자필기장 열기 검증 실패: Notebook ID가 반환되지 않았습니다."
}}
$one.NavigateTo($notebookId)
""",
        }

_publish_context(globals())

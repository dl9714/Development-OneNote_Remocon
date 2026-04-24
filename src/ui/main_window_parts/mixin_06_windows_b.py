# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())


class MainWindowWindowsTemplatesBMixin:

    def _codex_onenote_templates_windows_b(
        self, values, title, body, target, section_group_id, section_id
    ) -> Dict[str, str]:
        return {
            "navigate_to_id": f"""# OneNote COM: ID로 이동
# 대상 입력값에는 Notebook/SectionGroup/Section/Page ID 중 하나를 넣습니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$targetId = {target}
if ([string]::IsNullOrWhiteSpace($targetId)) {{
    throw "ID로 이동 실패: 대상 ID가 비어 있습니다."
}}

$one.NavigateTo($targetId)
[ordered]@{{
    navigated_at = (Get-Date).ToString("s")
    target_id = $targetId
}} | ConvertTo-Json | Set-Clipboard
""",
            "find_pages": f"""# OneNote COM: 페이지 검색
# 제목 입력값을 검색어로 사용합니다. 대상 입력값이 OneNote ID이면 해당 범위에서, 아니면 전체 열린 전자필기장에서 검색합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$startNodeId = {target}
$query = {title}
if ($startNodeId -notmatch '^\\{{') {{
    $startNodeId = ""
}}
if ([string]::IsNullOrWhiteSpace($query)) {{
    throw "페이지 검색 실패: 검색어가 비어 있습니다."
}}

$resultXml = ""
$one.FindPages($startNodeId, $query, [ref]$resultXml, $true, $false)
if ([string]::IsNullOrWhiteSpace($resultXml)) {{
    throw "페이지 검색 검증 실패: 검색 결과 XML이 비어 있습니다."
}}

[xml]$doc = $resultXml
$nsUri = $doc.DocumentElement.NamespaceURI
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("one", $nsUri)

$allPages = @($doc.SelectNodes("//one:Page", $ns))
$pages = $allPages |
    Select-Object -First 50 |
    ForEach-Object {{
        [ordered]@{{
            title = $_.GetAttribute("name")
            id = $_.GetAttribute("ID")
            lastModifiedTime = $_.GetAttribute("lastModifiedTime")
        }}
    }}

$result = [ordered]@{{
    generated_at = (Get-Date).ToString("s")
    query = $query
    matched_page_count = @($allPages).Count
    returned_page_count = @($pages).Count
    truncated = (@($allPages).Count -gt @($pages).Count)
    pages = $pages
}}

$json = $result | ConvertTo-Json -Depth 8
$json | Set-Clipboard
$json
""",
            "get_object_link": f"""# OneNote COM: 링크 생성
# 대상 입력값에는 Notebook/SectionGroup/Section/Page ID를 넣습니다. 본문 입력값이 Content Object ID이면 해당 개체 링크를 생성합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$hierarchyId = {target}
$contentObjectId = {body}
if ([string]::IsNullOrWhiteSpace($hierarchyId) -or $hierarchyId -notmatch '^\\{{') {{
    throw "링크 생성 실패: 대상 OneNote ID가 비어 있거나 ID 형식이 아닙니다."
}}
if ($contentObjectId -notmatch '^\\{{') {{
    $contentObjectId = ""
}}

$link = ""
$one.GetHyperlinkToObject($hierarchyId, $contentObjectId, [ref]$link)
if ([string]::IsNullOrWhiteSpace($link)) {{
    throw "링크 생성 검증 실패: 링크가 반환되지 않았습니다."
}}

$link | Set-Clipboard
$link
""",
            "get_web_link": f"""# OneNote COM: 웹 링크 생성
# 대상 입력값에는 Notebook/SectionGroup/Section/Page ID를 넣습니다. 브라우저나 다른 앱에서 열 수 있는 웹 링크를 생성합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$hierarchyId = {target}
$contentObjectId = {body}
if ([string]::IsNullOrWhiteSpace($hierarchyId) -or $hierarchyId -notmatch '^\\{{') {{
    throw "웹 링크 생성 실패: 대상 OneNote ID가 비어 있거나 ID 형식이 아닙니다."
}}
if ($contentObjectId -notmatch '^\\{{') {{
    $contentObjectId = ""
}}

$link = ""
$one.GetWebHyperlinkToObject($hierarchyId, $contentObjectId, [ref]$link)
if ([string]::IsNullOrWhiteSpace($link)) {{
    throw "웹 링크 생성 검증 실패: 링크가 반환되지 않았습니다."
}}

$link | Set-Clipboard
$link
""",
            "get_parent_id": f"""# OneNote COM: 부모 ID 조회
# 대상 입력값에는 Page/Section/SectionGroup ID를 넣습니다. 부모 ID를 JSON과 클립보드로 반환합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$objectId = {target}
if ([string]::IsNullOrWhiteSpace($objectId) -or $objectId -notmatch '^\\{{') {{
    throw "부모 ID 조회 실패: 대상 OneNote ID가 비어 있거나 ID 형식이 아닙니다."
}}

$parentId = ""
$one.GetHierarchyParent($objectId, [ref]$parentId)
if ([string]::IsNullOrWhiteSpace($parentId)) {{
    throw "부모 ID 조회 검증 실패: 부모 ID가 반환되지 않았습니다."
}}

$result = [ordered]@{{
    generated_at = (Get-Date).ToString("s")
    object_id = $objectId
    parent_id = $parentId
}}
$json = $result | ConvertTo-Json
$json | Set-Clipboard
$json
""",
            "sync_hierarchy": f"""# OneNote COM: 계층 동기화
# 대상 입력값에는 Notebook/SectionGroup/Section/Page ID를 넣습니다. 동기화 후 hsSelf 조회로 접근 가능 여부를 검증합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$targetId = {target}
if ([string]::IsNullOrWhiteSpace($targetId) -or $targetId -notmatch '^\\{{') {{
    throw "계층 동기화 실패: 대상 OneNote ID가 비어 있거나 ID 형식이 아닙니다."
}}

$one.SyncHierarchy($targetId)
$verifyXml = ""
$one.GetHierarchy(
    $targetId,
    [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSelf,
    [ref]$verifyXml
)
if ([string]::IsNullOrWhiteSpace($verifyXml)) {{
    throw "계층 동기화 검증 실패: 대상 계층을 다시 조회하지 못했습니다."
}}

[ordered]@{{
    synced_at = (Get-Date).ToString("s")
    target_id = $targetId
    verify_xml_length = $verifyXml.Length
}} | ConvertTo-Json | Set-Clipboard
""",
            "export_pdf": f"""# OneNote COM: PDF 내보내기
# 대상 입력값에는 내보낼 OneNote ID를 넣습니다. 본문 입력값이 .pdf 경로이면 그 경로를 사용하고, 아니면 바탕화면에 저장합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$targetId = {target}
$requestedPath = {body}
$title = {title}
if ([string]::IsNullOrWhiteSpace($targetId) -or $targetId -notmatch '^\\{{') {{
    throw "PDF 내보내기 실패: 대상 OneNote ID가 비어 있거나 ID 형식이 아닙니다."
}}

if ([string]::IsNullOrWhiteSpace($requestedPath) -or -not $requestedPath.ToLower().EndsWith(".pdf")) {{
    $safeName = $title -replace '[<>:"/\\\\|?*]', '-'
    if ([string]::IsNullOrWhiteSpace($safeName)) {{ $safeName = "OneNote-Export" }}
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $requestedPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "$safeName-$stamp.pdf"
}}

$parentDir = Split-Path -Parent $requestedPath
if (-not [string]::IsNullOrWhiteSpace($parentDir)) {{
    New-Item -ItemType Directory -Force -Path $parentDir | Out-Null
}}

$one.Publish(
    $targetId,
    $requestedPath,
    [Microsoft.Office.Interop.OneNote.PublishFormat]::pfPDF
)

if (-not (Test-Path -LiteralPath $requestedPath)) {{
    throw "PDF 내보내기 검증 실패: 파일이 생성되지 않았습니다. $requestedPath"
}}

$requestedPath | Set-Clipboard
$requestedPath
""",
            "export_mhtml": f"""# OneNote COM: MHTML 내보내기
# 대상 입력값에는 내보낼 OneNote ID를 넣습니다. 본문 입력값이 .mht/.mhtml 경로이면 그 경로를 사용하고, 아니면 바탕화면에 저장합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$targetId = {target}
$requestedPath = {body}
$title = {title}
if ([string]::IsNullOrWhiteSpace($targetId) -or $targetId -notmatch '^\\{{') {{
    throw "MHTML 내보내기 실패: 대상 OneNote ID가 비어 있거나 ID 형식이 아닙니다."
}}

$lowerPath = $requestedPath.ToLower()
if ([string]::IsNullOrWhiteSpace($requestedPath) -or (-not $lowerPath.EndsWith(".mht") -and -not $lowerPath.EndsWith(".mhtml"))) {{
    $safeName = $title -replace '[<>:"/\\\\|?*]', '-'
    if ([string]::IsNullOrWhiteSpace($safeName)) {{ $safeName = "OneNote-Export" }}
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $requestedPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "$safeName-$stamp.mht"
}}

$parentDir = Split-Path -Parent $requestedPath
if (-not [string]::IsNullOrWhiteSpace($parentDir)) {{
    New-Item -ItemType Directory -Force -Path $parentDir | Out-Null
}}

$one.Publish(
    $targetId,
    $requestedPath,
    [Microsoft.Office.Interop.OneNote.PublishFormat]::pfMHTML
)

if (-not (Test-Path -LiteralPath $requestedPath)) {{
    throw "MHTML 내보내기 검증 실패: 파일이 생성되지 않았습니다. $requestedPath"
}}

$requestedPath | Set-Clipboard
$requestedPath
""",
            "export_xps": f"""# OneNote COM: XPS 내보내기
# 대상 입력값에는 내보낼 OneNote ID를 넣습니다. 본문 입력값이 .xps 경로이면 그 경로를 사용하고, 아니면 바탕화면에 저장합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$targetId = {target}
$requestedPath = {body}
$title = {title}
if ([string]::IsNullOrWhiteSpace($targetId) -or $targetId -notmatch '^\\{{') {{
    throw "XPS 내보내기 실패: 대상 OneNote ID가 비어 있거나 ID 형식이 아닙니다."
}}

if ([string]::IsNullOrWhiteSpace($requestedPath) -or -not $requestedPath.ToLower().EndsWith(".xps")) {{
    $safeName = $title -replace '[<>:"/\\\\|?*]', '-'
    if ([string]::IsNullOrWhiteSpace($safeName)) {{ $safeName = "OneNote-Export" }}
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $requestedPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "$safeName-$stamp.xps"
}}

$parentDir = Split-Path -Parent $requestedPath
if (-not [string]::IsNullOrWhiteSpace($parentDir)) {{
    New-Item -ItemType Directory -Force -Path $parentDir | Out-Null
}}

$one.Publish(
    $targetId,
    $requestedPath,
    [Microsoft.Office.Interop.OneNote.PublishFormat]::pfXPS
)

if (-not (Test-Path -LiteralPath $requestedPath)) {{
    throw "XPS 내보내기 검증 실패: 파일이 생성되지 않았습니다. $requestedPath"
}}

$requestedPath | Set-Clipboard
$requestedPath
""",
            "get_special_locations": f"""# OneNote COM: 특수 위치 조회
# 백업 폴더, 빠른 노트 섹션, 기본 전자필기장 폴더를 조회합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$backupFolder = ""
$unfiledNotesSection = ""
$defaultNotebookFolder = ""
$one.GetSpecialLocation([Microsoft.Office.Interop.OneNote.SpecialLocation]::slBackUpFolder, [ref]$backupFolder)
$one.GetSpecialLocation([Microsoft.Office.Interop.OneNote.SpecialLocation]::slUnfiledNotesSection, [ref]$unfiledNotesSection)
$one.GetSpecialLocation([Microsoft.Office.Interop.OneNote.SpecialLocation]::slDefaultNotebookFolder, [ref]$defaultNotebookFolder)

$result = [ordered]@{{
    generated_at = (Get-Date).ToString("s")
    backup_folder = $backupFolder
    unfiled_notes_section = $unfiledNotesSection
    default_notebook_folder = $defaultNotebookFolder
}}
$json = $result | ConvertTo-Json
$json | Set-Clipboard
$json
""",
            "close_notebook": f"""# OneNote COM: 전자필기장 닫기
# 대상 입력값에는 Notebook ID를 넣습니다. 닫기 전 Notebook ID인지 확인하고 강제 닫기는 사용하지 않습니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$notebookId = {target}
if ([string]::IsNullOrWhiteSpace($notebookId) -or $notebookId -notmatch '^\\{{') {{
    throw "전자필기장 닫기 실패: Notebook ID가 비어 있거나 ID 형식이 아닙니다."
}}

$selfXml = ""
$one.GetHierarchy(
    $notebookId,
    [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSelf,
    [ref]$selfXml
)
if ($selfXml -notmatch '<one:Notebook|<Notebook') {{
    throw "전자필기장 닫기 중단: 대상 ID가 Notebook이 아닌 것 같습니다."
}}

$one.CloseNotebook($notebookId, $false)
[ordered]@{{
    closed_at = (Get-Date).ToString("s")
    notebook_id = $notebookId
    force = $false
}} | ConvertTo-Json | Set-Clipboard
""",
            "navigate_to_url": f"""# OneNote COM: URL로 이동
# 대상 입력값에는 onenote:/https:/http: URL을 넣습니다. 외부 URL은 새 창에서 엽니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$url = {target}
if ([string]::IsNullOrWhiteSpace($url) -or $url -notmatch '^(onenote|https?|file):') {{
    throw "URL로 이동 실패: 대상 입력값이 URL 형식이 아닙니다."
}}

$one.NavigateToUrl($url, $true)
[ordered]@{{
    navigated_at = (Get-Date).ToString("s")
    url = $url
}} | ConvertTo-Json | Set-Clipboard
""",
        }

_publish_context(globals())

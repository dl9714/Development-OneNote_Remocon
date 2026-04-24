# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin04:

    def _codex_all_targets_copy_text(self) -> str:
        targets = self._load_codex_targets()
        lines = ["저장된 OneNote 작업 위치 목록"]
        if not targets:
            lines.append("- 저장된 작업 위치가 없습니다.")
            return "\n".join(lines)
        for idx, target in enumerate(targets, start=1):
            path = target.get("path") or target.get("name") or "경로 미지정"
            notebook = target.get("notebook") or "미지정"
            section_group = target.get("section_group") or "미지정"
            section = target.get("section") or "미지정"
            lines.extend(
                [
                    f"{idx}. {path}",
                    f"   - 전자필기장: {notebook}",
                    f"   - 섹션 그룹: {section_group}",
                    f"   - 섹션: {section}",
                ]
            )
        return "\n".join(lines)

    def _codex_targets_from_inventory_json_text(self, text: str) -> List[Dict[str, str]]:
        raw = (text or "").strip()
        if not raw:
            raise ValueError("클립보드에 OneNote 위치 조회 결과가 없습니다.")

        try:
            data = json.loads(raw)
        except Exception:
            starts = [idx for idx in (raw.find("{"), raw.find("[")) if idx >= 0]
            ends = [idx for idx in (raw.rfind("}"), raw.rfind("]")) if idx >= 0]
            if not starts or not ends:
                raise ValueError("위치 조회 결과의 시작과 끝을 찾지 못했습니다.")
            data = json.loads(raw[min(starts): max(ends) + 1])

        items = data.get("items") if isinstance(data, dict) else data
        if not isinstance(items, list):
            raise ValueError("위치 조회 결과에서 작업 위치 목록을 찾지 못했습니다.")

        section_group_ids: Dict[str, str] = {}
        for item in items:
            if not isinstance(item, dict):
                continue
            if str(item.get("type", "")).casefold() == "sectiongroup":
                path = str(item.get("path", "")).strip()
                item_id = str(item.get("id", "")).strip()
                if path and item_id:
                    section_group_ids[path] = item_id

        targets: List[Dict[str, str]] = []
        seen: Set[str] = set()
        for item in items:
            if not isinstance(item, dict):
                continue
            if str(item.get("type", "")).casefold() != "section":
                continue
            path = str(item.get("path", "")).strip()
            section_id = str(item.get("id", "")).strip()
            if not path or not section_id:
                continue
            parts = [part.strip() for part in path.split(">") if part.strip()]
            if not parts:
                continue
            notebook = parts[0]
            section = parts[-1]
            section_group = " > ".join(parts[1:-1])
            group_path = " > ".join(parts[:-1])
            dedupe_key = f"{path}|{section_id}"
            if dedupe_key in seen:
                continue
            seen.add(dedupe_key)
            targets.append(
                {
                    "name": f"조회 위치 - {section}",
                    "path": path,
                    "notebook": notebook,
                    "section_group": section_group,
                    "section": section,
                    "section_group_id": section_group_ids.get(group_path, ""),
                    "section_id": section_id,
                }
            )
        return targets

    def _codex_inventory_target_preview_text(self) -> str:
        targets = self._codex_targets_from_inventory_json_text(QApplication.clipboard().text())
        rows = [
            f"| {idx} | {target.get('name', '')} | {target.get('path', '')} | "
            f"{target.get('section_id', '')} |"
            for idx, target in enumerate(targets, start=1)
        ]
        if not rows:
            rows.append("| - | 후보 없음 | - | - |")
        return "\n".join(
            [
                "# OneNote 작업 위치 후보",
                "",
                f"후보 수: {len(targets)}",
                "",
                "| 번호 | 이름 | 경로 | Section ID |",
                "| ---: | --- | --- | --- |",
                *rows,
                "",
            ]
        )

    def _copy_codex_inventory_target_preview_to_clipboard(self) -> None:
        try:
            QApplication.clipboard().setText(self._codex_inventory_target_preview_text())
            try:
                self.connection_status_label.setText(
                    "OneNote 작업 위치 후보를 클립보드에 복사했습니다."
                )
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, "작업 위치 후보 생성 실패", str(e))

    def _import_codex_targets_from_clipboard_inventory(self) -> None:
        try:
            incoming = self._codex_targets_from_inventory_json_text(
                QApplication.clipboard().text()
            )
            if not incoming:
                QMessageBox.information(self, "작업 위치 후보 없음", "등록할 섹션 후보를 찾지 못했습니다.")
                return

            targets = self._load_codex_targets()
            existing_keys = {
                (target.get("path", ""), target.get("section_id", ""))
                for target in targets
                if isinstance(target, dict)
            }
            added = 0
            for target in incoming:
                key = (target.get("path", ""), target.get("section_id", ""))
                if key in existing_keys:
                    continue
                targets.append(target)
                existing_keys.add(key)
                added += 1

            self._write_codex_targets(targets)
            self._refresh_codex_target_combo()
            self._update_codex_status_summary()
            QMessageBox.information(
                self,
                "작업 위치 등록 완료",
                f"새 작업 위치 {added}개를 등록했습니다.\n전체 후보: {len(incoming)}개",
            )
        except Exception as e:
            QMessageBox.warning(self, "작업 위치 등록 실패", str(e))

    def _codex_onenote_inventory_script(self) -> str:
        return """# OneNote COM: 전체 구조 위치 조회 JSON
# 전자필기장/섹션그룹/섹션/페이지 목록을 평탄화해서 클립보드에 JSON으로 복사합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$xml = ""
$one.GetHierarchy(
    "",
    [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages,
    [ref]$xml
)

[xml]$doc = $xml
$nsUri = $doc.DocumentElement.NamespaceURI
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("one", $nsUri)

$items = New-Object System.Collections.Generic.List[object]

function Add-OneNoteNode {
    param(
        [System.Xml.XmlNode]$Node,
        [string[]]$PathParts
    )

    $local = $Node.LocalName
    if ($local -notin @("Notebook", "SectionGroup", "Section", "Page")) {
        return
    }

    $name = $Node.GetAttribute("name")
    if ([string]::IsNullOrWhiteSpace($name) -and $local -eq "Page") {
        $name = $Node.GetAttribute("name")
    }
    if ([string]::IsNullOrWhiteSpace($name)) {
        $name = "(이름 없음)"
    }

    $nextPath = @($PathParts + $name)
    $items.Add([ordered]@{
        type = $local
        name = $name
        id = $Node.GetAttribute("ID")
        path = [string]::Join(" > ", [string[]]$nextPath)
        isUnread = $Node.GetAttribute("isUnread")
        lastModifiedTime = $Node.GetAttribute("lastModifiedTime")
    })

    foreach ($child in $Node.ChildNodes) {
        if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            Add-OneNoteNode -Node $child -PathParts $nextPath
        }
    }
}

foreach ($notebook in @($doc.SelectNodes("//one:Notebook", $ns))) {
    Add-OneNoteNode -Node $notebook -PathParts @()
}

$summary = [ordered]@{
    generated_at = (Get-Date).ToString("s")
    total = $items.Count
    notebooks = @($items | Where-Object { $_.type -eq "Notebook" }).Count
    section_groups = @($items | Where-Object { $_.type -eq "SectionGroup" }).Count
    sections = @($items | Where-Object { $_.type -eq "Section" }).Count
    pages = @($items | Where-Object { $_.type -eq "Page" }).Count
    items = $items
}

$json = $summary | ConvertTo-Json -Depth 8
$json | Set-Clipboard
$json
"""

    def _copy_codex_onenote_inventory_script_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_onenote_location_request_text())
        try:
            self.connection_status_label.setText(
                "OneNote 위치 조회 요청을 한국어로 복사했습니다."
            )
        except Exception:
            pass

    def _codex_onenote_location_request_text(self) -> str:
        if IS_MACOS:
            return """작업:
OneNote for Mac에서 현재 열려 있는 전자필기장, 섹션 그룹, 섹션, 현재 보이는 페이지 위치를 조회해줘.

정리 방식:
- 왼쪽 패널 기준으로 전자필기장별 섹션 그룹과 섹션을 계층형으로 정리한다.
- 현재 선택된 섹션/페이지가 있으면 따로 표시한다.
- 사용자가 복사해서 작업 지시로 쓸 수 있게 경로를 `전자필기장 > 섹션 그룹 > 섹션` 형식으로 적는다.

보고:
- 조회한 전자필기장 수
- 섹션 그룹 수
- 섹션 수
- 현재 선택된 위치
- 바로 작업 위치로 지정할 만한 후보 목록
"""
        return """작업:
OneNote에서 현재 열려 있는 전자필기장, 섹션 그룹, 섹션을 조회해줘.

정리 방식:
- 전자필기장별로 섹션 그룹과 섹션을 계층형으로 정리한다.
- 작업 위치로 쓸 수 있는 섹션 후보를 따로 표시한다.
- 사용자가 복사해서 작업 지시로 쓸 수 있게 경로를 `전자필기장 > 섹션 그룹 > 섹션` 형식으로 적는다.

보고:
- 조회한 전자필기장 수
- 섹션 그룹 수
- 섹션 수
- 바로 작업 위치로 지정할 만한 후보 목록
"""

    def _codex_page_reader_script(self) -> str:
        profile = self._codex_target_from_fields()
        section_id = self._ps_single_quoted(profile.get("section_id", ""))
        section_path = self._ps_single_quoted(profile.get("path", ""))
        return f"""# OneNote COM: 현재 대상 섹션 페이지 읽기
# 현재 코덱스 작업 위치의 Section ID 기준으로 페이지 목록과 선택 페이지 내용을 조회합니다.
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote")
$one = New-Object -ComObject OneNote.Application

$sectionId = {section_id}
$sectionPath = {section_path}
$PageTitleContains = ""

if ([string]::IsNullOrWhiteSpace($sectionId)) {{
    throw "Section ID가 비어 있습니다. OneNote 조회 ON으로 세부 위치를 선택하거나 왼쪽 패널에서 섹션을 선택하세요. 대상: $sectionPath"
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
    section_path = $sectionPath
    page_count = @($pages).Count
    pages = $pages
    selected_page = $selectedPage
    selected_page_xml = $pageContent
}}

$json = $result | ConvertTo-Json -Depth 8
$json | Set-Clipboard
$json
"""

    def _copy_codex_page_reader_script_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._codex_page_reader_request_text())
        try:
            self.connection_status_label.setText(
                "현재 섹션의 페이지 읽기 요청을 한국어로 복사했습니다."
            )
        except Exception:
            pass

_publish_context(globals())

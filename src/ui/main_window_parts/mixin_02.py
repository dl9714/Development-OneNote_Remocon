# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())

class MainWindowMixin02:

    def _codex_onenote_location_lookup_script(self) -> str:
        return """# OneNote COM: 작업 위치 상세 조회
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

$targets = New-Object System.Collections.Generic.List[object]

function Join-TargetPath {
    param([string[]]$Parts)
    $clean = @($Parts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    return [string]::Join(" > ", [string[]]$clean)
}

function Add-Target {
    param(
        [string]$Kind,
        [string]$DisplayName,
        [string]$Path,
        [string]$Notebook,
        [string]$SectionGroup,
        [string]$Section,
        [string]$SectionGroupId,
        [string]$SectionId
    )
    $targets.Add([ordered]@{
        kind = $Kind
        name = $DisplayName
        path = $Path
        notebook = $Notebook
        section_group = $SectionGroup
        section = $Section
        section_group_id = $SectionGroupId
        section_id = $SectionId
    })
}

function Visit-OneNoteNode {
    param(
        [System.Xml.XmlNode]$Node,
        [string[]]$PathParts,
        [string]$NotebookName,
        [string]$SectionGroupPath,
        [string]$ParentContainerId
    )

    $local = $Node.LocalName
    if ($local -notin @("Notebook", "SectionGroup", "Section")) {
        return
    }

    $nodeName = $Node.GetAttribute("name")
    if ([string]::IsNullOrWhiteSpace($nodeName)) {
        $nodeName = "(이름 없음)"
    }
    $nodeId = $Node.GetAttribute("ID")

    if ($local -eq "Notebook") {
        $nextPathParts = @($nodeName)
        $path = Join-TargetPath -Parts $nextPathParts
        Add-Target `
            -Kind "notebook" `
            -DisplayName "전자필기장 - $nodeName" `
            -Path $path `
            -Notebook $nodeName `
            -SectionGroup "" `
            -Section "" `
            -SectionGroupId $nodeId `
            -SectionId ""

        foreach ($child in $Node.ChildNodes) {
            if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                Visit-OneNoteNode `
                    -Node $child `
                    -PathParts $nextPathParts `
                    -NotebookName $nodeName `
                    -SectionGroupPath "" `
                    -ParentContainerId $nodeId
            }
        }
        return
    }

    $nextPathParts = @($PathParts + $nodeName)
    $path = Join-TargetPath -Parts $nextPathParts

    if ($local -eq "SectionGroup") {
        $groupPath = if ([string]::IsNullOrWhiteSpace($SectionGroupPath)) {
            $nodeName
        } else {
            "$SectionGroupPath > $nodeName"
        }
        Add-Target `
            -Kind "section_group" `
            -DisplayName "그룹 - $nodeName" `
            -Path $path `
            -Notebook $NotebookName `
            -SectionGroup $groupPath `
            -Section "" `
            -SectionGroupId $nodeId `
            -SectionId ""

        foreach ($child in $Node.ChildNodes) {
            if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                Visit-OneNoteNode `
                    -Node $child `
                    -PathParts $nextPathParts `
                    -NotebookName $NotebookName `
                    -SectionGroupPath $groupPath `
                    -ParentContainerId $nodeId
            }
        }
        return
    }

    if ($local -eq "Section") {
        Add-Target `
            -Kind "section" `
            -DisplayName "섹션 - $nodeName" `
            -Path $path `
            -Notebook $NotebookName `
            -SectionGroup $SectionGroupPath `
            -Section $nodeName `
            -SectionGroupId $ParentContainerId `
            -SectionId $nodeId
    }
}

foreach ($notebook in @($doc.SelectNodes("//one:Notebook", $ns))) {
    Visit-OneNoteNode `
        -Node $notebook `
        -PathParts @() `
        -NotebookName "" `
        -SectionGroupPath "" `
        -ParentContainerId ""
}

$result = [ordered]@{
    generated_at = (Get-Date).ToString("s")
    count = $targets.Count
    targets = $targets
}

$result | ConvertTo-Json -Depth 8 -Compress
"""

    def _codex_location_lookup_targets_from_json_text(
        self, text: str
    ) -> List[Dict[str, str]]:
        data = self._codex_json_from_text(text)
        raw_targets = data.get("targets") if isinstance(data, dict) else data
        if not isinstance(raw_targets, list):
            raise ValueError("OneNote 위치 조회 결과에서 위치 목록을 찾지 못했습니다.")

        targets: List[Dict[str, str]] = []
        for item in raw_targets:
            if not isinstance(item, dict):
                continue
            kind = str(item.get("kind", "") or "").strip()
            path = str(item.get("path", "") or "").strip()
            notebook = str(item.get("notebook", "") or "").strip()
            section_group = str(item.get("section_group", "") or "").strip()
            section = str(item.get("section", "") or "").strip()
            section_group_id = str(item.get("section_group_id", "") or "").strip()
            section_id = str(item.get("section_id", "") or "").strip()
            name = str(item.get("name", "") or "").strip()
            if not path:
                continue
            targets.append(
                {
                    "kind": kind,
                    "name": name or path,
                    "path": path,
                    "notebook": notebook,
                    "section_group": section_group,
                    "section": section,
                    "section_group_id": section_group_id,
                    "section_id": section_id,
                }
            )
        return targets

    def _codex_location_lookup_label(self, profile: Dict[str, str]) -> str:
        kind = profile.get("kind", "")
        kind_label = {
            "notebook": "전자필기장",
            "section_group": "그룹",
            "section": "섹션",
        }.get(kind, "위치")
        return f"{kind_label} | {profile.get('path', '')}"

    def _codex_location_cache_path(self) -> str:
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return os.path.join(root, "docs", "codex", "onenote-location-cache.json")

    def _codex_location_targets_from_saved_targets(self) -> List[Dict[str, str]]:
        targets: List[Dict[str, str]] = []
        seen: Set[str] = set()
        for item in self._load_codex_targets():
            if not isinstance(item, dict):
                continue
            path = str(item.get("path", "") or "").strip()
            notebook = str(item.get("notebook", "") or "").strip()
            section_group = str(item.get("section_group", "") or "").strip()
            section = str(item.get("section", "") or "").strip()
            if not path:
                continue
            entries = [
                {
                    "kind": "notebook",
                    "name": notebook,
                    "path": notebook,
                    "notebook": notebook,
                    "section_group": "",
                    "section": "",
                    "section_group_id": "",
                    "section_id": "",
                },
                {
                    "kind": "section_group",
                    "name": section_group,
                    "path": " > ".join(part for part in [notebook, section_group] if part),
                    "notebook": notebook,
                    "section_group": section_group,
                    "section": "",
                    "section_group_id": str(item.get("section_group_id", "") or ""),
                    "section_id": "",
                },
                {
                    "kind": "section",
                    "name": section or path,
                    "path": path,
                    "notebook": notebook,
                    "section_group": section_group,
                    "section": section,
                    "section_group_id": str(item.get("section_group_id", "") or ""),
                    "section_id": str(item.get("section_id", "") or ""),
                },
            ]
            for profile in entries:
                if not profile.get("path"):
                    continue
                key = "|".join(
                    [
                        profile.get("kind", ""),
                        profile.get("path", ""),
                        profile.get("section_id", ""),
                    ]
                )
                if key in seen:
                    continue
                seen.add(key)
                targets.append(profile)
        return targets

    def _load_codex_location_lookup_cache(self) -> List[Dict[str, str]]:
        path = self._codex_location_cache_path()
        try:
            with open(path, "r", encoding="utf-8") as f:
                targets = self._codex_location_lookup_targets_from_json_text(f.read())
            if targets:
                return targets
        except Exception:
            pass
        return self._codex_location_targets_from_saved_targets()

    def _save_codex_location_lookup_cache(self, targets: List[Dict[str, str]]) -> None:
        payload = {
            "version": 1,
            "cached_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "targets": targets,
        }
        self._write_json_file_atomic(self._codex_location_cache_path(), payload)

    def _load_codex_location_lookup_cache_into_ui(self, selected_path: str = "") -> bool:
        targets = self._load_codex_location_lookup_cache()
        if not targets:
            return False
        self._codex_location_lookup_targets = targets
        self._populate_codex_location_lookup_combo(selected_path)
        return True

    def _codex_location_first_profile(
        self, *, kind: str = "", notebook: str = "", section_group: Optional[str] = None
    ) -> Optional[Dict[str, str]]:
        for profile in getattr(self, "_codex_location_lookup_targets", []):
            if not isinstance(profile, dict):
                continue
            if kind and profile.get("kind") != kind:
                continue
            if notebook and profile.get("notebook") != notebook:
                continue
            if section_group is not None and profile.get("section_group", "") != section_group:
                continue
            return profile
        return None

    def _codex_location_selected_notebook(self) -> str:
        combo = getattr(self, "codex_location_notebook_combo", None)
        if combo is None:
            return ""
        return str(combo.currentData() or "").strip()

    def _codex_location_selected_group(self) -> str:
        combo = getattr(self, "codex_location_group_combo", None)
        if combo is None:
            return ""
        data = combo.currentData()
        if isinstance(data, dict):
            return str(data.get("section_group", "") or "").strip()
        return ""

    def _configure_codex_lookup_combo(self, combo: QComboBox) -> None:
        combo.setMinimumWidth(0)
        combo.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        try:
            combo.setSizeAdjustPolicy(
                QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon
            )
            combo.setMinimumContentsLength(8)
            combo.setMaxVisibleItems(16)
        except Exception:
            pass

_publish_context(globals())

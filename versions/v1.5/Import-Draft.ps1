<#
.SYNOPSIS
Extracts a neutral, reviewable draft graph; never maps labels to Azure services.
.DESCRIPTION
Accepts UTF-8 .excalidraw or .json files (10 MiB, 2000 records maximum).
Images, ASCII, and other flowcharts must be interpreted by a person/agent first.
No OCR, network requests, content execution, or geometric endpoint inference.

Minimal manually interpreted JSON:
{"schemaVersion":1,"sourceFormat":"manual-ascii","components":[
  {"id":"a","label":"Client"},{"id":"b","label":"API"}],
 "relationships":[{"id":"ab","source":"a","target":"b","direction":"forward"}]}

schemaVersion=1 output contains sourceFormat, status, components, relationships,
annotations, groups, embeddedFiles, unresolved, warnings, and counts.
IDs remain case-sensitive source IDs; component safeId is a stable model-safe ID.
Components: id, safeId, label, sourceIds, kind (shape/group/image/sketch/unknown),
sourceType, geometry, labelParts, frameId, groupIds, link, fileId, requiresVisualReview.
Geometry preserves x/y/width/height/angle and relative points; absent values are null.
Label parts preserve text, originalText, source IDs, geometry, and explicit grouping.
Relationships: id, source, target, label, direction, sourceIds, confidence,
geometry, labelParts, provenance, dataFlowVerified=false. Excalidraw bindings and
arrowheads are retained. Start-only arrowheads swap source/target and use reverse.
Direction is forward/reverse/bidirectional/none/unknown. None/bidirectional/unknown
retain diagram start/end ordering, NOT a claim of data flow. Missing endpoints
are null and reported in unresolved, never glued to a nearby shape.
Annotations preserve standalone or ambiguously bound text. Frames are components
of kind group; groups list explicit groupIds membership only, never containment.
Warnings/unresolved entries have code, message, sourceIds. An empty component
list returns status=empty plus NoComponents diagnostics, not a successful topology.

Normalized input requires schemaVersion=1, sourceFormat (normalized, manual-image,
manual-ascii, manual-flowchart, or excalidraw), and components. Optional arrays:
relationships, annotations, groups, embeddedFiles, warnings, unresolved. Components/annotations
require id and label; relationships require id and direction, with nullable
source/target and optional label. sourceIds defaults to [id]; kind defaults to
shape. For normalized forward/reverse relationships, source and target ALREADY
mean arrow-origin and arrow-destination, exactly as in this importer's output.
reverse describes the recorded head relative to original diagram start/end
geometry; normalized import never swaps the endpoints again. A model adapter
connecting source to target must render its arrow forward, not reverse it twice.
Optional geometry, frameId, groupIds, labelParts, sourceType, link, fileId,
requiresVisualReview are retained. Connector binding/arrowhead evidence and embedded
file metadata also survive normalized re-import. No service selection or verified-flow
claims are imported. Original source files are never modified.
.EXAMPLE
.\Import-Draft.ps1 -InputPath C:\Diagrams\draft.excalidraw
.EXAMPLE
.\Import-Draft.ps1 -InputPath C:\Diagrams\interpreted.json -OutputPath C:\Diagrams\graph.json
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$InputPath,
    [string]$OutputPath
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
$maxBytes = 10 * 1024 * 1024
$maxRecords = 2000
$warnings = [System.Collections.Generic.List[object]]::new()
$unresolved = [System.Collections.Generic.List[object]]::new()
$components = [System.Collections.Generic.List[object]]::new()
$relationships = [System.Collections.Generic.List[object]]::new()
$annotations = [System.Collections.Generic.List[object]]::new()
$groups = [System.Collections.Generic.List[object]]::new()
$embeddedFiles = [System.Collections.Generic.List[object]]::new()

function Get-DraftValue($Object, [string]$Name, $Default = $null) {
    if ($null -eq $Object) { return $Default }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $Default }
    return ,$property.Value
}

function Test-DraftObject($Value) {
    # PS5.1 also reports JSON-wrapped arrays/scalars as "-is [pscustomobject]".
    return ($null -ne $Value -and $Value.GetType().FullName -ceq 'System.Management.Automation.PSCustomObject')
}

function Get-DraftPath([string]$Path) {
    if ($Path -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+(?:\\|$))' -or
        $Path -match '^\\\\[?.]\\' -or $Path.Substring(2).Contains(':')) {
        throw 'Use a fully qualified Windows filesystem path without device prefixes or alternate streams.'
    }
    return [IO.Path]::GetFullPath($Path)
}

function Get-DraftArray($Object, [string]$Name, [bool]$Required = $false) {
    $value = Get-DraftValue $Object $Name
    if ($null -eq $value) {
        if ($Required) { throw "Missing required array '$Name'." }
        return ,@()
    }
    if ($value -isnot [array]) { throw "'$Name' must be a JSON array." }
    return ,$value
}

function Get-DraftString($Object, [string]$Name, [bool]$Required = $false, $Default = $null) {
    $value = Get-DraftValue $Object $Name
    if ($null -eq $value) {
        if ($Required) { throw "Missing required string '$Name'." }
        return $Default
    }
    if ($value -isnot [string]) { throw "'$Name' must be a string." }
    if ($Required -and $Name -ne 'label' -and [string]::IsNullOrWhiteSpace($value)) {
        throw "'$Name' must not be empty."
    }
    return $value
}

function Get-DraftStrings($Object, [string]$Name) {
    $items = Get-DraftArray $Object $Name
    foreach ($item in $items) {
        if ($item -isnot [string] -or [string]::IsNullOrWhiteSpace($item)) {
            throw "'$Name' must contain nonempty strings."
        }
    }
    return ,$items
}

function Get-DraftBool($Object, [string]$Name) {
    $value = Get-DraftValue $Object $Name
    if ($null -eq $value) { return $false }
    if ($value -isnot [bool]) { throw "'$Name' must be a boolean." }
    return $value
}

function Test-DraftNumber($Value) {
    return ($null -ne $Value -and $Value -is [ValueType] -and
        $Value -isnot [bool] -and $Value -isnot [datetime] -and
        -not [double]::IsNaN([double]$Value) -and -not [double]::IsInfinity([double]$Value))
}

function Add-DraftDiagnostic($List, [string]$Code, [string]$Message, [object[]]$Ids = @()) {
    $List.Add([pscustomobject][ordered]@{code=$Code; message=$Message; sourceIds=@($Ids)})
}

function Get-DraftGeometry($Record, [string]$Id, [bool]$Nested = $false) {
    $inputGeometry = $Record
    if ($Nested) {
        $inputGeometry = Get-DraftValue $Record 'geometry'
        if ($null -eq $inputGeometry) { return $null }
        if (-not (Test-DraftObject $inputGeometry)) { throw "'geometry' must be an object or null." }
    }
    $geometry = [ordered]@{}
    $missing = $false
    foreach ($name in @('x','y','width','height','angle')) {
        $value = Get-DraftValue $inputGeometry $name
        if ($null -ne $value -and -not (Test-DraftNumber $value)) {
            throw "Geometry '$name' for '$Id' must be a finite number."
        }
        $geometry[$name] = $value
        if ($name -ne 'angle' -and $null -eq $value) { $missing = $true }
    }
    $points = Get-DraftValue $inputGeometry 'points'
    if ($null -ne $points) {
        if ($points -isnot [array]) { throw "Geometry points for '$Id' must be an array." }
        foreach ($point in $points) {
            if ($point -isnot [array] -or $point.Count -ne 2 -or
                -not (Test-DraftNumber $point[0]) -or -not (Test-DraftNumber $point[1])) {
                throw "Geometry points for '$Id' must contain numeric [x,y] pairs."
            }
        }
    }
    $geometry['points'] = $points
    if ($missing) {
        Add-DraftDiagnostic $warnings 'IncompleteGeometry' "Record '$Id' has incomplete geometry; no position was invented." @($Id)
    }
    if ($null -ne $geometry.angle -and [math]::Abs([double]$geometry.angle) -gt 0.000001) {
        Add-DraftDiagnostic $unresolved 'RotatedElementRequiresReview' "Record '$Id' is rotated; geometry is preserved without approximation." @($Id)
    }
    return [pscustomobject]$geometry
}

function New-DraftSafeId([string]$Id) {
    $slug = [regex]::Replace($Id.ToLowerInvariant(), '[^a-z0-9_]', '_')
    if ($slug.Length -gt 32) { $slug = $slug.Substring(0,32) }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Id))).Replace('-','').Substring(0,16).ToLowerInvariant() }
    finally { $sha.Dispose() }
    return "c_${slug}_$hash"
}

function New-DraftMap {
    return ,([System.Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal))
}

function Add-DraftRecord($Map, $Record) {
    if (-not (Test-DraftObject $Record)) { throw 'Every record must be a JSON object.' }
    $id = Get-DraftString $Record 'id' $true
    if ($Map.ContainsKey($id)) { throw "Duplicate record ID '$id'." }
    $Map.Add($id, $Record)
    return $id
}

function New-DraftLabelPart($Record, $Geometry) {
    $text = Get-DraftString $Record 'text' $false ''
    $original = Get-DraftString $Record 'originalText'
    $label = $text
    if ($null -ne $original) { $label = $original }
    return [pscustomobject][ordered]@{
        sourceId=$Record.id; text=$text; originalText=$original; label=$label
        geometry=$Geometry; frameId=(Get-DraftString $Record 'frameId')
        groupIds=(Get-DraftStrings $Record 'groupIds')
    }
}

function Get-DraftLabelParts($Record) {
    $parts = Get-DraftArray $Record 'labelParts'
    $result = [System.Collections.Generic.List[object]]::new()
    foreach ($part in $parts) {
        if (-not (Test-DraftObject $part)) { throw "'labelParts' entries must be objects." }
        $id = Get-DraftString $part 'sourceId' $true
        $result.Add([pscustomobject][ordered]@{
            sourceId=$id; text=(Get-DraftString $part 'text' $false '')
            originalText=(Get-DraftString $part 'originalText')
            label=(Get-DraftString $part 'label' $true)
            geometry=(Get-DraftGeometry $part $id $true)
            frameId=(Get-DraftString $part 'frameId')
            groupIds=(Get-DraftStrings $part 'groupIds')
        })
    }
    return ,$result.ToArray()
}

function Get-DraftSourceIds($Record, [string]$Id) {
    $ids = Get-DraftStrings $Record 'sourceIds'
    if ($ids.Count -eq 0) { return ,@($Id) }
    return ,$ids
}

function Resolve-DraftEndpoint($Binding, [string]$EdgeId, [string]$End, $ComponentMap) {
    if ($null -eq $Binding) {
        Add-DraftDiagnostic $unresolved 'UnboundEndpoint' "Connector '$EdgeId' has no explicit $End binding; no nearby shape was selected." @($EdgeId)
        return $null
    }
    if (-not (Test-DraftObject $Binding)) { throw "'$End' binding on '$EdgeId' must be an object or null." }
    $id = Get-DraftString $Binding 'elementId' $true
    if (-not $ComponentMap.ContainsKey($id)) {
        Add-DraftDiagnostic $unresolved 'UnresolvedEndpoint' "Connector '$EdgeId' refers to missing, deleted, or non-component '$id' at $End." @($EdgeId,$id)
        return $null
    }
    return $id
}

function Get-DraftConfidence($Source, $Target, [bool]$Manual = $false) {
    if ($null -ne $Source -and $null -ne $Target) {
        if ($Manual) { return 'declared' }
        return 'explicit-binding'
    }
    if ($null -ne $Source -or $null -ne $Target) { return 'partial-binding' }
    return 'unresolved'
}

function Import-DraftDiagnostics($Root, [string]$Name, $List) {
    foreach ($entry in (Get-DraftArray $Root $Name)) {
        if ($entry -is [string]) {
            Add-DraftDiagnostic $List 'ImportedNote' $entry
        } elseif (Test-DraftObject $entry) {
            Add-DraftDiagnostic $List (Get-DraftString $entry 'code' $true) (Get-DraftString $entry 'message' $true) (Get-DraftStrings $entry 'sourceIds')
        } else { throw "'$Name' must contain strings or diagnostic objects." }
    }
}

$inputFile = Get-DraftPath $InputPath
$outputFile = $null
if ($OutputPath) {
    $outputFile = Get-DraftPath $OutputPath
    if (Test-Path -LiteralPath $outputFile) { throw 'OutputPath already exists. Draft import never overwrites files.' }
    if (-not [IO.Directory]::Exists([IO.Path]::GetDirectoryName($outputFile))) {
        throw 'OutputPath parent directory must already exist.'
    }
}
if ([IO.Path]::GetExtension($inputFile).ToLowerInvariant() -notin @('.json','.excalidraw')) {
    throw 'Unsupported input format. Interpret images, ASCII, or other flowcharts visually and supply a schemaVersion=1 normalized JSON graph; this script does not perform OCR.'
}
if (-not [IO.File]::Exists($inputFile)) { throw 'InputPath must name an existing JSON or Excalidraw file.' }
$stream = [IO.File]::Open($inputFile, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
try {
    if ($stream.Length -gt $maxBytes) { throw 'Draft input exceeds the 10 MiB file size limit.' }
    $reader = [IO.StreamReader]::new($stream, [Text.UTF8Encoding]::new($false, $true), $true)
    try { $json = $reader.ReadToEnd() }
    catch [System.Text.DecoderFallbackException] { throw 'Draft input cannot be read as valid Unicode JSON text.' }
    finally { $reader.Dispose() }
} finally { $stream.Dispose() }
try { $root = ConvertFrom-Json -InputObject $json -ErrorAction Stop }
catch [System.ArgumentException] { throw "Invalid JSON in draft input: $($_.Exception.Message)" }
catch [System.InvalidOperationException] {
    if ($_.FullyQualifiedErrorId -ne 'DuplicateKeysInJsonString,Microsoft.PowerShell.Commands.ConvertFromJsonCommand') { throw }
    throw "Invalid JSON in draft input: $($_.Exception.Message)"
}
if (-not (Test-DraftObject $root)) { throw 'Draft input must be a JSON object, not an array, scalar, or null.' }

$recordMap = New-DraftMap
$componentMap = New-DraftMap
$targetMap = New-DraftMap
$geometryMap = New-DraftMap
$deletedRecords = 0
$attachedTexts = 0
$inputRecords = 0
$normalized = $null -ne (Get-DraftValue $root 'schemaVersion')
if (-not $normalized) {
    if ((Get-DraftString $root 'type') -cne 'excalidraw') {
        throw 'Unsupported JSON structure. Expected type=excalidraw or schemaVersion=1 normalized graph.'
    }
    $sourceFormat = 'excalidraw'
    $records = Get-DraftArray $root 'elements' $true
    $inputRecords = $records.Count
    if ($inputRecords -gt $maxRecords) { throw 'Draft input exceeds the 2000 record limit (including deleted elements).' }
    $active = [System.Collections.Generic.List[object]]::new()
    foreach ($record in $records) {
        $id = Add-DraftRecord $recordMap $record
        if (Get-DraftBool $record 'isDeleted') { $deletedRecords++; continue }
        [void](Get-DraftString $record 'type' $true)
        $active.Add($record)
        $geometryMap.Add($id, (Get-DraftGeometry $record $id))
    }
    foreach ($record in $active) {
        $id = $record.id
        $type = $record.type
        if ($type -ceq 'text') { continue }
        $frameId = Get-DraftString $record 'frameId'
        $groupIds = Get-DraftStrings $record 'groupIds'
        if ($type -cin @('arrow','line')) {
            $startHead = Get-DraftString $record 'startArrowhead'
            $endHead = Get-DraftString $record 'endArrowhead'
            $direction = 'none'
            if ($null -ne $startHead -and $startHead -ne '' -and $null -ne $endHead -and $endHead -ne '') { $direction = 'bidirectional' }
            elseif ($null -ne $startHead -and $startHead -ne '') { $direction = 'reverse' }
            elseif ($null -ne $endHead -and $endHead -ne '') { $direction = 'forward' }
            if ($null -eq $record.PSObject.Properties['startArrowhead'] -or $null -eq $record.PSObject.Properties['endArrowhead']) {
                if ($direction -eq 'none') {
                    $direction = 'unknown'
                    Add-DraftDiagnostic $unresolved 'MissingArrowheads' "Connector '$id' omits arrowhead metadata; direction was not assumed." @($id)
                } else {
                    Add-DraftDiagnostic $warnings 'IncompleteArrowheads' "Connector '$id' omits some arrowhead metadata; direction reflects only its recorded head(s)." @($id)
                }
            }
            $target = [pscustomobject][ordered]@{
                id=$id; source=$null; target=$null; label=''; direction=$direction
                sourceIds=@($id); confidence='unresolved'; geometry=$geometryMap[$id]
                labelParts=@(); frameId=$frameId; groupIds=@($groupIds)
                sourceType=$type; startBinding=(Get-DraftValue $record 'startBinding')
                endBinding=(Get-DraftValue $record 'endBinding')
                startArrowhead=$startHead; endArrowhead=$endHead
                provenance='source-diagram'; dataFlowVerified=$false
                link=(Get-DraftString $record 'link')
            }
            $relationships.Add($target)
        } else {
            $kind = 'unknown'
            switch -CaseSensitive ($type) {
                { $_ -cin @('rectangle','ellipse','diamond') } { $kind = 'shape' }
                { $_ -cin @('frame','magicframe') } { $kind = 'group' }
                'image' { $kind = 'image' }
                'freedraw' { $kind = 'sketch' }
            }
            $review = $kind -in @('image','sketch','unknown')
            if ($review) {
                Add-DraftDiagnostic $unresolved 'VisualReviewRequired' "Record '$id' of type '$type' needs visual review; no embedded content was decoded or fetched." @($id)
            }
            $label = Get-DraftString $record 'name' $false ''
            $target = [pscustomobject][ordered]@{
                id=$id; safeId=(New-DraftSafeId $id); label=$label; sourceIds=@($id)
                geometry=$geometryMap[$id]; kind=$kind; sourceType=$type; labelParts=@()
                frameId=$frameId; groupIds=@($groupIds); requiresVisualReview=$review
                fileId=(Get-DraftString $record 'fileId'); link=(Get-DraftString $record 'link')
            }
            $components.Add($target)
            $componentMap.Add($id,$target)
        }
        $targetMap.Add($id,$target)
    }

    # Collect BOTH directions of text ownership; conflicting evidence stays unresolved.
    $owners = New-DraftMap
    foreach ($record in $active) {
        if ($record.type -ceq 'text') {
            $containerId = Get-DraftString $record 'containerId'
            if (-not [string]::IsNullOrEmpty($containerId)) {
                $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
                [void]$set.Add($containerId)
                $owners.Add($record.id,$set)
            }
        }
    }
    foreach ($record in $active) {
        foreach ($bound in (Get-DraftArray $record 'boundElements')) {
            if (-not (Test-DraftObject $bound)) { throw "'boundElements' entries must be objects." }
            $boundId = Get-DraftString $bound 'id' $true
            $boundType = Get-DraftString $bound 'type' $true
            if (-not $recordMap.ContainsKey($boundId) -or (Get-DraftBool $recordMap[$boundId] 'isDeleted')) {
                Add-DraftDiagnostic $unresolved 'MissingBoundElement' "Record '$($record.id)' refers to missing or deleted bound record '$boundId'." @($record.id,$boundId)
                continue
            }
            if ($boundType -cne $recordMap[$boundId].type) {
                Add-DraftDiagnostic $unresolved 'BoundElementTypeMismatch' "Record '$boundId' does not match its declared binding type." @($record.id,$boundId)
                continue
            }
            if ($boundType -ceq 'text') {
                if (-not $owners.ContainsKey($boundId)) {
                    $owners.Add($boundId, [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal))
                }
                [void]$owners[$boundId].Add($record.id)
            }
        }
    }
    foreach ($record in $active) {
        if ($record.type -cne 'text') { continue }
        $part = New-DraftLabelPart $record $geometryMap[$record.id]
        $candidates = @()
        if ($owners.ContainsKey($record.id)) { $candidates = @($owners[$record.id]) }
        $owner = $null
        if ($candidates.Count -eq 1 -and $targetMap.ContainsKey($candidates[0])) { $owner = $targetMap[$candidates[0]] }
        if ($null -ne $owner) {
            if ($owner.label.Length -eq 0) { $owner.label = $part.label }
            else { $owner.label += "`n" + $part.label }
            $owner.labelParts += $part
            $owner.sourceIds += $record.id
            $attachedTexts++
        } else {
            if ($candidates.Count -gt 0) {
                Add-DraftDiagnostic $unresolved 'UnresolvedTextBinding' "Text '$($record.id)' has missing, invalid, or conflicting owners and remains an annotation." (@($record.id) + $candidates)
            }
            $annotations.Add([pscustomobject][ordered]@{
                id=$record.id; label=$part.label; text=$part.text; originalText=$part.originalText
                sourceIds=@($record.id); geometry=$part.geometry; kind='label'
                frameId=$part.frameId; groupIds=@($part.groupIds)
                containerId=(Get-DraftString $record 'containerId')
                link=(Get-DraftString $record 'link'); labelParts=@()
            })
        }
    }
    foreach ($edge in $relationships) {
        $start = Resolve-DraftEndpoint $edge.startBinding $edge.id 'start' $componentMap
        $end = Resolve-DraftEndpoint $edge.endBinding $edge.id 'end' $componentMap
        if ($edge.direction -eq 'reverse') { $edge.source = $end; $edge.target = $start }
        else { $edge.source = $start; $edge.target = $end }
        $edge.confidence = Get-DraftConfidence $edge.source $edge.target
    }
    $files = Get-DraftValue $root 'files'
    if ($null -ne $files) {
        if (-not (Test-DraftObject $files)) { throw "'files' must be an object." }
        $fileProperties = @($files.PSObject.Properties)
        if ($fileProperties.Count -gt $maxRecords) { throw 'Draft input exceeds the 2000 embedded file limit.' }
        foreach ($file in $fileProperties) {
            if (-not (Test-DraftObject $file.Value)) { throw 'Embedded file metadata must be an object.' }
            $embeddedFiles.Add([pscustomobject][ordered]@{
                id=$file.Name; mimeType=(Get-DraftString $file.Value 'mimeType')
                hasEmbeddedData=($null -ne (Get-DraftValue $file.Value 'dataURL'))
                requiresVisualReview=$true
            })
        }
        if ($embeddedFiles.Count -gt 0) {
            Add-DraftDiagnostic $warnings 'EmbeddedPayloadNotDecoded' 'Embedded file metadata was retained; image/SVG/HTML payloads were deliberately excluded, not decoded or executed.'
        }
    }
    # Group membership comes from explicit source metadata, not overlapping geometry.
    $groupMap = New-DraftMap
    foreach ($record in $active) {
        foreach ($groupId in (Get-DraftStrings $record 'groupIds')) {
            if (-not $groupMap.ContainsKey($groupId)) {
                $group = [pscustomobject][ordered]@{id=$groupId; memberIds=@(); provenance='explicit-groupIds'}
                $groupMap.Add($groupId,$group)
                $groups.Add($group)
            }
            if (-not (@($groupMap[$groupId].memberIds) -ccontains $record.id)) { $groupMap[$groupId].memberIds += $record.id }
        }
    }
} else {
    $version = Get-DraftValue $root 'schemaVersion'
    if (-not (Test-DraftNumber $version) -or $version -ne 1) { throw 'Unsupported schemaVersion. Expected numeric 1.' }
    $sourceFormat = Get-DraftString $root 'sourceFormat' $true
    if ($sourceFormat -cnotin @('normalized','manual-image','manual-ascii','manual-flowchart','excalidraw')) {
        throw 'Unsupported sourceFormat for normalized graph.'
    }
    $nodeRecords = Get-DraftArray $root 'components' $true
    $edgeRecords = Get-DraftArray $root 'relationships'
    $noteRecords = Get-DraftArray $root 'annotations'
    $groupRecords = Get-DraftArray $root 'groups'
    $fileRecords = Get-DraftArray $root 'embeddedFiles'
    $inputRecords = $nodeRecords.Count + $edgeRecords.Count + $noteRecords.Count
    if ($inputRecords + $groupRecords.Count -gt $maxRecords) { throw 'Draft input exceeds the 2000 record limit.' }
    if ($fileRecords.Count -gt $maxRecords) { throw 'Draft input exceeds the 2000 embedded file limit.' }
    foreach ($record in (@($nodeRecords) + @($edgeRecords) + @($noteRecords))) { [void](Add-DraftRecord $recordMap $record) }
    foreach ($record in $nodeRecords) {
        $id = $record.id
        $kind = Get-DraftString $record 'kind' $false 'shape'
        if ($kind -cnotin @('shape','group','image','sketch','unknown')) { throw "Unsupported component kind '$kind'; use neutral draft kinds, not Azure service types." }
        $review = (Get-DraftBool $record 'requiresVisualReview') -or $kind -in @('image','sketch','unknown')
        $node = [pscustomobject][ordered]@{
            id=$id; safeId=(New-DraftSafeId $id); label=(Get-DraftString $record 'label' $true)
            sourceIds=(Get-DraftSourceIds $record $id); kind=$kind
            sourceType=(Get-DraftString $record 'sourceType' $false 'manual')
            geometry=(Get-DraftGeometry $record $id $true)
            labelParts=(Get-DraftLabelParts $record); frameId=(Get-DraftString $record 'frameId')
            groupIds=(Get-DraftStrings $record 'groupIds'); requiresVisualReview=$review
            fileId=(Get-DraftString $record 'fileId'); link=(Get-DraftString $record 'link')
        }
        if ($review) { Add-DraftDiagnostic $unresolved 'VisualReviewRequired' "Component '$id' needs visual review." @($id) }
        $components.Add($node)
        $componentMap.Add($id,$node)
    }
    foreach ($record in $edgeRecords) {
        $id = $record.id
        $sourceId = Get-DraftString $record 'source'
        $targetId = Get-DraftString $record 'target'
        $direction = Get-DraftString $record 'direction' $true
        if ($direction -cnotin @('forward','reverse','bidirectional','none','unknown')) { throw "Unsupported relationship direction '$direction'." }
        $source = $null; $target = $null
        if (-not [string]::IsNullOrEmpty($sourceId) -and $componentMap.ContainsKey($sourceId)) { $source = $sourceId }
        else { Add-DraftDiagnostic $unresolved 'UnresolvedEndpoint' "Relationship '$id' has an absent or unknown declared source." @($id,$sourceId) }
        if (-not [string]::IsNullOrEmpty($targetId) -and $componentMap.ContainsKey($targetId)) { $target = $targetId }
        else { Add-DraftDiagnostic $unresolved 'UnresolvedEndpoint' "Relationship '$id' has an absent or unknown declared target." @($id,$targetId) }
        if (Get-DraftBool $record 'dataFlowVerified') {
            Add-DraftDiagnostic $warnings 'VerificationNotImported' "Relationship '$id' verification claim was not imported; source arrows are not verified data flow." @($id)
        }
        $provenance = Get-DraftString $record 'provenance' $false 'manual-interpretation'
        if ($provenance -cnotin @('source-diagram','manual-interpretation')) { throw "Unsupported relationship provenance '$provenance'." }
        foreach ($bindingName in @('startBinding','endBinding')) {
            $binding = Get-DraftValue $record $bindingName
            if ($null -ne $binding) {
                if (-not (Test-DraftObject $binding)) { throw "'$bindingName' must be an object or null." }
                [void](Get-DraftString $binding 'elementId' $true)
            }
        }
        $relationships.Add([pscustomobject][ordered]@{
            id=$id; source=$source; target=$target
            declaredSource=(Get-DraftString $record 'declaredSource' $false $sourceId)
            declaredTarget=(Get-DraftString $record 'declaredTarget' $false $targetId)
            label=(Get-DraftString $record 'label' $false ''); direction=$direction
            sourceIds=(Get-DraftSourceIds $record $id); confidence=(Get-DraftConfidence $source $target $true)
            geometry=(Get-DraftGeometry $record $id $true); labelParts=(Get-DraftLabelParts $record)
            frameId=(Get-DraftString $record 'frameId'); groupIds=(Get-DraftStrings $record 'groupIds')
            provenance=$provenance; dataFlowVerified=$false
            startBinding=(Get-DraftValue $record 'startBinding'); endBinding=(Get-DraftValue $record 'endBinding')
            startArrowhead=(Get-DraftString $record 'startArrowhead'); endArrowhead=(Get-DraftString $record 'endArrowhead')
            sourceType=(Get-DraftString $record 'sourceType' $false 'manual')
            link=(Get-DraftString $record 'link')
        })
    }
    foreach ($record in $noteRecords) {
        $id = $record.id
        $annotations.Add([pscustomobject][ordered]@{
            id=$id; label=(Get-DraftString $record 'label' $true)
            sourceIds=(Get-DraftSourceIds $record $id); kind=(Get-DraftString $record 'kind' $false 'label')
            text=(Get-DraftString $record 'text'); originalText=(Get-DraftString $record 'originalText')
            geometry=(Get-DraftGeometry $record $id $true); labelParts=(Get-DraftLabelParts $record)
            frameId=(Get-DraftString $record 'frameId'); groupIds=(Get-DraftStrings $record 'groupIds')
            containerId=(Get-DraftString $record 'containerId'); link=(Get-DraftString $record 'link')
        })
    }
    $knownSourceIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($record in (@($components.ToArray()) + @($relationships.ToArray()) + @($annotations.ToArray()))) {
        [void]$knownSourceIds.Add($record.id)
        foreach ($sourceId in $record.sourceIds) { [void]$knownSourceIds.Add($sourceId) }
        foreach ($part in $record.labelParts) { [void]$knownSourceIds.Add($part.sourceId); $attachedTexts++ }
    }
    $groupMap = New-DraftMap
    foreach ($record in $groupRecords) {
        $id = Add-DraftRecord $groupMap $record
        $members = Get-DraftStrings $record 'memberIds'
        foreach ($member in $members) {
            if (-not $knownSourceIds.Contains($member)) { Add-DraftDiagnostic $unresolved 'UnresolvedGroupMember' "Group '$id' names missing source record '$member'." @($id,$member) }
        }
        $group = [pscustomobject][ordered]@{id=$id; memberIds=@($members); provenance='declared'}
        $groups.Add($group)
        $groupMap[$id] = $group
    }
    foreach ($record in (@($components.ToArray()) + @($relationships.ToArray()) + @($annotations.ToArray()))) {
        foreach ($grouped in (@($record) + @($record.labelParts))) {
            $memberId = Get-DraftString $grouped 'id' $false (Get-DraftValue $grouped 'sourceId')
            foreach ($groupId in $grouped.groupIds) {
                if (-not $groupMap.ContainsKey($groupId)) {
                    $group = [pscustomobject][ordered]@{id=$groupId; memberIds=@(); provenance='explicit-groupIds'}
                    $groupMap.Add($groupId,$group)
                    $groups.Add($group)
                }
                if (-not (@($groupMap[$groupId].memberIds) -ccontains $memberId)) { $groupMap[$groupId].memberIds += $memberId }
            }
        }
    }
    $fileMap = New-DraftMap
    foreach ($record in $fileRecords) {
        $id = Add-DraftRecord $fileMap $record
        $embeddedFiles.Add([pscustomobject][ordered]@{
            id=$id; mimeType=(Get-DraftString $record 'mimeType')
            hasEmbeddedData=((Get-DraftBool $record 'hasEmbeddedData') -or $null -ne (Get-DraftValue $record 'dataURL'))
            requiresVisualReview=$true
        })
    }
    if ($embeddedFiles.Count -gt 0) {
        Add-DraftDiagnostic $warnings 'EmbeddedPayloadNotDecoded' 'Embedded file descriptors were retained without decoding or copying payloads.'
    }
    Import-DraftDiagnostics $root 'warnings' $warnings
    Import-DraftDiagnostics $root 'unresolved' $unresolved
    Add-DraftDiagnostic $warnings 'NormalizedInputNotVerified' 'Normalized graph records are supplied interpretations, not machine-read images/ASCII or verified Azure architecture.'
}

foreach ($record in (@($components.ToArray()) + @($relationships.ToArray()) + @($annotations.ToArray()))) {
    $groupedRecords = @($record) + @($record.labelParts)
    foreach ($grouped in $groupedRecords) {
        $groupedId = Get-DraftString $grouped 'id' $false (Get-DraftValue $grouped 'sourceId')
        if (-not [string]::IsNullOrEmpty($grouped.frameId)) {
            if (-not $componentMap.ContainsKey($grouped.frameId) -or $componentMap[$grouped.frameId].kind -ne 'group' -or $grouped.frameId -ceq $groupedId) {
                Add-DraftDiagnostic $unresolved 'UnresolvedFrame' "Record '$groupedId' has invalid explicit frame '$($grouped.frameId)'." @($groupedId,$grouped.frameId)
            }
        }
    }
}
$status = 'ready-for-review'
if ($components.Count -eq 0) {
    $status = 'empty'
    Add-DraftDiagnostic $unresolved 'NoComponents' 'No architecture components were extracted. Review annotations or supply a manually interpreted normalized graph; no default topology was generated.'
}
Add-DraftDiagnostic $warnings 'SourceDiagramNotVerified' 'Source labels, grouping, and connectors are evidence only. Azure service mapping, containment semantics, and data flow require explicit review.'
$graph = [pscustomobject][ordered]@{
    schemaVersion=1; sourceFormat=$sourceFormat; status=$status
    components=@($components.ToArray()); relationships=@($relationships.ToArray())
    annotations=@($annotations.ToArray()); groups=@($groups.ToArray())
    embeddedFiles=@($embeddedFiles.ToArray())
    unresolved=@($unresolved.ToArray()); warnings=@($warnings.ToArray())
    counts=[pscustomobject][ordered]@{
        inputRecords=$inputRecords; activeRecords=($inputRecords-$deletedRecords)
        deletedRecords=$deletedRecords; components=$components.Count
        relationships=$relationships.Count; annotations=$annotations.Count
        attachedTexts=$attachedTexts; groups=$groups.Count; embeddedFiles=$embeddedFiles.Count
        unresolved=$unresolved.Count; warnings=$warnings.Count
    }
}
$resultJson = ConvertTo-Json -InputObject $graph -Depth 60
if ($null -ne $outputFile) {
    # CreateNew provides the no-overwrite guarantee even if a file appears after validation.
    $outputStream = [IO.File]::Open($outputFile, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $bytes = [Text.UTF8Encoding]::new($false).GetBytes($resultJson + [Environment]::NewLine)
        $outputStream.Write($bytes,0,$bytes.Length)
        $outputStream.Flush()
    } finally { $outputStream.Dispose() }
} else { Write-Output $resultJson }

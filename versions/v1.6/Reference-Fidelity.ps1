<#
.SYNOPSIS
Fail-closed, filesystem-only comparison against a separately extracted source contract.
.DESCRIPTION
Contract (unknown properties are rejected):
{schemaVersion:1,mode:"source-faithful",
 source:{path:"source.bin",sha256:"64 hex characters",role:"authoritative-reference"},
 referencePage:"01 Reference architecture",allowAdditionalPages:false,
 components:[{id:"a",label:"Original title",requiredText:["Original detail"],
              parent:null,kind:"card"}],
 relationships:[{id:"ab",source:"a",target:"b",direction:"forward",requiredText:[]}],
 layout:{leftToRight:[["a","b"]],topToBottom:[],aspectRatio:1.7777778,aspectTolerance:0.15},
 unresolved:[]}
All fields above are required except allowAdditionalPages (false), aspectTolerance
(0.15), and relationship requiredText ([]). Component kind is card/container/note;
direction is forward/backward/both/none, relative to the recorded source/target.
Nonempty unresolved blocks acceptance; it is an array of explanatory strings.
This does NOT extract images or verify Azure services, data flow, or contract truth.

The reference page must contain exactly one node/edge per contract sourceId.
Native parent and endpoint IDs are mapped to case-sensitive sourceIds. Titles are
single lines; requiredText contains complete logical lines (the title may repeat
there). Nonblank model lines must match the contracted lines exactly, ignoring
case and whitespace, with no extra/duplicate lines. Component title is first.
Layout chains contain at least two IDs; boxes must be strictly separated.
Visio Y increases upward. Aspect tolerance is fractional relative error, [0,1).
Other model properties are intentionally left to controller structural validation.

Paths must be local drive paths, without streams, device names, link reparse
points, or traversal segments. Hydrated OneDrive cloud tags are allowed after
read-only Windows fsutil metadata inspection; offline files are never hydrated.
A relative source path stays inside the contract folder.
JSON is limited to 10 MiB, 64 levels, 2000 components/relationships each, 100 pages;
source bytes to 100 MiB. Source contents are only hashed and are never executed.
Emits JSON always; invalid input throws unless -ReportOnly is specified.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ModelPath,
    [Parameter(Mandatory)][string]$ReferencePath,
    [switch]$ReportOnly
)
Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
$issues = [Collections.Generic.List[object]]::new()
$cloudPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$report = [ordered]@{
    schemaVersion=1; mode='source-faithful'; valid=$false; modelPath=$ModelPath; referencePath=$ReferencePath
    referencePage=$null; source=[ordered]@{path=$null; expectedSha256=$null; actualSha256=$null; verified=$false}
    counts=[ordered]@{expectedComponents=0; matchedComponents=0; expectedRelationships=0; matchedRelationships=0}
    errors=@(); comUsed=$false; imageRecognitionPerformed=$false; dataFlowVerified=$false
}
function Fail([string]$Code, [string]$Message) {
    $exception = [IO.InvalidDataException]::new($Message)
    $exception.Data['code'] = $Code
    throw $exception
}
function Issue([string]$Code, [string]$Message) { $issues.Add([pscustomobject]@{code=$Code; message=$Message}) }
function Value($Object, [string]$Name, $Default=$null) {
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return ,$Default }
    return ,$property.Value
}
function ObjectFields($Object, [string[]]$Required, [string[]]$Optional=@(), [switch]$Open) {
    if ($null -eq $Object -or $Object.GetType().FullName -cne 'System.Management.Automation.PSCustomObject') { Fail 'InvalidInput' 'Expected a JSON object.' }
    $names = @($Object.PSObject.Properties.Name)
    foreach ($name in $Required) { if ($names -cnotcontains $name) { Fail 'InvalidInput' "Missing property '$name'." } }
    if (-not $Open) {
        foreach ($name in $names) { if (($Required+$Optional) -cnotcontains $name) { Fail 'InvalidInput' "Unknown property '$name'." } }
    }
}
function Text($Value, [string]$Context, [switch]$Empty) {
    if ($Value -isnot [string] -or $Value.Length -gt 32768 -or (-not $Empty -and [string]::IsNullOrWhiteSpace($Value))) { Fail 'InvalidInput' "$Context must be a bounded string." }
}
function ArrayValue($Value, [string]$Context, [int]$Maximum=2000) {
    if ($Value -isnot [array] -or $Value.Count -gt $Maximum) { Fail 'InvalidInput' "$Context must be an array of at most $Maximum items." }
}
function Number($Value, [string]$Context, [switch]$Positive) {
    if ($null -eq $Value -or $Value.GetType().Name -cnotin @('Int32','Int64','Double','Decimal') -or
        [double]::IsNaN($Value) -or [double]::IsInfinity($Value) -or [math]::Abs($Value) -gt 1000000 -or
        ($Positive -and $Value -le 0)) { Fail 'InvalidInput' "$Context must be a finite number in range." }
}
function Map { return ,([Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)) }
function Normalize([string]$Text) { return [regex]::Replace($Text.Trim(), '\s+', ' ').ToUpperInvariant() }
function Lines($Items) {
    ArrayValue $Items 'requiredText'
    foreach ($item in $Items) {
        Text $item 'requiredText entry'
        if ($item -match '[\r\n]') { Fail 'InvalidInput' 'Contract text entries must be single logical lines.' }
    }
}
function LocalPath([string]$Path, [string]$Base='') {
    if ($Path -notmatch '^[A-Za-z]:\\') {
        if (-not $Base -or $Path -match '^(?:[\\/]|[A-Za-z]:)') { Fail 'UnsafePath' 'An absolute local drive path is required.' }
        $Path = Join-Path $Base $Path
    }
    if ($Path.Substring(3) -match '[\x00-\x1f<>:"/|?*]') { Fail 'UnsafePath' 'Streams, URLs, and invalid path characters are forbidden.' }
    foreach ($part in ($Path.Substring(3) -split '\\')) {
        if (-not $part -or $part -match '[. ]$|^(?i:CON|PRN|AUX|NUL|COM[1-9\u00b9\u00b2\u00b3]|LPT[1-9\u00b9\u00b2\u00b3])(?:\.|$)') { Fail 'UnsafePath' 'Traversal, device names, and ambiguous path segments are forbidden.' }
    }
    $full = [IO.Path]::GetFullPath($Path)
    if ($Base -and -not $full.StartsWith($Base.TrimEnd('\')+'\', [StringComparison]::OrdinalIgnoreCase)) { Fail 'UnsafePath' 'Relative source escaped the contract folder.' }
    $drive = [IO.DriveInfo]::new([IO.Path]::GetPathRoot($full))
    if ($drive.DriveType -eq [IO.DriveType]::Network) { Fail 'UnsafePath' 'Network drives are forbidden.' }
    # Inspect ancestors before children, so a directory link cannot redirect a probe.
    $cursor = [IO.Path]::GetPathRoot($full)
    foreach ($part in ($full.Substring(3) -split '\\')) {
        $cursor=Join-Path $cursor $part
        try { $attributes=[IO.File]::GetAttributes($cursor) }
        catch [IO.FileNotFoundException] { break }
        catch [IO.DirectoryNotFoundException] { break }
        if (($attributes -band [IO.FileAttributes]::Directory) -eq 0 -and ([int]$attributes -band 0x441000) -ne 0) { Fail 'UnsafePath' 'Offline/recall files must already be locally hydrated; no download is permitted.' }
        if (($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -and -not $cloudPaths.Contains($cursor)) {
            $fsutil=Join-Path ([Environment]::SystemDirectory) 'fsutil.exe'
            $metadata=(& $fsutil reparsepoint query $cursor 2>&1) -join "`n"
            if ($LASTEXITCODE -ne 0 -or $metadata -notmatch '(?im)^[^\r\n]*:\s*0x9000[0-9a-f]01a\s*$') { Fail 'UnsafePath' 'Only non-link cloud reparse tags are allowed; symbolic links/junctions are forbidden.' }
            [void]$cloudPaths.Add($cursor)
        }
    }
    return $full
}
function ReadJson([string]$Path) {
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        if ($stream.Length -gt 10MB) { Fail 'InputLimit' 'JSON exceeds 10 MiB.' }
        $reader = [IO.StreamReader]::new($stream, [Text.UTF8Encoding]::new($false,$true), $true)
        try { $json = $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $stream.Dispose() }
    # ConvertFrom-Json discards duplicate keys. Inspect property tokens before conversion.
    $tokens = [regex]::Matches($json, '"(?:[^"\\\x00-\x1f]|\\(?:["\\/bfnrt]|u[0-9a-fA-F]{4}))*"|[{}\[\]:,]|[^\s{}\[\]:,]+')
    $stack = [Collections.Generic.List[object]]::new()
    for ($i=0; $i -lt $tokens.Count; $i++) {
        $token = $tokens[$i].Value
        if ($token -in @('{','[')) {
            if ($stack.Count -ge 64) { Fail 'InputLimit' 'JSON exceeds 64 nesting levels.' }
            $frame=$null
            if ($token -eq '{') { $frame=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
            $stack.Add($frame)
        } elseif ($token -in @('}',']')) {
            if ($stack.Count -eq 0) { Fail 'InvalidInput' 'Invalid JSON nesting.' }
            $stack.RemoveAt($stack.Count-1)
        } elseif ($token.StartsWith('"') -and $i+1 -lt $tokens.Count -and $tokens[$i+1].Value -eq ':') {
            if ($stack.Count -eq 0 -or $null -eq $stack[$stack.Count-1]) { Fail 'InvalidInput' 'Invalid JSON property.' }
            $key = $token.Substring(1,$token.Length-2)
            if ($key.Contains('\')) { $key = (ConvertFrom-Json -InputObject ('{"key":'+$token+'}')).key }
            if (-not $stack[$stack.Count-1].Add($key)) { Fail 'DuplicateProperty' "Duplicate JSON property '$key'." }
        }
    }
    if ($stack.Count -ne 0) { Fail 'InvalidInput' 'Unclosed JSON structure.' }
    # Wrapping preserves root arrays/scalars on Windows PowerShell 5.1.
    $wrapper = ConvertFrom-Json -InputObject ('{"document":'+$json+'}')
    return ,$wrapper.document
}
function CheckLabel($Actual, [string]$Title, $Required, [string]$Id) {
    Text $Actual "label for '$Id'" -Empty
    $expected = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    if ($Title) { [void]$expected.Add((Normalize $Title)) }
    foreach ($text in $Required) { [void]$expected.Add((Normalize $text)) }
    $actualLines = @($Actual -split '\r\n|\n|\r' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { Normalize $_ })
    if ($Title -and ($actualLines.Count -eq 0 -or $actualLines[0] -cne (Normalize $Title))) { Issue 'TitleMismatch' "Original title changed for '$Id'." }
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($line in $actualLines) {
        if (-not $expected.Contains($line)) { Issue 'UnexpectedText' "Uncontracted label line for '$Id'." }
        if (-not $seen.Add($line)) { Issue 'DuplicateText' "Repeated label line for '$Id'." }
    }
    foreach ($line in $expected) { if (-not $seen.Contains($line)) { Issue 'MissingText' "Original label detail missing for '$Id': $line" } }
}
try {
    $report.modelPath = LocalPath $ModelPath
    $report.referencePath = LocalPath $ReferencePath
    $contract = ReadJson $report.referencePath
    ObjectFields $contract @('schemaVersion','mode','source','referencePage','components','relationships','layout','unresolved') @('allowAdditionalPages')
    Number $contract.schemaVersion 'contract schemaVersion'
    if ($contract.schemaVersion -ne 1 -or $contract.mode -isnot [string] -or $contract.mode -cne 'source-faithful') { Fail 'InvalidInput' 'Require schemaVersion 1 and mode source-faithful.' }
    ObjectFields $contract.source @('path','sha256','role')
    Text $contract.source.path 'source.path'; Text $contract.source.sha256 'source.sha256'
    if ($contract.source.role -isnot [string] -or $contract.source.role -cne 'authoritative-reference') { Fail 'SourceRole' 'Only an authoritative-reference can be the source, never an output regression.' }
    if ($contract.source.sha256 -notmatch '\A[0-9A-Fa-f]{64}\z') { Fail 'InvalidInput' 'source.sha256 must contain exactly 64 hexadecimal characters.' }
    $base = ''; if ($contract.source.path -notmatch '^[A-Za-z]:\\') { $base = [IO.Path]::GetDirectoryName($report.referencePath) }
    $report.source.path = LocalPath $contract.source.path $base
    if ($report.source.path -ieq $report.modelPath -or $report.source.path -ieq $report.referencePath) { Fail 'SourceRole' 'Source must be a separate artifact from model and contract.' }
    $report.source.expectedSha256 = $contract.source.sha256.ToLowerInvariant()
    if (-not [IO.File]::Exists($report.source.path)) { Fail 'MissingSource' 'Persisted authoritative source file is missing; no fallback is permitted.' }
    $stream = [IO.File]::Open($report.source.path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        if ($stream.Length -gt 100MB -or $stream.Length -eq 0) { Fail 'InputLimit' 'Source must contain 1 byte to 100 MiB.' }
        $sha = [Security.Cryptography.SHA256]::Create()
        try { $report.source.actualSha256 = [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','').ToLowerInvariant() } finally { $sha.Dispose() }
    } finally { $stream.Dispose() }
    $report.source.verified = $report.source.actualSha256 -ceq $report.source.expectedSha256
    if (-not $report.source.verified) { Issue 'SourceHashMismatch' 'Persisted source SHA256 changed; reacquire/review the source rather than redesigning.' }
    Text $contract.referencePage 'referencePage'; $report.referencePage = $contract.referencePage
    $allowPages = Value $contract 'allowAdditionalPages' $false
    if ($allowPages -isnot [bool]) { Fail 'InvalidInput' 'allowAdditionalPages must be boolean.' }
    ArrayValue $contract.unresolved 'unresolved'
    foreach ($item in $contract.unresolved) { Text $item 'unresolved entry' }
    if ($contract.unresolved.Count) { Issue 'UnresolvedSource' 'Resolve source extraction uncertainties before accepting a faithful rendition.' }
    ArrayValue $contract.components 'components'; ArrayValue $contract.relationships 'relationships'
    if (-not $contract.components.Count) { Fail 'InvalidInput' 'The source contract must contain components.' }
    $components = Map; $relationships = Map; $allSourceIds = Map
    foreach ($category in @('components','relationships')) {
        foreach ($item in $contract.$category) {
            if ($category -eq 'components') { ObjectFields $item @('id','label','requiredText','parent','kind') }
            else { ObjectFields $item @('id','source','target','direction') @('requiredText') }
            Text $item.id 'source id'
            if ($allSourceIds.ContainsKey($item.id)) { Fail 'DuplicateSourceId' "Duplicate contract source id '$($item.id)'." }
            $allSourceIds.Add($item.id,$true)
            Lines (Value $item 'requiredText' @())
            if ($category -eq 'components') {
                Text $item.label 'component label'
                if ($item.label -match '[\r\n]' -or $item.kind -isnot [string] -or $item.kind -cnotin @('card','container','note')) { Fail 'InvalidInput' 'Component title must be one line and kind card/container/note.' }
                if ($null -ne $item.parent) { Text $item.parent 'component parent' }
                $components.Add($item.id,$item)
            } else {
                Text $item.source 'relationship source'; Text $item.target 'relationship target'
                if ($item.direction -isnot [string] -or $item.direction -cnotin @('forward','backward','both','none')) { Fail 'InvalidInput' 'Relationship direction must be explicit.' }
                $relationships.Add($item.id,$item)
            }
        }
    }
    $report.counts.expectedComponents=$components.Count; $report.counts.expectedRelationships=$relationships.Count
    foreach ($item in $components.Values) {
        $ancestors = Map; $current = $item
        while ($null -ne $current.parent) {
            if (-not $components.ContainsKey($current.parent) -or $components[$current.parent].kind -cne 'container') { Fail 'InvalidInput' 'Contract parent must reference a source container.' }
            if ($ancestors.ContainsKey($current.parent)) { Fail 'InvalidInput' 'Contract parent cycle.' }
            $ancestors.Add($current.parent,$true); $current=$components[$current.parent]
        }
    }
    foreach ($item in $relationships.Values) {
        if (-not $components.ContainsKey($item.source) -or -not $components.ContainsKey($item.target)) { Fail 'InvalidInput' 'Contract relationship has an unknown endpoint.' }
    }
    ObjectFields $contract.layout @('leftToRight','topToBottom','aspectRatio') @('aspectTolerance')
    Number $contract.layout.aspectRatio 'aspectRatio' -Positive
    $tolerance = Value $contract.layout 'aspectTolerance' 0.15
    Number $tolerance 'aspectTolerance'
    if ($tolerance -lt 0 -or $tolerance -ge 1) { Fail 'InvalidInput' 'aspectTolerance must be at least zero and less than one.' }
    foreach ($axis in @('leftToRight','topToBottom')) {
        ArrayValue $contract.layout.$axis $axis
        foreach ($chain in $contract.layout.$axis) {
            ArrayValue $chain 'layout chain'
            if ($chain.Count -lt 2) { Fail 'InvalidInput' 'Layout chains require at least two source IDs.' }
            $seen=Map
            foreach ($id in $chain) {
                Text $id 'layout id'
                if (-not $components.ContainsKey($id) -or $seen.ContainsKey($id)) { Fail 'InvalidInput' 'Layout has an unknown or repeated source ID.' }
                $seen.Add($id,$true)
            }
        }
    }
    $model = ReadJson $report.modelPath
    ObjectFields $model @('schemaVersion','pages') -Open
    Number $model.schemaVersion 'model schemaVersion'
    if ($model.schemaVersion -ne 1) { Fail 'InvalidInput' 'Model schemaVersion must be 1.' }
    ArrayValue $model.pages 'pages' 100
    $pageNames=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); $reference=$null
    foreach ($page in $model.pages) {
        ObjectFields $page @('name','width','height','nodes','edges') -Open
        Text $page.name 'page name'; Number $page.width 'page width' -Positive; Number $page.height 'page height' -Positive
        if (-not $pageNames.Add($page.name)) { Fail 'InvalidInput' 'Duplicate page name.' }
        ArrayValue $page.nodes 'nodes'; ArrayValue $page.edges 'edges'
        $native=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($category in @('nodes','edges')) {
            foreach ($item in $page.$category) {
                ObjectFields $item @('id') -Open; Text $item.id 'model id'
                if (-not $native.Add($item.id)) { Fail 'DuplicateModelId' "Duplicate model id '$($item.id)'." }
                if ($null -ne $item.PSObject.Properties['sourceId']) { Text $item.sourceId 'model sourceId' }
                if ($category -eq 'nodes') {
                    ObjectFields $item @('id','kind','label','x','y','width','height') -Open
                    Text $item.kind 'node kind'; Text $item.label 'node label' -Empty
                    if ($item.kind -cnotin @('card','container','note')) { Fail 'InvalidInput' 'Unsupported model node kind.' }
                    foreach ($coordinate in @('x','y')) { Number $item.$coordinate $coordinate }
                    foreach ($size in @('width','height')) { Number $item.$size $size -Positive }
                    $parent=Value $item 'parent'; if ($null -ne $parent) { Text $parent 'model parent' -Empty }
                } else {
                    ObjectFields $item @('id','source','target') -Open
                    foreach ($endpoint in @('source','target')) { if ($null -ne $item.$endpoint) { Text $item.$endpoint 'model endpoint' } }
                    Text (Value $item 'label' '') 'edge label' -Empty
                    if ($null -ne $item.PSObject.Properties['direction']) { Text $item.direction 'edge direction' }
                }
            }
        }
        if ($page.name -ceq $contract.referencePage) { $reference=$page }
    }
    if (-not $allowPages -and $model.pages.Count -gt 1) { Issue 'AdditionalPages' 'Additional proposal/review pages require explicit contract permission.' }
    if ($null -eq $reference) { Fail 'MissingReferencePage' 'The designated reference page is missing; components elsewhere cannot replace it.' }
    $nodes=Map; $nativeNodes=Map; $edges=Map
    foreach ($category in @('nodes','edges')) {
        $expected=$components; $mapped=$nodes; if ($category -eq 'edges') { $expected=$relationships; $mapped=$edges }
        foreach ($item in $reference.$category) {
            if ($category -eq 'nodes') {
                $nativeNodes.Add($item.id,$item)
            } else { ObjectFields $item @('id','source','target','direction') -Open }
            $sourceId=Value $item 'sourceId'
            if ($null -eq $sourceId -or -not $expected.ContainsKey($sourceId)) { Issue 'UnexpectedItem' "Uncontracted $category item '$($item.id)' on reference page."; continue }
            if ($mapped.ContainsKey($sourceId)) { Issue 'DuplicateSourceId' "Repeated model sourceId '$sourceId'."; continue }
            $mapped.Add($sourceId,$item)
        }
        foreach ($id in $expected.Keys) { if (-not $mapped.ContainsKey($id)) { Issue 'MissingItem' "Missing contracted $category sourceId '$id' on reference page." } }
    }
    $report.counts.matchedComponents=$nodes.Count; $report.counts.matchedRelationships=$edges.Count
    foreach ($id in $nodes.Keys) {
        $item=$nodes[$id]; $expected=$components[$id]
        if ($item.kind -cne $expected.kind) { Issue 'KindMismatch' "Component kind changed for '$id'." }
        CheckLabel $item.label $expected.label $expected.requiredText $id
        $parent=Value $item 'parent'; $parentSource=$null
        if ($null -ne $parent) { Text $parent 'model parent' -Empty }
        if ($parent) {
            if (-not $nativeNodes.ContainsKey($parent)) { Issue 'ParentMismatch' "Unknown parent for '$id'."; continue }
            $parentSource=Value $nativeNodes[$parent] 'sourceId'
            if ($null -eq $parentSource) { Issue 'ParentMismatch' "Uncontracted parent for '$id'."; continue }
        }
        if ($parentSource -cne $expected.parent) { Issue 'ParentMismatch' "Source grouping changed for '$id'." }
    }
    foreach ($id in $edges.Keys) {
        $item=$edges[$id]; $expected=$relationships[$id]
        foreach ($endpoint in @('source','target')) {
            Text $item.$endpoint 'model relationship endpoint'
            if (-not $nativeNodes.ContainsKey($item.$endpoint) -or (Value $nativeNodes[$item.$endpoint] 'sourceId') -cne $expected.$endpoint) { Issue 'EndpointMismatch' "Relationship $endpoint changed for '$id'." }
        }
        Text $item.direction 'model direction'
        if ($item.direction -cne $expected.direction) { Issue 'DirectionMismatch' "Relationship direction changed for '$id'." }
        CheckLabel (Value $item 'label' '') '' (Value $expected 'requiredText' @()) $id
    }
    foreach ($axis in @('leftToRight','topToBottom')) {
        foreach ($chain in $contract.layout.$axis) {
            for ($i=1; $i -lt $chain.Count; $i++) {
                if (-not $nodes.ContainsKey($chain[$i-1]) -or -not $nodes.ContainsKey($chain[$i])) { continue }
                $a=$nodes[$chain[$i-1]]; $b=$nodes[$chain[$i]]
                $gap=$b.x-$b.width/2-($a.x+$a.width/2)
                if ($axis -eq 'topToBottom') { $gap=$a.y-$a.height/2-($b.y+$b.height/2) }
                if ($gap -le 0) { Issue 'LayoutMismatch' "$axis ordering/clear separation lost between '$($chain[$i-1])' and '$($chain[$i])'." }
            }
        }
    }
    $aspectError=[math]::Abs(($reference.width/$reference.height)/$contract.layout.aspectRatio-1)
    if ($aspectError -gt $tolerance) { Issue 'AspectMismatch' 'Reference page aspect ratio is outside the contracted tolerance.' }
} catch [IO.InvalidDataException], [IO.IOException], [UnauthorizedAccessException], [ArgumentException], [Text.DecoderFallbackException] {
    $code='InvalidInput'
    if ($_.Exception.Data.Contains('code')) { $code=$_.Exception.Data['code'] }
    Issue $code $_.Exception.Message
}
$report.errors=@($issues.ToArray())
$report.valid=$issues.Count -eq 0
ConvertTo-Json -InputObject $report -Depth 12
if (-not $report.valid -and -not $ReportOnly) { throw "Reference fidelity failed: $(($issues | ForEach-Object { $_.code+': '+$_.message }) -join '; ')" }

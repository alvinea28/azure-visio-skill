[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$OutputDirectory
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
if ($OutputDirectory -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+(?:\\|$))') {
    throw 'Provide an absolute OutputDirectory for filesystem-only draft import tests.'
}
[void][IO.Directory]::CreateDirectory($OutputDirectory)
$fixtures = Join-Path $OutputDirectory ('Draft-Import-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($fixtures)
$importer = Join-Path $PSScriptRoot 'Import-Draft.ps1'
$checks = [System.Collections.Generic.List[string]]::new()
$utf8 = [Text.UTF8Encoding]::new($false)

function Assert-Draft([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "DRAFT ASSERTION FAILED: $Message" }
}

function Assert-DraftThrows([scriptblock]$Operation, [string]$Pattern) {
    $thrown = $false
    try { & $Operation | Out-Null }
    catch {
        if ($_.Exception.Message -notlike $Pattern) { throw }
        $thrown = $true
    }
    Assert-Draft $thrown "Expected error matching $Pattern"
}

function Write-DraftJson([string]$Name, $Value) {
    $path = Join-Path $fixtures $Name
    [IO.File]::WriteAllText($path, (ConvertTo-Json -InputObject $Value -Depth 30), $utf8)
    return $path
}

function Write-DraftText([string]$Name, [string]$Text) {
    $path = Join-Path $fixtures $Name
    [IO.File]::WriteAllText($path, $Text, $utf8)
    return $path
}

function Invoke-Draft([string]$Path) {
    return (ConvertFrom-Json -InputObject ((& $importer -InputPath $Path) -join "`n"))
}

function New-DraftFixtureElement([string]$Id, [string]$Type) {
    return [ordered]@{id=$Id; type=$Type; x=0; y=0; width=100; height=60; angle=0; isDeleted=$false}
}

function New-DraftFixtureEdge([string]$Id, $StartHead, $EndHead) {
    $edge = New-DraftFixtureElement $Id 'arrow'
    $edge.startArrowhead = $StartHead
    $edge.endArrowhead = $EndHead
    $edge.startBinding = @{elementId='a'; focus=0; gap=1}
    $edge.endBinding = @{elementId='b'; focus=0; gap=1}
    $edge.points = @(@(0,0),@(100,20))
    return $edge
}

try {
    foreach ($path in @($importer, $PSCommandPath)) {
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
        Assert-Draft ($errors.Count -eq 0) "PowerShell syntax in $path : $errors"
        if ($path -eq $importer) {
            $unsafe = @($ast.FindAll({
                param($node)
                $node -is [Management.Automation.Language.CommandAst] -and
                    $node.GetCommandName() -in @('Invoke-Expression','Invoke-WebRequest','Invoke-RestMethod','Start-Process','iex','curl','Add-Type')
            }, $true))
            Assert-Draft ($unsafe.Count -eq 0) 'Importer has no code execution, network, or external tool commands.'
            $broadCatches = @($ast.FindAll({
                param($node)
                $node -is [Management.Automation.Language.CatchClauseAst] -and $node.CatchTypes.Count -eq 0
            }, $true))
            Assert-Draft ($broadCatches.Count -eq 0) 'Importer only catches expected exception types.'
        }
    }
    $checks.Add('PS5.1 parsing; no execution/network/external-tool commands.')

    $a = New-DraftFixtureElement 'a' 'rectangle'
    $a.boundElements = @(@{id='at'; type='text'},@{id='ab'; type='arrow'})
    $a.frameId = 'frame'
    $a.groupIds = @('explicit-group')
    $b = New-DraftFixtureElement 'b' 'diamond'
    $b.x = 250
    $b.groupIds = @('explicit-group')
    $frame = New-DraftFixtureElement 'frame' 'frame'
    $frame.name = 'Application area'
    $frame.width = 600
    $frame.height = 400
    $at = New-DraftFixtureElement 'at' 'text'
    $at.text = 'API wrapped'
    $at.originalText = 'API original'
    $at.frameId = 'frame'
    $at.groupIds = @('explicit-group')
    $bt = New-DraftFixtureElement 'bt' 'text'
    $bt.containerId = 'b'
    $bt.text = 'Database'
    $edge = New-DraftFixtureEdge 'ab' $null 'arrow'
    $edge.boundElements = @(@{id='edge-label';type='text'})
    $edgeLabel = New-DraftFixtureElement 'edge-label' 'text'
    $edgeLabel.containerId = 'ab'
    $edgeLabel.text = 'request'
    $note = New-DraftFixtureElement 'note' 'text'
    $note.text = 'Standalone note'
    $note.frameId = 'frame'
    $deleted = New-DraftFixtureElement 'deleted' 'rectangle'
    $deleted.isDeleted = $true
    $deletedText = @{id='deleted-text';isDeleted=$true}
    $happyPath = Write-DraftJson 'bindings.excalidraw' @{
        type='excalidraw'; elements=@($at,$edgeLabel,$frame,$a,$b,$bt,$edge,$note,$deleted,$deletedText)
    }
    $before = (Get-FileHash -LiteralPath $happyPath -Algorithm SHA256).Hash
    $graph = Invoke-Draft $happyPath
    Assert-Draft ($graph.schemaVersion -eq 1 -and $graph.sourceFormat -eq 'excalidraw') 'Versioned Excalidraw graph.'
    Assert-Draft ($graph.status -eq 'ready-for-review') 'Ready for review, not verified architecture.'
    Assert-Draft ($graph.components.Count -eq 3 -and $graph.relationships.Count -eq 1 -and $graph.annotations.Count -eq 1) 'Every active record is represented or attached.'
    Assert-Draft ($graph.counts.inputRecords -eq 10 -and $graph.counts.deletedRecords -eq 2 -and $graph.counts.attachedTexts -eq 3) 'Deleted records counted; attached text counted once.'
    Assert-Draft (($graph.components | Where-Object id -CEQ 'a').label -ceq 'API original') 'Reverse boundElements text ownership and originalText.'
    Assert-Draft (($graph.components | Where-Object id -CEQ 'b').label -ceq 'Database') 'containerId-only text ownership.'
    $actualEdge = $graph.relationships[0]
    Assert-Draft ($actualEdge.source -ceq 'a' -and $actualEdge.target -ceq 'b' -and $actualEdge.direction -eq 'forward') 'Explicit connector bindings.'
    Assert-Draft ($actualEdge.label -ceq 'request' -and $actualEdge.sourceIds.Count -eq 2 -and $actualEdge.labelParts.Count -eq 1) 'Dual-owner evidence does not duplicate bound edge label.'
    Assert-Draft ($actualEdge.confidence -eq 'explicit-binding' -and -not $actualEdge.dataFlowVerified -and $actualEdge.provenance -eq 'source-diagram') 'Diagram arrows are not verified flow.'
    Assert-Draft ($actualEdge.geometry.points.Count -eq 2 -and $actualEdge.geometry.points[1][1] -eq 20) 'Relative connector points preserved.'
    Assert-Draft ((Get-FileHash -LiteralPath $happyPath -Algorithm SHA256).Hash -eq $before) 'Source unchanged.'
    $checks.Add('Bindings, label ownership/original text, deleted records, point geometry, neutral provenance, input immutability.')

    $actualA = $graph.components | Where-Object id -CEQ 'a'
    $actualB = $graph.components | Where-Object id -CEQ 'b'
    Assert-Draft ($actualA.frameId -ceq 'frame' -and $actualA.labelParts[0].frameId -ceq 'frame') 'Explicit nested frame metadata preserved on component and text.'
    Assert-Draft ($actualA.labelParts[0].groupIds[0] -is [string] -and $actualA.labelParts[0].groupIds[0] -ceq 'explicit-group') 'Label group IDs are flat strings, not nested arrays.'
    Assert-Draft ($null -eq $actualB.frameId) 'Overlapping geometry does not invent containment.'
    Assert-Draft (($graph.components | Where-Object id -CEQ 'frame').kind -eq 'group') 'Frame is an explicit group component.'
    Assert-Draft ($graph.groups.Count -eq 1 -and $graph.groups[0].memberIds.Count -eq 3) 'Only explicit groupIds membership, including bound text.'
    Assert-Draft ($graph.annotations[0].label -ceq 'Standalone note' -and $graph.annotations[0].frameId -ceq 'frame') 'Standalone nested text retained.'
    Assert-Draft ($actualA.safeId -match '^c_[a-z0-9_]+$' -and $actualA.sourceIds -ccontains 'at') 'Safe IDs separate from preserved source IDs.'
    $checks.Add('Explicit frames/groups and nested standalone text; no geometric containment guesses.')

    $reverse = New-DraftFixtureEdge 'reverse' 'arrow' $null
    $startOnly = New-DraftFixtureEdge 'start-only' 'arrow' $null
    $startOnly.Remove('endArrowhead')
    $both = New-DraftFixtureEdge 'both' 'arrow' 'arrow'
    $none = New-DraftFixtureEdge 'none' $null $null
    $unknown = New-DraftFixtureElement 'unknown' 'arrow'
    $unknown.startBinding = @{elementId='a'}
    $unknown.endBinding = @{elementId='b'}
    $unbound = New-DraftFixtureEdge 'unbound' $null 'arrow'
    $unbound.startBinding = $null
    $unbound.endBinding = $null
    $dangling = New-DraftFixtureEdge 'dangling' $null 'arrow'
    $dangling.endBinding = @{elementId='absent'}
    $line = New-DraftFixtureElement 'line' 'line'
    $line.startArrowhead = $null; $line.endArrowhead = $null
    $graph = Invoke-Draft (Write-DraftJson 'directions.json' @{
        type='excalidraw'; elements=@((New-DraftFixtureElement 'a' 'rectangle'),(New-DraftFixtureElement 'b' 'ellipse'),$reverse,$startOnly,$both,$none,$unknown,$unbound,$dangling,$line)
    })
    $r = $graph.relationships | Where-Object id -CEQ 'reverse'
    Assert-Draft ($r.source -ceq 'b' -and $r.target -ceq 'a' -and $r.direction -eq 'reverse') 'Start-only head reverses logical source and target.'
    $r = $graph.relationships | Where-Object id -CEQ 'start-only'
    Assert-Draft ($r.direction -eq 'reverse' -and $r.source -ceq 'b') 'A sole recorded startArrowhead reverses the arrow even when endArrowhead is omitted.'
    $r = $graph.relationships | Where-Object id -CEQ 'both'
    Assert-Draft ($r.direction -eq 'bidirectional' -and $r.source -ceq 'a' -and $r.target -ceq 'b') 'Both heads preserved.'
    Assert-Draft (($graph.relationships | Where-Object id -CEQ 'none').direction -eq 'none') 'Explicit absence of heads stays undirected.'
    Assert-Draft (($graph.relationships | Where-Object id -CEQ 'unknown').direction -eq 'unknown') 'Missing head fields do not guess direction.'
    $r = $graph.relationships | Where-Object id -CEQ 'unbound'
    Assert-Draft ($null -eq $r.source -and $null -eq $r.target -and $r.confidence -eq 'unresolved') 'No nearest-neighbor glue for overlapping unbound endpoints.'
    $r = $graph.relationships | Where-Object id -CEQ 'dangling'
    Assert-Draft ($r.source -ceq 'a' -and $null -eq $r.target -and $r.endBinding.elementId -ceq 'absent') 'Dangling binding retained and unresolved.'
    Assert-Draft ($graph.relationships.Count -eq 8 -and $graph.unresolved.Count -ge 6) 'No dangling arrows or unbound lines silently dropped.'
    $reverseRoundTrip = Invoke-Draft (Write-DraftJson 'reverse-round-trip.json' $graph)
    $r = $reverseRoundTrip.relationships | Where-Object id -CEQ 'reverse'
    Assert-Draft ($r.source -ceq 'b' -and $r.target -ceq 'a' -and $r.direction -eq 'reverse') 'Normalized reverse endpoints are already arrow-origin/destination and are not reversed twice.'
    $checks.Add('Forward/reverse/bidirectional/undirected/unknown heads, dangling and unbound connectors.')

    $c1 = New-DraftFixtureElement 'c1' 'rectangle'
    $c1.boundElements = @(@{id='t';type='text'},@{id='missing-label';type='text'})
    $c2 = New-DraftFixtureElement 'c2' 'rectangle'
    $text = New-DraftFixtureElement 't' 'text'
    $text.text = 'Conflict'; $text.containerId = 'c2'
    $orphan = New-DraftFixtureElement 'orphan' 'text'
    $orphan.text = 'Orphan'; $orphan.containerId = 'missing-owner'
    $graph = Invoke-Draft (Write-DraftJson 'ambiguous.json' @{type='excalidraw';elements=@($c1,$c2,$text,$orphan)})
    Assert-Draft ($graph.annotations.Count -eq 2 -and $graph.counts.attachedTexts -eq 0) 'Conflicting and missing owners stay annotations.'
    Assert-Draft (@($graph.unresolved | Where-Object code -EQ 'UnresolvedTextBinding').Count -eq 2) 'Every unresolved text owner reported.'
    Assert-Draft (@($graph.unresolved | Where-Object code -EQ 'MissingBoundElement').Count -eq 1) 'Stale boundElements reported.'
    $checks.Add('Conflicting text ownership and stale references remain explicit, not guessed.')

    $rotated = New-DraftFixtureElement 'rotated' 'rectangle'
    $rotated.angle = 0.5
    $minimal = @{id='minimal';type='ellipse'}
    $unsupported = New-DraftFixtureElement 'future-type' 'custom-unknown-type'
    $graph = Invoke-Draft (Write-DraftJson 'geometry.json' @{type='excalidraw';elements=@($rotated,$minimal,$unsupported)})
    Assert-Draft (($graph.components | Where-Object id -CEQ 'rotated').geometry.angle -eq 0.5) 'Rotation preserved.'
    Assert-Draft (@($graph.unresolved | Where-Object code -EQ 'RotatedElementRequiresReview').Count -eq 1) 'Rotated geometry review flag.'
    Assert-Draft ($null -eq ($graph.components | Where-Object id -CEQ 'minimal').geometry.x) 'Missing geometry is null, not fabricated.'
    Assert-Draft (@($graph.warnings | Where-Object code -EQ 'IncompleteGeometry').Count -eq 1) 'Missing geometry reported.'
    Assert-Draft (($graph.components | Where-Object id -CEQ 'future-type').kind -eq 'unknown') 'Unknown source types preserved for review.'
    $checks.Add('Rotated/missing geometry and unknown source elements preserved and flagged.')

    $unicodePath = Write-DraftText 'unicode.json' '{"type":"excalidraw","elements":[{"id":"shape \u00e9/\u6771","type":"rectangle"},{"id":"txt","type":"text","containerId":"shape \u00e9/\u6771","text":"Caf\u00e9 \u6771\u4eac \ud83c\udf10"}]}'
    $graph = Invoke-Draft $unicodePath
    $expectedUnicode = 'Caf' + [char]0xe9 + ' ' + [char]0x6771 + [char]0x4eac + ' ' + [char]::ConvertFromUtf32(0x1f310)
    Assert-Draft ($graph.components[0].label -ceq $expectedUnicode) 'Unicode and non-BMP label preservation.'
    $unicodeJson = Join-Path $fixtures 'unicode-utf8.json'
    & $importer -InputPath $unicodePath -OutputPath $unicodeJson
    $outputGraph = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($unicodeJson, $utf8))
    Assert-Draft ($outputGraph.components[0].label -ceq $expectedUnicode) 'UTF8 output file preserves Unicode.'
    Assert-Draft ($outputGraph.components[0].safeId -ceq $graph.components[0].safeId) 'Stable sanitized IDs.'
    $checks.Add('UTF8/non-BMP labels, stable safe IDs, and new-file JSON output.')

    $sentinel = Join-Path $fixtures 'MUST-NOT-EXECUTE.txt'
    $payload = '$([IO.File]::WriteAllText(''' + $sentinel + ''',''executed'')) <script>alert(1)</script>'
    $image = New-DraftFixtureElement 'image' 'image'
    $image.fileId = 'svg-file'; $image.link = 'https://example.invalid/do-not-fetch'
    $literal = New-DraftFixtureElement 'literal' 'text'
    $literal.text = $payload
    $graph = Invoke-Draft (Write-DraftJson 'inert-content.excalidraw' @{
        type='excalidraw'; elements=@($image,$literal)
        files=@{'svg-file'=@{mimeType='image/svg+xml';dataURL=('<svg onload="' + $payload + '"></svg>')}}
    })
    Assert-Draft (-not (Test-Path -LiteralPath $sentinel)) 'Source content never executed.'
    Assert-Draft ($graph.annotations[0].label -ceq $payload -and $graph.components[0].link -ceq $image.link) 'Text and links remain inert original data.'
    Assert-Draft ($graph.components[0].fileId -ceq 'svg-file' -and $graph.components[0].requiresVisualReview) 'Image file ID preserved and visual review required.'
    Assert-Draft ($graph.embeddedFiles.Count -eq 1 -and $graph.embeddedFiles[0].hasEmbeddedData) 'Embedded metadata retained.'
    Assert-Draft ($null -eq $graph.embeddedFiles[0].PSObject.Properties['dataURL']) 'Embedded SVG/HTML payload excluded, never decoded.'
    $imageRoundTrip = Invoke-Draft (Write-DraftJson 'image-round-trip.json' $graph)
    Assert-Draft ($imageRoundTrip.embeddedFiles[0].id -ceq 'svg-file' -and $imageRoundTrip.components[0].fileId -ceq 'svg-file') 'Embedded source file identifiers survive normalized re-import.'
    $checks.Add('Embedded images retain file IDs; SVG/HTML, links, and executable-looking labels remain inert.')

    $manualPath = Write-DraftText 'manual.json' @'
{"schemaVersion":1,"sourceFormat":"manual-ascii","components":[
 {"id":"web/client","label":"Web client","sourceIds":["screenshot-box-1"]},
 {"id":"api","label":"API","frameId":"area"},
 {"id":"area","label":"Observed area","kind":"group"}],
 "relationships":[{"id":"r","source":"web/client","target":"api","direction":"forward","label":"HTTPS","dataFlowVerified":true},
 {"id":"d","source":"api","target":"missing","direction":"none"}],
 "annotations":[{"id":"n","label":"Unclear service"}],
 "groups":[{"id":"g","memberIds":["web/client","api"]}],
 "warnings":["Human interpretation"],"unresolved":[{"code":"Question","message":"Confirm service","sourceIds":["api"]}]}
'@
    $graph = Invoke-Draft $manualPath
    Assert-Draft ($graph.components.Count -eq 3 -and $graph.annotations.Count -eq 1 -and $graph.groups.Count -eq 1) 'Normalized graph structural import.'
    Assert-Draft ($graph.components[0].id -ceq 'web/client' -and $graph.components[0].sourceIds[0] -ceq 'screenshot-box-1') 'Manual source provenance preserved.'
    Assert-Draft ($graph.relationships[0].confidence -eq 'declared' -and -not $graph.relationships[0].dataFlowVerified) 'Manual interpretation not promoted to verification.'
    Assert-Draft (@($graph.warnings | Where-Object code -EQ 'VerificationNotImported').Count -eq 1) 'Untrusted verification claim explicitly rejected.'
    Assert-Draft ($null -eq $graph.relationships[1].target -and $graph.relationships[1].declaredTarget -ceq 'missing') 'Manual missing target remains unresolved.'
    Assert-Draft (@($graph.unresolved | Where-Object code -EQ 'Question').Count -eq 1) 'Human unresolved questions retained.'
    $checks.Add('Normalized manual graph import; provenance, annotations, explicit groups, unresolved questions, and unverified flows.')

    $happyGraph = Invoke-Draft $happyPath
    $roundTrip = Invoke-Draft (Write-DraftJson 'round-trip.json' $happyGraph)
    Assert-Draft ($roundTrip.relationships[0].startBinding.elementId -ceq 'a' -and $roundTrip.relationships[0].endArrowhead -ceq 'arrow') 'Normalized re-import retains diagram binding/head evidence.'
    Assert-Draft ($roundTrip.relationships[0].labelParts[0].label -ceq 'request' -and $roundTrip.relationships[0].provenance -eq 'source-diagram') 'Normalized re-import retains original bound labels and provenance.'
    Assert-Draft (@($roundTrip.unresolved | Where-Object code -EQ 'UnresolvedGroupMember').Count -eq 0) 'Source-text group members are recognized on re-import.'
    Assert-Draft ($roundTrip.components[1].safeId -ceq $happyGraph.components[1].safeId) 'Safe IDs stable through normalized re-import.'
    $nested = Invoke-Draft (Write-DraftText 'nested-frame.json' @'
{"type":"excalidraw","elements":[{"id":"f","type":"frame","frameId":"outer"},
{"id":"outer","type":"frame"},{"id":"ft","type":"text","text":"Frame label","containerId":"f","frameId":"f"}]}
'@)
    Assert-Draft (@($nested.unresolved | Where-Object code -EQ 'UnresolvedFrame').Count -eq 0) 'Nested frame label can explicitly belong to the frame it labels.'
    $grouped = Invoke-Draft (Write-DraftText 'groupids-only.json' '{"schemaVersion":1,"sourceFormat":"manual-image","components":[{"id":"a","label":"Box","groupIds":["g"]}]}')
    Assert-Draft ($grouped.groups[0].id -ceq 'g' -and $grouped.groups[0].memberIds[0] -ceq 'a') 'Normalized groupIds become explicit group metadata even without a groups array.'
    $checks.Add('Normalized re-import preserves bindings, labels, nested frames, image descriptors, source IDs, and explicit groups.')

    $empty = Invoke-Draft (Write-DraftText 'empty.json' '{"type":"excalidraw","elements":[]}')
    Assert-Draft ($empty.status -eq 'empty' -and $empty.components.Count -eq 0) 'Empty diagram gives meaningful empty report.'
    Assert-Draft (@($empty.unresolved | Where-Object code -EQ 'NoComponents').Count -eq 1) 'Empty report explains lack of components.'
    $empty = Invoke-Draft (Write-DraftText 'only-text.json' '{"type":"excalidraw","elements":[{"id":"t","type":"text","text":"title"}]}')
    Assert-Draft ($empty.status -eq 'empty' -and $empty.annotations.Count -eq 1) 'Text-only diagram not fabricated into architecture.'
    $empty = Invoke-Draft (Write-DraftText 'deleted-only.json' '{"type":"excalidraw","elements":[{"id":"d","isDeleted":true}]}')
    Assert-Draft ($empty.status -eq 'empty' -and $empty.counts.deletedRecords -eq 1) 'Deleted-only diagram reports empty without requiring deleted content.'
    $checks.Add('Empty, text-only, and deleted-only inputs produce meaningful reports without default topology.')

    $invalidCases = @(
        @('bad-json.json','{"type":','*Invalid JSON*'),
        @('case-conflicting-json-keys.json','{"type":"excalidraw","Type":"excalidraw","elements":[]}','*Invalid JSON*'),
        @('root-array.json','[]','*must be a JSON object*'),
        @('root-null.json','null','*must be a JSON object*'),
        @('root-number.json','1','*must be a JSON object*'),
        @('root-string.json','"hello"','*must be a JSON object*'),
        @('root-one-array.json','[{}]','*must be a JSON object*'),
        @('unknown.json','{"elements":[]}','*Unsupported JSON structure*'),
        @('missing-elements.json','{"type":"excalidraw"}','*Missing required array*'),
        @('wrong-elements.json','{"type":"excalidraw","elements":{}}','*must be a JSON array*'),
        @('missing-id.json','{"type":"excalidraw","elements":[{"type":"rectangle"}]}','*Missing required string*id*'),
        @('empty-id.json','{"type":"excalidraw","elements":[{"id":"","type":"rectangle"}]}','*must not be empty*'),
        @('missing-type.json','{"type":"excalidraw","elements":[{"id":"a"}]}','*Missing required string*type*'),
        @('duplicate.json','{"type":"excalidraw","elements":[{"id":"a","type":"rectangle"},{"id":"a","type":"ellipse"}]}','*Duplicate record ID*'),
        @('duplicate-deleted.json','{"type":"excalidraw","elements":[{"id":"a","type":"rectangle"},{"id":"a","isDeleted":true}]}','*Duplicate record ID*'),
        @('bad-number.json','{"type":"excalidraw","elements":[{"id":"a","type":"rectangle","x":"0"}]}','*finite number*'),
        @('bad-points.json','{"type":"excalidraw","elements":[{"id":"r","type":"arrow","points":[[0,"x"]]}]}','*numeric*x,y*pairs*'),
        @('bad-deleted.json','{"type":"excalidraw","elements":[{"id":"a","type":"rectangle","isDeleted":"true"}]}','*must be a boolean*'),
        @('bad-binding.json','{"type":"excalidraw","elements":[{"id":"r","type":"arrow","startBinding":{}}]}','*Missing required string*elementId*'),
        @('bad-bound-elements.json','{"type":"excalidraw","elements":[{"id":"a","type":"rectangle","boundElements":["text"]}]}','*entries must be objects*'),
        @('schema-version.json','{"schemaVersion":2,"sourceFormat":"normalized","components":[]}','*Unsupported schemaVersion*'),
        @('schema-version-string.json','{"schemaVersion":"1","sourceFormat":"normalized","components":[]}','*Unsupported schemaVersion*'),
        @('missing-components.json','{"schemaVersion":1,"sourceFormat":"normalized"}','*Missing required array*components*'),
        @('missing-label.json','{"schemaVersion":1,"sourceFormat":"normalized","components":[{"id":"a"}]}','*Missing required string*label*'),
        @('bad-kind.json','{"schemaVersion":1,"sourceFormat":"normalized","components":[{"id":"a","label":"x","kind":"azureFirewall"}]}','*Unsupported component kind*'),
        @('bad-direction.json','{"schemaVersion":1,"sourceFormat":"normalized","components":[],"relationships":[{"id":"r","direction":"traffic"}]}','*Unsupported relationship direction*'),
        @('duplicate-manual.json','{"schemaVersion":1,"sourceFormat":"normalized","components":[{"id":"a","label":"a"}],"annotations":[{"id":"a","label":"note"}]}','*Duplicate record ID*'),
        @('bad-record.json','{"schemaVersion":1,"sourceFormat":"normalized","components":[null]}','*Every record must be a JSON object*'),
        @('bad-sourceids.json','{"schemaVersion":1,"sourceFormat":"normalized","components":[{"id":"a","label":"a","sourceIds":[5]}]}','*nonempty strings*')
    )
    foreach ($case in $invalidCases) {
        $path = Write-DraftText $case[0] $case[1]
        Assert-DraftThrows { & $importer -InputPath $path } $case[2]
    }
    $invalidUtf8 = Join-Path $fixtures 'invalid-utf8.json'
    [IO.File]::WriteAllBytes($invalidUtf8, [byte[]]@(0x7b,0xc3,0x28,0x7d))
    Assert-DraftThrows { & $importer -InputPath $invalidUtf8 } '*valid Unicode JSON text*'
    $checks.Add("Clear validation errors for $($invalidCases.Count) malformed JSON/record/schema/binding cases.")

    $caseGraph = Invoke-Draft (Write-DraftText 'case-sensitive.json' '{"type":"excalidraw","elements":[{"id":"A","type":"rectangle"},{"id":"a","type":"rectangle"}]}')
    Assert-Draft ($caseGraph.components.Count -eq 2 -and $caseGraph.components[0].safeId -cne $caseGraph.components[1].safeId) 'Case-sensitive IDs do not collide after sanitization.'
    $checks.Add('Case-sensitive source identifiers and distinct stable model-safe IDs.')

    $limitRecords = @()
    for ($i=0; $i -lt 2001; $i++) { $limitRecords += @{id="e$i";type='rectangle';isDeleted=$true} }
    $recordLimit = Write-DraftJson 'record-limit.json' @{type='excalidraw';elements=$limitRecords}
    Assert-DraftThrows { & $importer -InputPath $recordLimit } '*2000 record limit*'
    $boundary = Invoke-Draft (Write-DraftJson 'record-boundary.json' @{type='excalidraw';elements=@($limitRecords[0..1999])})
    Assert-Draft ($boundary.counts.inputRecords -eq 2000 -and $boundary.counts.deletedRecords -eq 2000 -and $boundary.status -eq 'empty') 'Exactly 2000 records accepted.'
    $sizeLimit = Join-Path $fixtures 'size-limit.json'
    $sizeStream = [IO.File]::Open($sizeLimit, [IO.FileMode]::CreateNew)
    try { $sizeStream.SetLength(10 * 1024 * 1024 + 1) } finally { $sizeStream.Dispose() }
    Assert-DraftThrows { & $importer -InputPath $sizeLimit } '*10 MiB file size limit*'
    $checks.Add('10 MiB and 2000-record resource limits, including deleted records.')

    foreach ($path in @('relative.json','C:relative.json','\relative.json')) {
        Assert-DraftThrows { & $importer -InputPath $path } '*fully qualified Windows filesystem path*'
    }
    Assert-DraftThrows { & $importer -InputPath $happyPath -OutputPath 'relative-output.json' } '*fully qualified Windows filesystem path*'
    Assert-DraftThrows { & $importer -InputPath ($happyPath + ':stream') } '*fully qualified Windows filesystem path*'
    Assert-DraftThrows { & $importer -InputPath (Join-Path $fixtures 'missing.json') } '*existing JSON or Excalidraw file*'
    foreach ($name in @('image.png','diagram.svg','diagram.txt','flowchart.drawio')) {
        $path = Write-DraftText $name 'must not be parsed'
        Assert-DraftThrows { & $importer -InputPath $path } '*Interpret images, ASCII*'
    }
    Assert-DraftThrows { & $importer -InputPath $happyPath -OutputPath (Join-Path $fixtures 'missing-parent\graph.json') } '*parent directory must already exist*'
    Assert-DraftThrows { & $importer -InputPath $happyPath -OutputPath $happyPath } '*never overwrites files*'
    Assert-DraftThrows { & $importer -InputPath $happyPath -OutputPath $unicodeJson } '*never overwrites files*'
    Assert-Draft ((Get-FileHash -LiteralPath $happyPath -Algorithm SHA256).Hash -eq $before) 'Refused overwrite leaves source unchanged.'
    $checks.Add('Absolute filesystem paths, explicit unsupported-format guidance, no overwrites or input mutation.')
} finally {
    # Delete only this invocation's unique child, never the supplied output directory.
    if ([IO.Directory]::Exists($fixtures)) { Remove-Item -LiteralPath $fixtures -Recurse -Force }
}

[pscustomobject]@{
    suite='Draft-Import'
    passed=$true
    checkGroups=$checks.Count
    checks=@($checks.ToArray())
    fixtureDirectoryRemoved=(-not (Test-Path -LiteralPath $fixtures))
}

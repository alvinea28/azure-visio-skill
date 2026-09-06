[CmdletBinding()]
param([Parameter(Mandatory)][string]$OutputDirectory)
Set-StrictMode -Version 2
$ErrorActionPreference='Stop'
if ($OutputDirectory -notmatch '^[A-Za-z]:\\' -or $OutputDirectory.Substring(3) -match '[\x00-\x1f<>:"/|?*]' -or
    $OutputDirectory -match '(?:^|\\)\.\.?(?:\\|$)') { throw 'OutputDirectory must be an absolute local directory without traversal or streams.' }
[void][IO.Directory]::CreateDirectory($OutputDirectory)
$fixtures=Join-Path $OutputDirectory ('Reference-Fidelity-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($fixtures)
$helper=Join-Path $PSScriptRoot 'Reference-Fidelity.ps1'
$utf8=[Text.UTF8Encoding]::new($false)
$checks=[Collections.Generic.List[string]]::new()
$assertions=0
function Assert([bool]$Condition, [string]$Message) {
    $script:assertions++
    if (-not $Condition) { throw "REFERENCE ASSERTION FAILED: $Message" }
}
function WriteJson([string]$Name, $Object) {
    $path=Join-Path $fixtures $Name
    [IO.File]::WriteAllText($path,(ConvertTo-Json -InputObject $Object -Depth 40),$utf8)
    return $path
}
function WriteText([string]$Name, [string]$Text) {
    $path=Join-Path $fixtures $Name; [IO.File]::WriteAllText($path,$Text,$utf8); return $path
}
function CopyJson($Object) { return (ConvertFrom-Json -InputObject (ConvertTo-Json -InputObject $Object -Depth 40)) }
function Gate([string]$Model=$modelPath, [string]$Contract=$contractPath) {
    return (ConvertFrom-Json -InputObject ((& $helper -ModelPath $Model -ReferencePath $Contract -ReportOnly) -join "`n"))
}
function Reject([string]$Model=$modelPath, [string]$Contract=$contractPath, [string[]]$Codes=@()) {
    $result=Gate $Model $Contract
    Assert ($result.valid -is [bool] -and -not $result.valid) 'Every negative report has the boolean valid:false.'
    Assert ($result.errors.Count -gt 0) 'Every rejection explains the failure.'
    foreach ($code in $Codes) { Assert (@($result.errors | Where-Object code -CEQ $code).Count -gt 0) "Expected $code; received $($result.errors | ConvertTo-Json -Compress)" }
}
function Node([string]$Id, [string]$Kind, [string]$Label, $X, $Y, $Width, $Height, [string]$Parent='') {
    return @{id="n-$Id";sourceId="src-$Id";kind=$Kind;label=$Label;x=$X;y=$Y;width=$Width;height=$Height;parent=$Parent}
}
try {
    foreach ($path in @($helper,$PSCommandPath)) {
        $tokens=$null; $errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
        Assert ($errors.Count -eq 0) "Windows PowerShell parsing: $path"
        if ($path -eq $helper) {
            $unsafe=@($ast.FindAll({
                param($node)
                $node -is [Management.Automation.Language.CommandAst] -and
                $node.GetCommandName() -in @('Invoke-Expression','iex','Invoke-WebRequest','Invoke-RestMethod','Start-Process','curl','Add-Type','New-Object')
            },$true))
            Assert ($unsafe.Count -eq 0) 'No COM, content execution, network, or general-purpose execution commands.'
        }
    }
    $checks.Add('PowerShell 5.1 parsing; filesystem-only helper without COM/network/content execution.')
    $sourcePath=Join-Path $fixtures 'source.bin'
    $sourceBytes=$utf8.GetBytes('Synthetic source bytes, NOT image recognition. $(throw "must stay inert")')
    [IO.File]::WriteAllBytes($sourcePath,$sourceBytes)
    $sourceHash=(Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
    $model=@{
        schemaVersion=1;title='Synthetic faithful fixture';pages=@(@{
            name='01 Reference architecture';width=20;height=10;furniture=$false
            nodes=@(
                (Node 'governance' 'card' "Source governance`nPolicy and audit" 10 9 18 1),
                (Node 'lane' 'container' 'Application boundary' 5 5 8 6),
                (Node 'api' 'card' "API tier`nOriginal processing detail" 5 6.5 6 1 'n-lane'),
                (Node 'db' 'card' "Data store`nOriginal persistence detail" 5 3.5 6 1 'n-lane'),
                (Node 'note' 'note' "Source note`nDiagram arrow is conceptual" 15 5 7 2),
                (Node 'identity' 'card' "Identity layer`nOriginal identity detail" 15 7.5 7 1)
            )
            edges=@(
                @{id='e-api-db';sourceId='src-api-db';source='n-api';target='n-db';direction='forward';kind='association';label='Conceptual linkage only'},
                @{id='e-identity-api';sourceId='src-identity-api';source='n-identity';target='n-api';direction='none';kind='association';label=''}
            )
        })
    }
    $contract=@{
        schemaVersion=1;mode='source-faithful';source=@{path='source.bin';sha256=$sourceHash;role='authoritative-reference'}
        referencePage='01 Reference architecture';unresolved=@()
        components=@(
            @{id='src-governance';label='Source governance';requiredText=@('Source governance','Policy and audit');parent=$null;kind='card'},
            @{id='src-lane';label='Application boundary';requiredText=@();parent=$null;kind='container'},
            @{id='src-api';label='API tier';requiredText=@('Original processing detail');parent='src-lane';kind='card'},
            @{id='src-db';label='Data store';requiredText=@('Original persistence detail');parent='src-lane';kind='card'},
            @{id='src-note';label='Source note';requiredText=@('Diagram arrow is conceptual');parent=$null;kind='note'},
            @{id='src-identity';label='Identity layer';requiredText=@('Original identity detail');parent=$null;kind='card'}
        )
        relationships=@(
            @{id='src-api-db';source='src-api';target='src-db';direction='forward';requiredText=@('Conceptual linkage only')},
            @{id='src-identity-api';source='src-identity';target='src-api';direction='none'}
        )
        layout=@{leftToRight=@(,@('src-lane','src-note'));topToBottom=@(@('src-governance','src-lane'),@('src-api','src-db'),@('src-identity','src-note'));aspectRatio=2}
    }
    $modelPath=WriteJson 'model.json' $model; $contractPath=WriteJson 'contract.json' $contract
    $baselineModelHash=(Get-FileHash -LiteralPath $modelPath -Algorithm SHA256).Hash
    $baselineContractHash=(Get-FileHash -LiteralPath $contractPath -Algorithm SHA256).Hash
    $result=Gate
    Assert $result.valid "Faithful fixture passes: $($result.errors | ConvertTo-Json -Compress)"
    Assert ($result.source.verified -and $result.source.actualSha256 -ieq $sourceHash) 'Persisted bytes verified against separate contract.'
    Assert ($result.counts.expectedComponents -eq 6 -and $result.counts.matchedComponents -eq 6 -and $result.counts.expectedRelationships -eq 2 -and $result.counts.matchedRelationships -eq 2) 'Exact node/edge counts.'
    Assert (-not $result.comUsed -and -not $result.imageRecognitionPerformed -and -not $result.dataFlowVerified) 'A passing contract check does not invent image recognition or verified flow.'
    $defaultResult=ConvertFrom-Json -InputObject ((& $helper -ModelPath $modelPath -ReferencePath $contractPath) -join "`n")
    Assert $defaultResult.valid 'Default invocation returns JSON for success.'
    $checks.Add('Faithful synthetic source fixture, hash, explicit provenance limits, exact counts, default invocation.')

    $m=CopyJson $model
    $m.pages[0].nodes[2].label="  api    TIER `r`n ORIGINAL   PROCESSING detail "
    Assert (Gate (WriteJson 'normalized.json' $m)).valid 'Whitespace/case normalization preserves complete lines.'
    $c=CopyJson $contract; $c.source.path=$sourcePath
    Assert (Gate $modelPath (WriteJson 'absolute-source.json' $c)).valid 'Absolute local source accepted.'
    [void][IO.Directory]::CreateDirectory((Join-Path $fixtures 'source-child'))
    [IO.File]::WriteAllBytes((Join-Path $fixtures 'source-child\source.bin'),$sourceBytes)
    $c.source.path='source-child\source.bin'
    Assert (Gate $modelPath (WriteJson 'contained-source.json' $c)).valid 'Contained relative source accepted.'
    $checks.Add('Exact normalized text and absolute/contained-relative source paths.')

    $m=CopyJson $model
    $m.pages[0].nodes=@($m.pages[0].nodes | Where-Object sourceId -CNotIn @('src-governance','src-identity','src-note'))
    $m.pages[0].nodes+=Node 'custom-bridge' 'card' 'Custom bridge' 15 5 7 2
    $m.pages[0].nodes[1].y=3.5; $m.pages[0].nodes[2].y=6.5
    $outputLike=WriteJson 'output-like-redesign.json' $m
    Reject $outputLike $contractPath @('MissingItem','UnexpectedItem','LayoutMismatch')
    $threw=$false
    try { & $helper -ModelPath $outputLike -ReferencePath $contractPath | Out-Null }
    catch { $threw=$_.Exception.Message -like 'Reference fidelity failed:*' }
    Assert $threw 'Default invocation throws, rather than accepting an output-like redesign.'
    $m | Add-Member NoteProperty valid $true
    Reject (WriteJson 'forged-valid-flag.json' $m) $contractPath @('MissingItem')
    $checks.Add('Output-like substitution drops original services/governance, adds bridge, changes layout; throws and cannot forge valid.')

    $c=CopyJson $contract; $c.source.path='missing.bin'
    Reject $modelPath (WriteJson 'missing-source.json' $c) @('MissingSource')
    [IO.File]::WriteAllText($sourcePath,'mutated source',$utf8)
    Reject $modelPath $contractPath @('SourceHashMismatch')
    [IO.File]::WriteAllBytes($sourcePath,$sourceBytes)
    $c=CopyJson $contract; $c.source.role='output-regression'
    Reject $modelPath (WriteJson 'output-is-not-source.json' $c) @('SourceRole')
    $c=CopyJson $contract; $c.source.path=$modelPath
    Reject $modelPath (WriteJson 'model-is-not-source.json' $c) @('SourceRole')
    $checks.Add('Missing/changed source, wrong source role, and using model as source fail without fallback.')

    foreach ($variant in @('duplicate-source','duplicate-native','missing-source-id','case-source-id','annotation-escape')) {
        $m=CopyJson $model
        switch ($variant) {
            'duplicate-source' { $m.pages[0].nodes[1].sourceId='src-api' }
            'duplicate-native' { $m.pages[0].nodes[1].id=$m.pages[0].nodes[0].id }
            'missing-source-id' { $m.pages[0].nodes[1].PSObject.Properties.Remove('sourceId') }
            'case-source-id' { $m.pages[0].nodes[1].sourceId='SRC-LANE' }
            'annotation-escape' { $extra=Node 'extra' 'note' 'New service' 15 2 6 1; $extra.annotation=$true; $m.pages[0].nodes+=$extra }
        }
        Reject (WriteJson ($variant+'.json') $m)
    }
    $c=CopyJson $contract; $c.components[1].id='src-api'
    Reject $modelPath (WriteJson 'duplicate-contract-id.json' $c) @('DuplicateSourceId')
    $m=CopyJson $model; $m.pages[0].edges[1].sourceId='src-api-db'
    Reject (WriteJson 'duplicate-edge-source.json' $m) $contractPath @('DuplicateSourceId','MissingItem')
    $checks.Add('Duplicate/missing/case-changed source IDs, duplicate native IDs, and annotation escape rejected.')

    $m=CopyJson $model; $m.pages[0].name='02 Proposal'
    Reject (WriteJson 'renamed-reference-page.json' $m) $contractPath @('MissingReferencePage')
    $m=CopyJson $model; $extra=CopyJson $m.pages[0]; $extra.name='02 Proposal'; $m.pages+=@($extra)
    $extraPath=WriteJson 'supplemental-pages.json' $m
    Reject $extraPath $contractPath @('AdditionalPages')
    $c=CopyJson $contract; $c | Add-Member NoteProperty allowAdditionalPages $true
    $allowedPath=WriteJson 'allow-additional-pages.json' $c
    Assert (Gate $extraPath $allowedPath).valid 'Explicitly allowed supplemental page leaves faithful reference intact.'
    $m.pages[0].nodes=@(); $m.pages[0].edges=@()
    Reject (WriteJson 'source-moved-to-supplement.json' $m) $allowedPath @('MissingItem')
    $checks.Add('Designated page required; supplemental pages require permission and cannot replace source page.')

    foreach ($variant in @('reordered-layers','horizontal-reorder','overlap','touching','aspect','governance','hierarchy','renamed','title-suffix','lost-detail','detail-substring','extra-detail','duplicate-line','kind')) {
        $m=CopyJson $model
        switch ($variant) {
            'reordered-layers' { $m.pages[0].nodes[2].y=3.5; $m.pages[0].nodes[3].y=6.5 }
            'horizontal-reorder' { $m.pages[0].nodes[4].x=1 }
            'overlap' { $m.pages[0].nodes[3].y=6.5 }
            'touching' { $m.pages[0].nodes[3].y=5.5 }
            'aspect' { $m.pages[0].width=40 }
            'governance' { $m.pages[0].nodes=@($m.pages[0].nodes | Where-Object sourceId -CNE 'src-governance') }
            'hierarchy' { $m.pages[0].nodes[2].parent='' }
            'renamed' { $m.pages[0].nodes[2].label="Replacement service`nOriginal processing detail" }
            'title-suffix' { $m.pages[0].nodes[2].label="API tier alternative`nOriginal processing detail" }
            'lost-detail' { $m.pages[0].nodes[2].label='API tier' }
            'detail-substring' { $m.pages[0].nodes[2].label="API tier`nNOT Original processing detail" }
            'extra-detail' { $m.pages[0].nodes[2].label+="`nAlso a new service" }
            'duplicate-line' { $m.pages[0].nodes[2].label+="`nOriginal processing detail" }
            'kind' { $m.pages[0].nodes[2].kind='note' }
        }
        Reject (WriteJson ($variant+'.json') $m)
    }
    $checks.Add('Layer/lane order, overlap/touching, aspect, governance, hierarchy, title/details, duplicate lines and kinds.')
    $c=CopyJson $contract; $m=CopyJson $model
    $c.components+=@([pscustomobject]@{id='src-inner';label='Inner boundary';requiredText=@();parent='src-lane';kind='container'})
    $m.pages[0].nodes+=Node 'inner' 'container' 'Inner boundary' 5 5 6 4 'n-lane'
    $c.components[2].parent='src-inner'; $c.components[3].parent='src-inner'
    $m.pages[0].nodes[2].parent='n-inner'; $m.pages[0].nodes[3].parent='n-inner'
    $nestedContract=WriteJson 'nested-contract.json' $c
    Assert (Gate (WriteJson 'nested-model.json' $m) $nestedContract).valid 'Nested source containers map through native parent IDs.'
    $m.pages[0].nodes[2].parent='n-lane'
    Reject (WriteJson 'flattened-hierarchy.json' $m) $nestedContract @('ParentMismatch')
    foreach ($variant in @('reversed','swapped-endpoints','missing-edge','extra-edge','missing-direction','lost-edge-label','edge-substring')) {
        $m=CopyJson $model
        switch ($variant) {
            'reversed' { $m.pages[0].edges[0].direction='backward' }
            'swapped-endpoints' { $m.pages[0].edges[0].source='n-db'; $m.pages[0].edges[0].target='n-api' }
            'missing-edge' { $m.pages[0].edges=@($m.pages[0].edges[1]) }
            'extra-edge' { $edge=CopyJson $m.pages[0].edges[0]; $edge.id='new-edge'; $edge.sourceId='new-source'; $m.pages[0].edges+=@($edge) }
            'missing-direction' { $m.pages[0].edges[0].PSObject.Properties.Remove('direction') }
            'lost-edge-label' { $m.pages[0].edges[0].label='' }
            'edge-substring' { $m.pages[0].edges[0].label='Not Conceptual linkage only' }
        }
        Reject (WriteJson ($variant+'.json') $m)
    }
    foreach ($direction in @('forward','backward','both','none')) {
        $c=CopyJson $contract; $m=CopyJson $model
        $c.relationships[0].direction=$direction; $m.pages[0].edges[0].direction=$direction
        Assert (Gate (WriteJson 'direction-model.json' $m) (WriteJson 'direction-contract.json' $c)).valid "Explicit $direction stays a conceptual source arrow."
    }
    $checks.Add('Relationship identity, endpoint orientation, labels and all four explicit arrow directions.')

    foreach ($variant in @('unknown-field','unknown-source-field','unresolved','mode','boolean-string','version-string','components-object','details-string','bad-parent','parent-cycle','unknown-layout','flat-layout','unknown-endpoint','bad-hash')) {
        $c=CopyJson $contract
        switch ($variant) {
            'unknown-field' { $c | Add-Member NoteProperty allowExtraPages $true }
            'unknown-source-field' { $c.source | Add-Member NoteProperty downloadUrl 'https://example.invalid' }
            'unresolved' { $c.unresolved=@('Unreadable source text') }
            'mode' { $c.mode='requirements-only' }
            'boolean-string' { $c | Add-Member NoteProperty allowAdditionalPages 'true' }
            'version-string' { $c.schemaVersion='1' }
            'components-object' { $c.components=$c.components[0] }
            'details-string' { $c.components[0].requiredText='Policy and audit' }
            'bad-parent' { $c.components[2].parent='src-note' }
            'parent-cycle' { $c.components[1].parent='src-lane' }
            'unknown-layout' { $c.layout.leftToRight=@(,@('absent','src-lane')) }
            'flat-layout' { $c.layout.leftToRight=@('src-lane','src-note') }
            'unknown-endpoint' { $c.relationships[0].target='absent' }
            'bad-hash' { $c.source.sha256='not-a-sha256' }
        }
        Reject $modelPath (WriteJson ($variant+'.json') $c)
    }
    foreach ($variant in @('number-string','label-object','id-array','pages-object','null-source-id','schema-bool')) {
        $m=CopyJson $model
        switch ($variant) {
            'number-string' { $m.pages[0].nodes[0].width='18' }
            'label-object' { $m.pages[0].nodes[0].label=@{text='Source governance'} }
            'id-array' { $m.pages[0].nodes[0].id=@('n-governance') }
            'pages-object' { $m.pages=$m.pages[0] }
            'null-source-id' { $m.pages[0].nodes[0].sourceId=$null }
            'schema-bool' { $m.schemaVersion=$true }
        }
        Reject (WriteJson ($variant+'.json') $m)
    }
    foreach ($json in @('null','1','true','[]',('['+[IO.File]::ReadAllText($modelPath)+']'),'{"schemaVersion":1,"schemaVersion":1}','{"schemaVersion":1,"SCHEMAVERSION":1}','{"schemaVersion":1,"schema\u0056ersion":1}')) {
        Reject (WriteText 'bad-json.json' $json)
    }
    $checks.Add('Strict contract fields, unresolved extraction, typed JSON, duplicate properties including escaped/case variants, root arrays.')

    foreach ($path in @('https://example.invalid/source.png','\\server\share\source.png','..\source.bin','source-child\..\source.bin','source.bin:stream','C:source.bin','\\?\C:\source.bin','NUL','source.bin.','source.bin ')) {
        $c=CopyJson $contract; $c.source.path=$path
        Reject $modelPath (WriteJson 'unsafe-source.json' $c) @('UnsafePath')
    }
    Reject 'relative-model.json' $contractPath @('UnsafePath')
    Reject $modelPath 'relative-contract.json' @('UnsafePath')
    $junction=Join-Path $fixtures 'redirect'
    [void](New-Item -ItemType Junction -Path $junction -Value (Join-Path $fixtures 'source-child'))
    try {
        $c=CopyJson $contract; $c.source.path='redirect\source.bin'
        Reject $modelPath (WriteJson 'junction-source.json' $c) @('UnsafePath')
    } finally {
        # Remove the link itself, never recursively traverse its destination.
        [IO.Directory]::Delete($junction)
    }
    $sentinel=Join-Path $fixtures 'must-not-execute.txt'
    $payload='$(Set-Content -LiteralPath "'+$sentinel+'" -Value executed); https://example.invalid'
    $c=CopyJson $contract; $m=CopyJson $model
    $c.components[4].label=$payload; $m.pages[0].nodes[4].label=$payload+"`nDiagram arrow is conceptual"
    Assert (Gate (WriteJson 'inert-model.json' $m) (WriteJson 'inert-contract.json' $c)).valid 'Code/URL strings are compared only as text.'
    Assert (-not [IO.File]::Exists($sentinel)) 'Label and source strings were not executed.'
    $checks.Add('URL/UNC/device/traversal/ADS/ambiguous paths rejected; source and label strings remain inert.')

    $limitPath=Join-Path $fixtures 'limit.json'
    $stream=[IO.File]::Open($limitPath,[IO.FileMode]::CreateNew)
    try { $stream.SetLength(10MB+1) } finally { $stream.Dispose() }
    Reject $limitPath $contractPath @('InputLimit')
    Reject $modelPath $limitPath @('InputLimit')
    Reject (WriteText 'deep.json' (('['*65)+'0'+(']'*65))) $contractPath @('InputLimit')
    $invalidUtf8=Join-Path $fixtures 'invalid-utf8.json'
    [IO.File]::WriteAllBytes($invalidUtf8,[byte[]]@(0x7b,0xc3,0x28,0x7d))
    Reject $invalidUtf8 $contractPath @('InvalidInput')
    $c=CopyJson $contract; $c.components=@($c.components[0])*2001
    Reject $modelPath (WriteJson 'record-limit.json' $c) @('InvalidInput')
    $largeSource=Join-Path $fixtures 'large-source.bin'
    $stream=[IO.File]::Open($largeSource,[IO.FileMode]::CreateNew)
    try { $stream.SetLength(100MB+1) } finally { $stream.Dispose() }
    $c=CopyJson $contract; $c.source.path='large-source.bin'
    Reject $modelPath (WriteJson 'large-source.json' $c) @('InputLimit')
    [IO.File]::WriteAllBytes((Join-Path $fixtures 'empty.bin'),[byte[]]@())
    $c=CopyJson $contract; $c.source.path='empty.bin'
    Reject $modelPath (WriteJson 'empty-source.json' $c) @('InputLimit')
    Assert ((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -ceq $sourceHash) 'Persisted source remains unchanged.'
    Assert ((Get-FileHash -LiteralPath $modelPath -Algorithm SHA256).Hash -ceq $baselineModelHash) 'Model input remains unchanged.'
    Assert ((Get-FileHash -LiteralPath $contractPath -Algorithm SHA256).Hash -ceq $baselineContractHash) 'Contract input remains unchanged.'
    $checks.Add('JSON size/depth/record bounds, empty source, and input immutability.')
} finally {
    if ([IO.Directory]::Exists($fixtures)) { Remove-Item -LiteralPath $fixtures -Recurse -Force }
}
[pscustomobject]@{passed=$true;checkGroups=$checks.Count;checks=@($checks.ToArray());assertions=$assertions;fixtureDirectoryRemoved=(-not [IO.Directory]::Exists($fixtures))}

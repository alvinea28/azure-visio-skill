<#
.SYNOPSIS
Checks local presentation guardrails, not architecture correctness or Microsoft certification.
.DESCRIPTION
Requires a schemaVersion=1 enterprise model or an enterprise/reference v1.6 pack.
Diagram pages use icon-led cards and short captions. Only contract-designated
notes pages may hold detailed explanations. These thresholds are local
skill defaults, not measured Microsoft standards.
Full canonical label/details are retained as semantic data; only the visible
caption is counted for icon/label styles. No source content is read or changed.
#>
[CmdletBinding(DefaultParameterSetName='File')]
param(
    [Parameter(Mandatory,ParameterSetName='File')][string]$ModelPath,
    [Parameter(Mandatory,ParameterSetName='Object')][pscustomobject]$InputModel,
    [switch]$ReportOnly
)
Set-StrictMode -Version 2
$ErrorActionPreference='Stop'
if ($PSCmdlet.ParameterSetName -eq 'File') {
    if($ModelPath -notmatch '^[A-Za-z]:\\' -or [IO.Path]::GetExtension($ModelPath) -ine '.json'){
        throw 'Use an absolute local JSON ModelPath.'
    }
    $file=Get-Item -LiteralPath $ModelPath
    if($file.Length -gt 10MB){throw 'Style input exceeds 10 MB.'}
    $model=Get-Content -LiteralPath $ModelPath -Raw | ConvertFrom-Json
} else {
    $model=$InputModel
}
$errors=[Collections.Generic.List[object]]::new()
$metrics=[Collections.Generic.List[object]]::new()
function Value($Object,[string]$Name,$Default=$null){
    if($null -eq $Object){return $Default}
    $p=$Object.PSObject.Properties[$Name]
    if($null -eq $p){return $Default}
    if($p.Value -is [array]){return ,$p.Value}
    return $p.Value
}
function Issue([string]$Code,[string]$Page,[string]$Id,[string]$Message){
    $errors.Add([pscustomobject]@{code=$Code;page=$Page;id=$Id;message=$Message})
}
function Words([string]$Text){return [regex]::Matches($Text,'[\p{L}\p{N}]+(?:[-/][\p{L}\p{N}]+)*').Count}
function CardStyle($Node){
    $style=Value $Node 'cardStyle'
    if($null -ne $style){return $style}
    if((Value $Node 'icon' '') -or (Value $Node 'iconRef' '')){return 'icon'}
    return 'label'
}
function Caption($Node){
    $style=$(if((Value $Node 'kind') -eq 'card'){CardStyle $Node}else{'standard'})
    if($style -in @('icon','label')){
        $caption=Value $Node 'displayLabel'
        if($null -ne $caption){return [string]$caption}
        return ([string](Value $Node 'label' '') -split '\r\n|\n|\r')[0]
    }
    return [string](Value $Node 'label' '')
}
$hasContract=$null -ne $model.PSObject.Properties['outputContract']
if((Value $model 'schemaVersion') -ne 1 -or
    ((Value $model 'presentationProfile') -cne 'enterprise' -and -not ($hasContract -and (Value $model 'presentationProfile') -ceq 'reference'))){
    throw 'Style audit requires schemaVersion=1 and presentationProfile=enterprise (or reference with outputContract).'
}
if ($hasContract) {
    if (-not (Get-Command Confirm-ArchitecturePack -ErrorAction SilentlyContinue)) {
        $tokens=$null; $parseErrors=$null
        $controller=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'AzureVisio.ps1'),[ref]$tokens,[ref]$parseErrors)
        if ($parseErrors.Count) { throw 'Cannot load architecture pack validator.' }
        foreach ($function in $controller.FindAll({
            param($node)
            $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -in @(
                'Get-Value','Test-Field','Confirm-Number','Confirm-Fields','Confirm-Metadata','Confirm-ArchitecturePack')
        },$true)) { . ([scriptblock]::Create($function.Extent.Text)) }
    }
    try { Confirm-ArchitecturePack $model -Required }
    catch { Issue 'ArchitecturePack' '' '' $_.Exception.Message }
}
$pages=Value $model 'pages'
if($pages -isnot [array] -or $pages.Count -eq 0){throw 'Style model requires pages.'}
$diagramCount=0
$boundaryTypes=@('cloud','tenant','region','subscription','resource-group','vnet','subnet',
    'availability-zone','cluster','environment','system','trust','control-plane','functional')
foreach($page in $pages){
    $name=[string](Value $page 'name' '')
    $role=Value $page 'role' 'diagram'
    if($role -notin @('diagram','notes')){Issue 'PageRole' $name '' 'Page role must be diagram or notes.';continue}
    if($role -eq 'notes'){
        if(-not $hasContract){Issue 'NotesContract' $name '' 'Notes-page exemptions require the architecture-pack-v1.6 output contract.'}
        continue
    }
    $diagramCount++
    $nodes=Value $page 'nodes';$edges=Value $page 'edges'
    if($nodes -isnot [array] -or $edges -isnot [array]){throw 'Each diagram requires nodes and edges arrays.'}
    $cards=@($nodes | Where-Object {$_.kind -eq 'card'})
    $icons=@($cards | Where-Object {(CardStyle $_) -eq 'icon'})
    $boxed=@($cards | Where-Object {(CardStyle $_) -in @('standard','detail')})
    $notes=@($nodes | Where-Object {$_.kind -eq 'note'})
    if($cards.Count -lt 2){Issue 'NoArchitectureEntities' $name '' 'A diagram needs at least two architectural entities; explanations belong on notes pages.'}
    $iconRatio=0.0;$boxRatio=0.0
    if($cards.Count){
        $iconRatio=$icons.Count/[double]$cards.Count
        $boxRatio=$boxed.Count/[double]$cards.Count
        if($iconRatio -lt 0.6){Issue 'IconFirst' $name '' 'At least 60% of entities should be icon-led; do not invent icons to reach the threshold.'}
        if($boxRatio -gt 0.2){Issue 'TooManyCards' $name '' 'At most 20% of entities may be boxed process/detail cards. Use symbols and meaningful boundaries.'}
    }
    if($notes.Count -gt 3){Issue 'TooManyCallouts' $name '' 'Move detailed commentary to notes; keep at most three short diagram callouts.'}
    $total=Words ([string](Value $page 'title' ''))
    $total+=Words ([string](Value $page 'subtitle' ''))
    $total+=Words ([string](Value $page 'footer' ''))
    foreach($node in $nodes){
        $id=[string](Value $node 'id' '')
        $visible=Caption $node
        $wordCount=Words $visible
        $total+=$wordCount
        if($node.kind -eq 'container'){
            if((Value $node 'boundaryType') -notin $boundaryTypes){
                Issue 'BoundaryMeaning' $name $id 'A boundary needs a real ownership, deployment, trust, environment, or functional scope.'
            }
            if($wordCount -gt 10){Issue 'BoundaryCaption' $name $id 'Use a short boundary name, not an explanatory sentence.'}
            continue
        }
        if($node.kind -eq 'note'){
            if($wordCount -gt 18){Issue 'CalloutLength' $name $id 'Use at most 18 words in a main-diagram callout; retain detail in notes.'}
            continue
        }
        if($node.kind -ne 'card'){Issue 'NodeKind' $name $id 'Unsupported diagram node kind.';continue}
        $style=CardStyle $node
        if($style -notin @('icon','label','standard','detail')){Issue 'CardStyle' $name $id 'Unknown card style.'}
        if($style -eq 'icon' -and -not ((Value $node 'icon' '') -or (Value $node 'iconRef' ''))){
            Issue 'MissingGlyph' $name $id 'An icon node requires an appropriate catalog or builtin icon.'
        }
        $lines=@($visible -split '\r\n|\n|\r' | Where-Object {-not [string]::IsNullOrWhiteSpace($_)})
        if([string]::IsNullOrWhiteSpace($visible) -or $wordCount -gt 8 -or $lines.Count -gt 3 -or $visible.Length -gt 80){
            Issue 'EntityCaption' $name $id 'Use one to three short caption lines, no more than eight words/80 characters.'
        }
        if($visible -match '(?m)^\s*(?:[-*\u2022]|\d+[.)])\s+'){
            Issue 'BulletBlock' $name $id 'Do not put implementation checklists in primary entity captions.'
        }
    }
    foreach($edge in $edges){
        $caption=[string](Value $edge 'label' '')
        $count=Words $caption;$total+=$count
        if($count -gt 6){Issue 'EdgeCaption' $name ([string]$edge.id) 'Use short relationship labels or keyed numbers, not sentences.'}
    }
    if($total -gt 300){Issue 'TextDensity' $name '' 'Main diagram exceeds 300 visible words; split useful views or move details to notes.'}
    $metrics.Add([pscustomobject]@{
        page=$name;entities=$cards.Count;iconEntities=$icons.Count;boxedEntities=$boxed.Count
        boundaries=@($nodes | Where-Object {$_.kind -eq 'container'}).Count
        iconRatio=$iconRatio;boxedRatio=$boxRatio;visibleWords=$total
    })
}
if($diagramCount -eq 0){Issue 'MissingDiagram' '' '' 'An enterprise model must contain a diagram page, not only notes.'}
$report=[pscustomobject]@{
    valid=($errors.Count -eq 0);profile=(Value $model 'presentationProfile');metrics=@($metrics.ToArray());errors=@($errors.ToArray())
    routingPolicy='orthogonal-v1.6'
    routingNote='New normalizes authored straight routes to orthogonal routing, retaining declared endpoints and ports. Icon bottom ports attach below the fitted caption; other icon ports touch the glyph. Native Inspect reports attachment regions and checks actual strokes against every service caption, including endpoint captions.'
    note='Local presentation guardrails only; visual review and architecture validation remain required.'
}
$report | ConvertTo-Json -Depth 6
if(-not $report.valid -and -not $ReportOnly){
    throw ('Enterprise presentation failed: '+(($errors | ForEach-Object {$_.code+': '+$_.message}) -join '; '))
}

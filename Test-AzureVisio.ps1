[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$OutputDirectory,
    [switch]$PortableOnly,
    [switch]$CrudOnly,
    [switch]$EnterpriseOnly,
    [string]$IconDirectory,
    [string]$ReferenceModelPath,
    [string]$ReferencePath
)
Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
if ($EnterpriseOnly -and ($PortableOnly -or $CrudOnly)) { throw 'EnterpriseOnly is a separate native opt-in; do not combine it with PortableOnly or CrudOnly.' }
if ($EnterpriseOnly -and -not $IconDirectory) { throw 'EnterpriseOnly requires -IconDirectory with audited official SVG assets.' }
if ([bool]$ReferenceModelPath -ne [bool]$ReferencePath) { throw 'Supply both ReferenceModelPath and ReferencePath for a real-reference regression.' }
if ($OutputDirectory -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+(?:\\|$))' -or (Test-Path -LiteralPath $OutputDirectory)) {
    throw 'Provide a new absolute OutputDirectory. Tests never reuse existing output.'
}
[void][IO.Directory]::CreateDirectory($OutputDirectory)
$helper = Join-Path $PSScriptRoot 'AzureVisio.ps1'
$baseline = Join-Path $PSScriptRoot 'hub-spoke-reference.json'
$expanded = Join-Path $OutputDirectory 'reference-3-spokes.json'
$drawing = Join-Path $OutputDirectory ('Azure-Hub-Spoke-Expanded-' + [guid]::NewGuid().ToString('N') + '.vsdx')
$checks = [System.Collections.Generic.List[string]]::new()
$requestedIconDirectory=$IconDirectory
$script:helperIconDirectory=''
$script:enterpriseStylePath=Join-Path $PSScriptRoot 'Enterprise-Style.ps1'

function Assert([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
}
function Run-Helper([hashtable]$Parameters) {
    if ($script:helperIconDirectory -and -not $Parameters.ContainsKey('IconDirectory')) { $Parameters.IconDirectory=$script:helperIconDirectory }
    $text = & $helper @Parameters
    return ($text | ConvertFrom-Json)
}
function Assert-Throws([scriptblock]$Operation, [string]$Pattern) {
    $blocked = $false
    try { & $Operation | Out-Null }
    catch [System.Management.Automation.RuntimeException] {
        if ($_.Exception.Message -notlike $Pattern) { throw }
        $blocked = $true
    }
    Assert $blocked "Expected failure: $Pattern"
}

$ast = $null
foreach ($scriptFile in Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File) {
    $tokens = $null; $parseErrors = $null
    $parsed = [Management.Automation.Language.Parser]::ParseFile($scriptFile.FullName, [ref]$tokens, [ref]$parseErrors)
    Assert ($parseErrors.Count -eq 0) "PowerShell syntax: $($scriptFile.Name): $parseErrors"
    if ($scriptFile.FullName -eq $helper) { $ast = $parsed }
}
Assert ($null -ne $ast) 'Controller parsed.'
$checks.Add('All companion PowerShell scripts parse.')

$pureFunctions = @(
    'Get-Value','Get-AbsolutePath','Read-Environment','Get-CloudPath','Get-OpenDocument','Resolve-StencilPath','Connect-Visio',
    'Test-Field','Confirm-Number','Confirm-Identifier','Confirm-Fields','Confirm-Metadata','Confirm-Model',
    'Confirm-ArchitecturePack','Confirm-PackMerge','Confirm-NativeArchitecturePack',
    'Confirm-EdgeSemantics','Confirm-TargetRef','Confirm-Changes','Resolve-Target','Get-UpdatePlan','Confirm-Unlocked',
    'Read-Property','Test-Container','Test-PortableBoundary','Get-IconCatalog','Resolve-IconRef','Get-TextHash',
    'Get-CardStyle','Get-DisplayLabel','Get-CaptionLayout','Get-RoleShape',
    'Get-ServicePort','Glue-End','Get-ShapeIndex','Get-NativeBounds','Get-NativeCaptionBounds','Get-NativeRouteGeometry','Test-RouteCrossing','Find-BoundaryRoute','Repair-NativeBoundaryRoutes','Get-NativeRoutingReport','Complete-NativeRouting',
    'Get-NativeShapes','Inspect-Shape','Inspect-Document','Export-Model','Get-CanonicalTargets','Get-DeletePlan','Confirm-DeleteApproval'
)
foreach ($function in $ast.FindAll({
    param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -in $pureFunctions
}, $true)) {
    . ([scriptblock]::Create($function.Extent.Text))
}
foreach ($assignment in $ast.FindAll({
    param($node)
    $node -is [Management.Automation.Language.AssignmentStatementAst] -and
        $node.Left.Extent.Text -in @('$script:icons','$script:edgeStyles')
}, $true)) { . ([scriptblock]::Create($assignment.Extent.Text)) }
$script:catalog=$null
$crudModelPath=Join-Path $OutputDirectory 'semantic-model.json'
$crudModel = @'
{
  "schemaVersion":1,"title":"Enterprise AI semantic CRUD test",
  "pages":[{
    "name":"AI experience","width":11,"height":8,"title":"Enterprise AI","subtitle":"","furniture":false,"footer":null,
    "nodes":[
      {"id":"users","kind":"card","label":"Business users","x":1.7,"y":6,"width":2,"height":1,"state":"Proposed","sourceRef":"requirements:interviews","requirementIds":["R1"],"confidence":0.6},
      {"id":"copilot","kind":"card","label":"Copilot / Teams","x":4.5,"y":6,"width":2,"height":1,"cardStyle":"detail","state":"Proposed","purpose":"User experience","sourceRef":"requirements:channels","sourceId":"ref-copilot","requirementIds":["R2"],"confidence":0.8},
      {"id":"platform","kind":"container","label":"AI platform","x":8,"y":4.5,"width":4,"height":6,"state":"Proposed"},
      {"id":"foundry","kind":"card","label":"Microsoft Foundry","x":8,"y":6,"width":2,"height":1,"parent":"platform","state":"Proposed"},
      {"id":"data","kind":"card","label":"Fabric / Microsoft 365","x":8,"y":3,"width":2.8,"height":1,"parent":"platform","state":"Observed"}
    ],
    "edges":[
      {"id":"experience","source":"users","target":"copilot","kind":"logical","label":"Experience","direction":"forward","dashed":false,"state":"Proposed","sourceRef":"requirements:R1","requirementIds":["R1"],"confidence":0.7},
      {"id":"query","source":"copilot","target":"foundry","kind":"query","label":"Query","state":"Proposed","direction":"both","dashed":false},
      {"id":"ingest","source":"data","target":"foundry","kind":"ingestion","label":"Ingestion","state":"Observed","direction":"forward","dashed":true}
    ]
  }]
}
'@ | ConvertFrom-Json
$crudModel | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $crudModelPath -Encoding UTF8
$fixtures = Join-Path $OutputDirectory 'Portability-fixtures'
[void][IO.Directory]::CreateDirectory($fixtures)
try {
    $script:visioCase='running'; $script:launchCalls=0
    function Get-RunningVisio {
        if ($script:visioCase -eq 'running') { return [pscustomobject]@{mode='existing'} }
        $code=$(if ($script:visioCase -in @('absent','not-installed')) { -2147221021 } else { -2147024891 })
        throw [Runtime.InteropServices.COMException]::new('Fixture COM failure', $code)
    }
    function Start-InstalledVisio {
        $script:launchCalls++
        if ($script:visioCase -eq 'not-installed') { throw 'Desktop Visio is not installed or registered.' }
        return [pscustomobject]@{mode='started';Visible=$true}
    }
    Assert ((Connect-Visio).mode -eq 'existing' -and $script:launchCalls -eq 0) 'Reuse running Visio without starting a second instance.'
    $script:visioCase='absent'
    Assert ((Connect-Visio).mode -eq 'started' -and $script:launchCalls -eq 1) 'Hands-free connection starts Visio when no instance is running.'
    Assert-Throws { Connect-Visio $false } '*automatic launch was explicitly disabled*'
    Assert ($script:launchCalls -eq 1) 'Explicit no-launch must not start Visio.'
    $script:visioCase='not-installed'
    Assert-Throws { Connect-Visio } '*not installed or registered*'
    Assert ($script:launchCalls -eq 2) 'Missing installation is surfaced without retry loops.'
    $script:visioCase='denied'; $denied=$false
    try { [void](Connect-Visio) }
    catch [Runtime.InteropServices.COMException] {
        if ($_.Exception.HResult -ne -2147024891) { throw }
        $denied=$true
    }
    Assert ($denied -and $script:launchCalls -eq 2) 'Unexpected COM or policy errors propagate without starting another instance.'
    $launchParameter=@($ast.ParamBlock.Parameters | Where-Object {$_.Name.VariablePath.UserPath -eq 'LaunchVisio'})[0]
    Assert ($launchParameter.DefaultValue.Extent.Text -eq '$true') 'Controller enables automatic launch by default.'
    Assert (@($ast.ParamBlock.Parameters | Where-Object {$_.Name.VariablePath.UserPath -eq 'NoLaunchVisio'}).Count -eq 1) 'Explicit no-launch opt-out remains available.'
    $checks.Add('Hands-free Visio reuse/start defaults, no-launch opt-out, missing installation, and error propagation work without changing running applications.')
    foreach ($path in @('C:', 'c:', 'C:relative.vsdx', '\relative.vsdx', 'relative.vsdx')) {
        Assert-Throws { Get-AbsolutePath $path } '*fully qualified Windows path*'
    }
    Assert ((Get-AbsolutePath 'C:\Diagrams\example.vsdx') -eq 'C:\Diagrams\example.vsdx') 'Absolute drive path.'
    Assert ((Get-AbsolutePath '\\server\share\example.vsdx') -eq '\\server\share\example.vsdx') 'UNC path.'
    $checks.Add('Drive-relative and root-relative paths are rejected; full drive and UNC paths are accepted.')

    $configPath = Join-Path $fixtures 'mapping.json'
    $config = @{syncRoots=@(
        @{local='C:\AVROOT'; cloud='https://example.invalid/Documents'},
        @{local='C:\AVROOT\Nested'; cloud='https://example.invalid/Other%20Documents'}
    )}
    $config | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $configPath -Encoding UTF8
    $script:syncRoots = @(Read-Environment $configPath)
    Assert ((Get-CloudPath 'C:\AVROOT\Nested\Drawing.vsdx') -eq 'https://example.invalid/Other Documents/Drawing.vsdx') 'Most specific sync root.'
    Assert ((Get-CloudPath 'C:\AVROOT\Drawing.vsdx') -eq 'https://example.invalid/Documents/Drawing.vsdx') 'Outer sync root.'
    Assert ((Get-CloudPath 'C:\AVROOT-other\Drawing.vsdx') -eq '') 'Sync root prefix boundary.'
    $config.syncRoots += @{local='C:\AVROOT'; cloud='https://example.invalid/Duplicate'}
    $config | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $configPath -Encoding UTF8
    Assert-Throws { Read-Environment $configPath } '*Duplicate local sync root*'
    @{syncRoots=@(@{local='C:\AVROOT';cloud='http://example.invalid/Documents'})} |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $configPath -Encoding UTF8
    Assert-Throws { Read-Environment $configPath } '*observed HTTPS*'
    Assert (@(Read-Environment (Join-Path $PSScriptRoot 'environment.example.json')).Count -eq 0) 'Empty portable mapping.'
    $checks.Add('Verified mappings use longest-prefix matching, decode URLs, and reject ambiguous or invalid configuration.')

    $documents = [pscustomobject]@{Count=1; Entries=@([pscustomobject]@{FullName='https://example.invalid/Other%20Documents/Drawing.vsdx'})}
    $documents | Add-Member -MemberType ScriptMethod -Name Item -Value { param($index) return $this.Entries[$index-1] }
    $script:app = [pscustomobject]@{Documents=$documents}
    Assert ($null -ne (Get-OpenDocument 'C:\AVROOT\Nested\Drawing.vsdx')) 'Exact mapped cloud document resolves.'
    Assert-Throws { Get-OpenDocument 'C:\Other\Drawing.vsdx' } '*cannot be matched safely*'
    $documents.Entries = @([pscustomobject]@{FullName='C:\Other\Drawing.vsdx'})
    Assert ($null -eq (Get-OpenDocument 'C:\AVROOT\Drawing.vsdx')) 'Same local filename is never enough.'
    $checks.Add('Document targeting never falls back to a basename or an unverified same-name cloud document.')

    $office = Join-Path $fixtures 'Office'
    $localized = Join-Path $office 'Visio Content\1041'
    [void][IO.Directory]::CreateDirectory($localized)
    $metric = Join-Path $localized 'AZURENETWORKING_M.VSSX'
    [void](New-Item -ItemType File -Path $metric)
    $script:app = [pscustomobject]@{Path=$office; MyShapesPath=''; StencilPaths=''; BuiltIn=(Join-Path $localized 'SDCONT_U.VSSX')}
    $script:app | Add-Member -MemberType ScriptMethod -Name GetBuiltInStencilFile -Value { param($type,$units) return $this.BuiltIn }
    $script:stencilPaths = @{}
    $StencilDirectory = ''
    Assert ((Resolve-StencilPath 'AZURENETWORKING_U.VSSX') -eq $metric) 'Localized metric stencil discovery.'
    $override = Join-Path $fixtures 'Override'
    [void][IO.Directory]::CreateDirectory($override)
    $StencilDirectory = $override
    $script:stencilPaths = @{}
    Assert-Throws { Resolve-StencilPath 'AZURENETWORKING_U.VSSX' } '*Required Azure stencil*'
    $custom = Join-Path $override 'AZURENETWORKING_U.VSSX'
    [void](New-Item -ItemType File -Path $custom)
    Assert ((Resolve-StencilPath 'AZURENETWORKING_U.VSSX') -eq $custom) 'Explicit stencil directory.'
    $checks.Add('Stencil discovery is not tied to English content and honors explicit overrides without silent fallback.')

    $valid=Run-Helper @{Action='Validate';ModelPath=$crudModelPath}
    Assert ($valid.valid -and -not $valid.comUsed) 'Validation must not acquire Visio.'
    Confirm-Model $crudModel
    $original=Get-Content -LiteralPath $crudModelPath -Raw
    foreach ($case in @(
        @{field='width';value=[double]::NaN}, @{field='height';value=[double]::PositiveInfinity},
        @{field='width';value=0}, @{field='height';value='8'}
    )) {
        $bad=$original | ConvertFrom-Json; $bad.pages[0].($case.field)=$case.value
        Assert-Throws { Confirm-Model $bad } '*finite*'
    }
    $bad=$original | ConvertFrom-Json; $bad.pages[0].edges[0].id='edge;unsafe'
    Assert-Throws { Confirm-Model $bad } '*Invalid shape identifier*'
    $bad=$original | ConvertFrom-Json; $bad.pages[0].nodes[0].confidence=1.1
    Assert-Throws { Confirm-Model $bad } '*confidence*'
    $qualitative=$original | ConvertFrom-Json
    $qualitative.pages[0].nodes[0].confidence='High'; $qualitative.pages[0].edges[0].confidence='Medium'
    Confirm-Model $qualitative
    $qualitative.pages[0].nodes[0].confidence='Definitely'
    Assert-Throws { Confirm-Model $qualitative } '*confidence*'
    $bad=$original | ConvertFrom-Json; $bad.schemaVersion=$true
    Assert-Throws { Confirm-Model $bad } '*finite JSON number*'
    $bad=$original | ConvertFrom-Json; $bad.pages[0].nodes[0].requirementIds=@(17)
    Assert-Throws { Confirm-Model $bad } '*nonempty strings*'
    $bad=$original | ConvertFrom-Json; $bad.pages[0].edges[0].source=$null
    Assert-Throws { Confirm-Model $bad } '*unattached edge*'
    $bad.pages[0].edges[0] | Add-Member -NotePropertyName beginX -NotePropertyValue 1.0
    $bad.pages[0].edges[0] | Add-Member -NotePropertyName beginY -NotePropertyValue 2.0
    Confirm-Model $bad
    $catalog=Run-Helper @{Action='Catalog'}
    Assert ($catalog.builtinIcons.Count -eq 13) 'Catalog retains all 13 built-in icon keys.'
    $checks.Add('COM-free Validate accepts provenance/logical models and explicit unglued edges, rejecting nonfinite dimensions, malformed IDs, and invalid metadata.')
    foreach ($mode in @('faithful','reference-plus-proposal')) {
        $referenceModel=$original | ConvertFrom-Json
        $referenceModel | Add-Member -NotePropertyName conversionMode -NotePropertyValue $mode
        $referenceModel | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $fixtures 'reference-model.json') -Encoding UTF8
        Assert-Throws { Run-Helper @{Action='Validate';ModelPath=(Join-Path $fixtures 'reference-model.json')} } '*persisted source contract*'
    }
    $bad=$original | ConvertFrom-Json
    $bad.pages[0].nodes[1].cardStyle='unrecognized-style'
    Assert-Throws { Confirm-Model $bad } '*cardStyle*'
    $checks.Add('Reference conversions cannot validate without a persisted source contract; unknown visual styles fail explicitly.')
    $iconModel=$original | ConvertFrom-Json
    $iconModel | Add-Member -NotePropertyName presentationProfile -NotePropertyValue 'enterprise'
    $iconModel.pages[0].nodes=@(
        [pscustomobject]@{id='glyph';kind='card';label="Storage account`nLong source semantics retained outside the caption.";displayLabel='Storage';details='Metadata-only notes';icon='storage';x=2;y=3;width=1.8;height=1.4},
        [pscustomobject]@{id='function';kind='card';cardStyle='label';label="Route selection`nOriginal functional description.";displayLabel='Select route';x=5;y=3;width=2;height=0.5}
    )
    $iconModel.pages[0].edges=@()
    Confirm-Model $iconModel
    Assert ((Get-CardStyle $iconModel.pages[0].nodes[0] 'enterprise') -eq 'icon') 'Enterprise defaults to icon rendering when a glyph is provided.'
    Assert ((Get-CardStyle $iconModel.pages[0].nodes[0]) -eq 'standard') 'Legacy defaults remain boxed standard cards.'
    Assert ((Get-DisplayLabel ([pscustomobject]@{label="First line`r`nFull detail"})) -eq 'First line') 'Default display caption is the first logical line.'
    Assert ((Get-DisplayLabel ([pscustomobject]@{displayLabel='Caption-only edit'})) -eq 'Caption-only edit') 'Caption-only changes never require a replacement canonical label.'
    $bad=$iconModel | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $bad.pages[0].nodes[0].PSObject.Properties.Remove('icon')
    $bad.pages[0].nodes[0] | Add-Member -NotePropertyName cardStyle -NotePropertyValue 'icon'
    Assert-Throws { Confirm-Model $bad } '*requires a valid icon or iconRef*'
    $bad.pages[0].nodes[0] | Add-Member -NotePropertyName icon -NotePropertyValue 'storage'
    $bad.pages[0].nodes[0].height=0.3
    Assert-Throws { Confirm-Model $bad } '*geometry is too small*'
    $bad=$iconModel | ConvertTo-Json -Depth 20 | ConvertFrom-Json; $bad.presentationProfile='automatic'
    Assert-Throws { Confirm-Model $bad } '*presentationProfile*'
    $bad=$iconModel | ConvertTo-Json -Depth 20 | ConvertFrom-Json; $bad.pages[0].nodes[0].displayLabel="One`nTwo`nThree`nFour"
    Assert-Throws { Confirm-Model $bad } '*1-3 nonempty caption lines*'
    $checks.Add('Enterprise renderer schema separates canonical and display labels, defaults safely, and rejects missing glyphs, small geometry, and invalid profiles/captions.')
    $styleHelper=Join-Path $PSScriptRoot 'Enterprise-Style.ps1'
    Assert (Test-Path -LiteralPath $styleHelper -PathType Leaf) 'Enterprise style helper is packaged.'
    $stylePath=Join-Path $fixtures 'enterprise-style.json'
    $styleModel=$iconModel | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $styleModel.pages[0].nodes=@($styleModel.pages[0].nodes[1])
    for ($i=0; $i -lt 5; $i++) {
        $styleModel.pages[0].nodes += [pscustomobject]@{
            id="storage-$i";kind='card';cardStyle='icon';icon='storage';label=("Storage $i`n"+('Canonical source detail. '*120))
            displayLabel="Storage $i";details=('Native metadata only. '*120);x=(1.5+$i*2);y=5.5;width=1.7;height=1.4
        }
    }
    $styleModel.pages[0].nodes += [pscustomobject]@{id='system';kind='container';containerStyle='boundary';boundaryType='system';label='Application system';x=5.5;y=4;width=10;height=7}
    $styleModel.pages[0].edges=@([pscustomobject]@{id='request';source='function';target='storage-0';kind='logical';label='Request'})
    $styleOriginal=$styleModel | ConvertTo-Json -Depth 30
    $styleOriginal | Set-Content -LiteralPath $stylePath -Encoding UTF8
    $styleValid=Run-Helper @{Action='Validate';ModelPath=$stylePath}
    Assert ($styleValid.valid -and -not $styleValid.comUsed -and $styleValid.styleReport.valid) 'Enterprise gate accepts concise icon-first rendering without counting hidden canonical labels/details.'
    $styleCases=@(
        @{name='boxed card ratio';change={param($m) foreach ($n in $m.pages[0].nodes) { if ($n.kind -eq 'card') { $n.cardStyle='standard'; $n.PSObject.Properties.Remove('displayLabel') } }}},
        @{name='insufficient icon fraction';change={param($m) foreach ($n in @($m.pages[0].nodes | Where-Object {$_.id -in @('storage-0','storage-1')})) { $n.cardStyle='label'; $n.PSObject.Properties.Remove('icon') }}},
        @{name='caption word limit';change={param($m) $m.pages[0].nodes[1].displayLabel='One two three four five six seven eight nine'}},
        @{name='caption line limit';change={param($m) $m.pages[0].nodes[1].displayLabel="One`nTwo`nThree`nFour"}},
        @{name='caption character limit';change={param($m) $m.pages[0].nodes[1].displayLabel=('x'*81)}},
        @{name='numbered primary caption';change={param($m) $m.pages[0].nodes[1].displayLabel="1. First`n2. Second"}},
        @{name='edge word limit';change={param($m) $m.pages[0].edges[0].label='One two three four five six seven'}},
        @{name='long primary note';change={param($m) $m.pages[0].nodes += [pscustomobject]@{id='note';kind='note';label=('word '*19).Trim();x=5;y=2;width=4;height=1}}},
        @{name='too many primary notes';change={param($m) for ($j=0;$j -lt 4;$j++) { $m.pages[0].nodes += [pscustomobject]@{id="note-$j";kind='note';label='Short note';x=5;y=2;width=4;height=1} }}},
        @{name='missing semantic boundary';change={param($m) $m.pages[0].nodes[-1].PSObject.Properties.Remove('boundaryType')}},
        @{name='invented boundary type';change={param($m) $m.pages[0].nodes[-1].boundaryType='guessed-geography'}},
        @{name='no primary diagram';change={param($m) $m.pages[0] | Add-Member -NotePropertyName role -NotePropertyValue 'notes'}},
        @{name='page visible-word budget';change={param($m) $m.pages[0].nodes=@(); for ($j=0;$j -lt 40;$j++) {
            $m.pages[0].nodes += [pscustomobject]@{id="many-$j";kind='card';cardStyle='icon';icon='storage';label='One two three four five six seven eight';displayLabel='One two three four five six seven eight';x=5;y=4;width=4;height=2}
        }; $m.pages[0].edges=@()}}
    )
    foreach ($case in $styleCases) {
        $bad=$styleOriginal | ConvertFrom-Json; & $case.change $bad
        $bad | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $stylePath -Encoding UTF8
        $report=& $styleHelper -ModelPath $stylePath -ReportOnly | ConvertFrom-Json
        Assert (-not $report.valid -and @($report.errors).Count -gt 0) "Enterprise style rejects $($case.name)."
        Assert-Throws { & $styleHelper -ModelPath $stylePath } '*'
    }
    $bad=$styleOriginal | ConvertFrom-Json; $bad.presentationProfile='reference'
    $bad | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $stylePath -Encoding UTF8
    Assert-Throws { & $styleHelper -ModelPath $stylePath } '*presentationProfile*'
    $notesExempt=$styleOriginal | ConvertFrom-Json
    $notesExempt.pages += [pscustomobject]@{name='Source notes';role='notes';width=11;height=8;furniture=$false;nodes=@(
        [pscustomobject]@{id='source-notes';kind='note';label=('Detailed supporting source notes. '*100);x=5.5;y=4;width=10;height=7}
    );edges=@()}
    $notesExempt | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $stylePath -Encoding UTF8
    Assert-Throws { Run-Helper @{Action='Validate';ModelPath=$stylePath} } '*NotesContract*'
    $inferred=$styleOriginal | ConvertFrom-Json
    foreach ($node in @($inferred.pages[0].nodes | Where-Object {$_.kind -eq 'card' -and (Get-Value $_ 'icon' '')})) {
        $node.PSObject.Properties.Remove('cardStyle'); $node.PSObject.Properties.Remove('displayLabel')
    }
    $inferred | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $stylePath -Encoding UTF8
    $inferredReport=& $styleHelper -ModelPath $stylePath -ReportOnly | ConvertFrom-Json
    Assert $inferredReport.valid 'Enterprise gate must infer icon style from icon/iconRef under enterprise profile, matching the renderer, and use the first canonical line as its default caption.'
    Assert (Run-Helper @{Action='Validate';ModelPath=$stylePath}).valid 'Controller and style helper agree on inferred icon styling.'
    $checks.Add('Enterprise style: inferred/default captions, hidden semantics, contract-only notes exemption, and 14 negative ratio/text/boundary/profile/word-budget cases pass without COM.')

    $pack=$styleOriginal | ConvertFrom-Json
    $pack | Add-Member -NotePropertyName outputContract -NotePropertyValue 'architecture-pack-v1.6'
    $pack | Add-Member -NotePropertyName hardening -NotePropertyValue ([pscustomobject]@{
        status='already-enterprise';reason='The main design already includes the controls required for this approved architecture.'
    })
    $pack.pages[0].name='Main architecture'
    $pack.pages[0] | Add-Member -NotePropertyName view -NotePropertyValue 'main'
    $pack.pages[0] | Add-Member -NotePropertyName role -NotePropertyValue 'diagram'
    $reviewText='The main architecture already includes the security and operational controls required by its approved scope. This review therefore does not propose another architecture or duplicate existing components. Ownership, access boundaries, monitoring responsibilities, and data handling remain those specified on the main page. Additional enterprise controls require an identified gap and separate approval rather than an automatic redesign.'
    $narrative='The main architecture begins with route selection and passes the request to the storage service through the request relationship. The storage services retain the information required by their respective workloads. This flowchart describes those main components and relationships only. It does not introduce proposed hardening services. Operators use the canonical component descriptions and source requirements to understand ownership, expected behavior, and the responsibilities associated with every storage component.'
    $pack.pages+=@(
        [pscustomobject]@{name='Hardening applicability';title='Hardening applicability review';view='hardening';role='notes';width=11;height=8;furniture=$false;nodes=@(
            [pscustomobject]@{id='review';kind='note';label=$reviewText;x=5.5;y=4;width=10;height=6}
        );edges=@()},
        [pscustomobject]@{name='Main flowchart';title='Main architecture flowchart and writeup';view='flowchart';role='notes';width=11;height=8;furniture=$false;nodes=@(
            [pscustomobject]@{id='select-step';kind='card';cardStyle='detail';label='Select the correct storage route for the incoming request using the main architecture routing function.';mainNodeIds=@('function');mainEdgeIds=@('request');x=2.8;y=5.4;width=4.5;height=2.2},
            [pscustomobject]@{id='store-step';kind='card';cardStyle='standard';label='Store workload information in the five designated storage services while preserving their documented ownership and responsibilities.';mainNodeIds=@('storage-0','storage-1','storage-2','storage-3','storage-4');mainEdgeIds=@();x=8.2;y=5.4;width=4.5;height=2.2},
            [pscustomobject]@{id='writeup';kind='note';label=$narrative;x=5.5;y=2;width=10;height=2.4}
        );edges=@([pscustomobject]@{id='flow-request';source='select-step';target='store-step';kind='logical';label='Route then store';direction='forward'})}
    )
    $packOriginal=ConvertTo-Json -InputObject $pack -Depth 40
    $packPath=Join-Path $fixtures 'architecture-pack.json'
    $packOriginal | Set-Content -LiteralPath $packPath -Encoding UTF8
    Confirm-Model $pack
    Confirm-ArchitecturePack $pack -Required
    Assert (Run-Helper @{Action='Validate';ModelPath=$packPath}).styleReport.valid 'Valid three-page enterprise pack passes style and contract without COM.'
    $referencePack=$packOriginal | ConvertFrom-Json; $referencePack.presentationProfile='reference'
    Confirm-Model $referencePack
    Assert ((& $styleHelper -InputModel $referencePack | ConvertFrom-Json).valid) 'Reference pack main/proposed pages obey the same icon-first presentation guardrails.'
    $referencePack.pages[0].nodes[1].PSObject.Properties.Remove('cardStyle')
    Confirm-Model $referencePack
    Assert ((& $styleHelper -InputModel $referencePack | ConvertFrom-Json).valid) 'Reference pack inferred icon styling matches enterprise rendering without changing legacy reference defaults.'
    $notApplicable=$packOriginal | ConvertFrom-Json; $notApplicable.hardening.status='not-applicable'
    Confirm-Model $notApplicable
    $proposed=$packOriginal | ConvertFrom-Json
    $proposed.hardening.status='proposed'; $proposed.hardening.reason='Storage requires private access controls that are absent from the approved main design.'
    $proposed.pages[1]=$proposed.pages[0] | ConvertTo-Json -Depth 40 | ConvertFrom-Json
    $proposed.pages[1].name='Proposed enterprise hardening'; $proposed.pages[1].view='hardening'
    $proposed.pages[1].nodes[1].details='Use private access controls to close the identified public access gap.'
    Confirm-Model $proposed
    Assert ((& $styleHelper -InputModel $proposed | ConvertFrom-Json).valid) 'A reasoned semantic change on a proposed architecture passes.'
    $proposedOriginal=ConvertTo-Json -InputObject $proposed -Depth 40
    $containment=$proposedOriginal | ConvertFrom-Json
    $containment.pages[1].nodes=$containment.pages[0].nodes | ConvertTo-Json -Depth 40 | ConvertFrom-Json
    $containment.pages[1].nodes[0] | Add-Member -NotePropertyName parent -NotePropertyValue 'system'
    Confirm-Model $containment
    $packCases=@(
        @{name='unknown output contract';change={param($m) $m.outputContract='architecture-pack-v9'}},
        @{name='null output contract';change={param($m) $m.outputContract=$null}},
        @{name='array output contract';change={param($m) $m.outputContract=@('architecture-pack-v1.6')}},
        @{name='legacy presentation';change={param($m) $m.presentationProfile='legacy'}},
        @{name='missing hardening';change={param($m) $m.PSObject.Properties.Remove('hardening')}},
        @{name='unknown status';change={param($m) $m.hardening.status='skip'}},
        @{name='empty reason';change={param($m) $m.hardening.reason=' '}},
        @{name='nonstring reason';change={param($m) $m.hardening.reason=17}},
        @{name='extra page';change={param($m) $m.pages+=($m.pages[2] | ConvertTo-Json -Depth 40 | ConvertFrom-Json);$m.pages[3].name='Fourth page'}},
        @{name='missing page';change={param($m) $m.pages=@($m.pages[0],$m.pages[2])}},
        @{name='wrong page order';change={param($m) $m.pages=@($m.pages[1],$m.pages[0],$m.pages[2])}},
        @{name='missing view';change={param($m) $m.pages[2].PSObject.Properties.Remove('view')}},
        @{name='array view';change={param($m) $m.pages[2].view=@('flowchart')}},
        @{name='main notes bypass';change={param($m) $m.pages[0].role='notes'}},
        @{name='hardening diagram bypass';change={param($m) $m.pages[1].role='diagram'}},
        @{name='flowchart diagram role';change={param($m) $m.pages[2].role='diagram'}},
        @{name='redundant non-proposed diagram';change={param($m) $m.pages[1].nodes=$m.pages[0].nodes}},
        @{name='short applicability review';change={param($m) $m.pages[1].nodes[0].label='Already enterprise.'}},
        @{name='short writeup';change={param($m) $m.pages[2].nodes[2].label='Main flow.'}},
        @{name='short flowchart label';change={param($m) $m.pages[2].nodes[0].label='Select route'}},
        @{name='caption flowchart card';change={param($m) $m.pages[2].nodes[0].cardStyle='label'}},
        @{name='missing reference array';change={param($m) $m.pages[2].nodes[0].PSObject.Properties.Remove('mainEdgeIds')}},
        @{name='string reference array';change={param($m) $m.pages[2].nodes[0].mainNodeIds='function'}},
        @{name='nonstrings in reference array';change={param($m) $m.pages[2].nodes[0].mainNodeIds=@(1)}},
        @{name='unknown node reference';change={param($m) $m.pages[2].nodes[0].mainNodeIds=@('hardening-only')}},
        @{name='wrong-case node reference';change={param($m) $m.pages[2].nodes[0].mainNodeIds=@('FUNCTION')}},
        @{name='unknown edge reference';change={param($m) $m.pages[2].nodes[0].mainEdgeIds=@('flow-request')}},
        @{name='missing card coverage';change={param($m) $m.pages[2].nodes[1].mainNodeIds=@('storage-0')}},
        @{name='missing edge coverage';change={param($m) $m.pages[2].nodes[0].mainEdgeIds=@()}},
        @{name='empty card mapping';change={param($m) $m.pages[2].nodes[0].mainNodeIds=@();$m.pages[2].nodes[0].mainEdgeIds=@()}},
        @{name='mapping on wrong page';change={param($m) $m.pages[0].nodes[0] | Add-Member -NotePropertyName mainNodeIds -NotePropertyValue @('function')}},
        @{name='no connected edges';change={param($m) $m.pages[2].edges=@()}},
        @{name='self edge';change={param($m) $m.pages[2].edges[0].target='select-step'}},
        @{name='undirected edge';change={param($m) $m.pages[2].edges[0].direction='none'}},
        @{name='edge to narrative';change={param($m) $m.pages[2].edges[0].target='writeup'}},
        @{name='disconnected flow card';change={param($m) $extra=$m.pages[2].nodes[0] | ConvertTo-Json -Depth 40 | ConvertFrom-Json;$extra.id='isolated';$m.pages[2].nodes+=$extra}}
    )
    foreach ($case in $packCases) {
        $bad=$packOriginal | ConvertFrom-Json; & $case.change $bad
        Assert-Throws { Confirm-Model $bad } '*'
    }
    $bad=$proposedOriginal | ConvertFrom-Json
    $bad.pages[1].nodes=$bad.pages[0].nodes; $bad.pages[1].edges=$bad.pages[0].edges
    Assert-Throws { Confirm-Model $bad } '*semantic architecture change*'
    $bad=$proposedOriginal | ConvertFrom-Json; $bad.pages[1].edges=@()
    Assert-Throws { Confirm-Model $bad } '*real card entities*'
    $bad=$proposedOriginal | ConvertFrom-Json
    $bad.pages[1].nodes=$bad.pages[0].nodes | ConvertTo-Json -Depth 40 | ConvertFrom-Json
    foreach ($node in $bad.pages[1].nodes) { $node.x+=0.01 }
    Assert-Throws { Confirm-Model $bad } '*semantic architecture change*'
    $bad=$packOriginal | ConvertFrom-Json; $bad.pages[0].nodes[1].displayLabel='One two three four five six seven eight nine'
    Assert-Throws { & $styleHelper -InputModel $bad } '*EntityCaption*'
    Assert-Throws { Confirm-ArchitecturePack $crudModel -Required } '*-LegacyModel*'
    Confirm-ArchitecturePack $crudModel
    Assert-Throws { Run-Helper @{Action='New';ModelPath=$crudModelPath;DocumentPath=$drawing;NoLaunchVisio=$true} } '*-LegacyModel*'
    $bad=$packOriginal | ConvertFrom-Json; $bad.outputContract='other'
    $bad | ConvertTo-Json -Depth 40 | Set-Content -LiteralPath $packPath -Encoding UTF8
    Assert-Throws { Run-Helper @{Action='Validate';ModelPath=$packPath} } '*Unknown outputContract*'
    Assert-Throws { Run-Helper @{Action='New';LegacyModel=$true;ModelPath=$packPath;DocumentPath=$drawing;NoLaunchVisio=$true} } '*Unknown outputContract*'
    $checks.Add("Architecture pack: three statuses, reference/enterprise style, multi-ID coverage, explicit legacy opt-in and $($packCases.Count+5) malformed/downgrade/duplicate cases pass without COM.")

    $IconDirectory=Join-Path $fixtures 'icons'
    [void][IO.Directory]::CreateDirectory($IconDirectory)
    $svg=Join-Path $IconDirectory 'fixture.svg'
    '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="12"><rect width="24" height="12"/></svg>' |
        Set-Content -LiteralPath $svg -Encoding UTF8
    $entry=[pscustomobject]@{id='fixture';path='fixture.svg';sha256=(Get-FileHash -LiteralPath $svg -Algorithm SHA256).Hash;usable=$true;issue=''}
    ConvertTo-Json -InputObject @($entry) | Set-Content -LiteralPath (Join-Path $IconDirectory 'catalog.json') -Encoding UTF8
    $script:catalog=$null
    Assert ((Resolve-IconRef 'fixture') -eq $svg) 'Contained SVG hash resolves.'
    Assert-Throws { Resolve-IconRef 'missing' } '*Missing iconRef*'
    $script:catalog['fixture'].usable=$false; $script:catalog['fixture'].issue='External resource rejected by SVG audit.'
    Assert-Throws { Resolve-IconRef 'fixture' } '*not usable*External resource*'
    $script:catalog['fixture'].usable='false'
    Assert-Throws { Resolve-IconRef 'fixture' } '*usable must be a boolean*'
    $script:catalog['fixture'].usable=$true
    $script:catalog['fixture'].path='..\fixture.svg'
    Assert-Throws { Resolve-IconRef 'fixture' } '*unsafe path segments*'
    $script:catalog['fixture'].path='https://example.invalid/icon.svg'
    Assert-Throws { Resolve-IconRef 'fixture' } '*contained relative SVG path*'
    $script:catalog['fixture'].path='fixture.png'
    Assert-Throws { Resolve-IconRef 'fixture' } '*only .svg*'
    $script:catalog['fixture'].path='fixture.svg'; $script:catalog['fixture'].sha256=('0'*64)
    Assert-Throws { Resolve-IconRef 'fixture' } '*SHA256 mismatch*'
    ConvertTo-Json -InputObject @($entry,$entry) | Set-Content -LiteralPath (Join-Path $IconDirectory 'catalog.json') -Encoding UTF8
    $script:catalog=$null
    Assert-Throws { Get-IconCatalog } '*Ambiguous catalog id*'
    $checks.Add('SVG catalog rejects unusable audit results, traversal, URLs, unsupported extensions, unknown/ambiguous IDs, and mismatched hashes.')

    $update='{"schemaVersion":1,"updates":[{"pageId":0,"visioId":7,"set":{"label":"Approved","state":"Confirmed","requirementIds":["R1"],"confidence":1}}]}' | ConvertFrom-Json
    Confirm-Changes $update 'Update'
    foreach ($invalid in @(
        '{"schemaVersion":1,"updates":[{"visioId":7,"set":{"label":"X"}}]}',
        '{"schemaVersion":1,"updates":[{"pageId":0,"visioId":7,"shapeId":"node","set":{"label":"X"}}]}',
        '{"schemaVersion":1,"updates":[{"pageId":0,"visioId":7,"set":{"delete":true}}]}',
        '{"schemaVersion":1,"updates":[{"pageId":0,"visioId":7,"set":{"width":-1}}]}',
        '{"schemaVersion":1,"updates":[{"pageId":0,"visioId":7,"set":{"source":{"pageId":2,"visioId":8}}}]}',
        '{"schemaVersion":1,"targets":[{"pageId":0,"visioId":0}]}'
    )) {
        $request=$invalid | ConvertFrom-Json
        Assert-Throws { Confirm-Changes $request $(if (Test-Field $request 'updates') { 'Update' } else { 'Delete' }) } '*'
    }
    $shape=[pscustomobject]@{ID=7;OneD=0;Text='Original';CellMap=@{}}
    foreach ($cell in @{PinX=2.0;PinY=2.0;Width=2.0;Height=1.0;LocPinX=1.0;LocPinY=0.5;Angle=0.0}.GetEnumerator()) {
        $shape.CellMap[$cell.Key]=[pscustomobject]@{ResultIU=$cell.Value}
    }
    $shape | Add-Member ScriptMethod CellsU {param($name) return $this.CellMap[$name]}
    $shape | Add-Member ScriptMethod CellExistsU {param($name,$exists) return $this.CellMap.ContainsKey($name)}
    $shapes=[pscustomobject]@{Count=1;Entries=@($shape)}
    $shapes | Add-Member ScriptMethod Item {param($index) return $this.Entries[$index-1]}
    $sheet=[pscustomobject]@{}
    $sheet | Add-Member ScriptMethod CellsU {param($name) return [pscustomobject]@{ResultIU=10.0}}
    $pageMock=[pscustomobject]@{ID=0;Shapes=$shapes;PageSheet=$sheet}
    $pagesMock=[pscustomobject]@{Count=1;Entries=@($pageMock)}
    $pagesMock | Add-Member ScriptMethod Item {param($index) return $this.Entries[$index-1]}
    $documentMock=[pscustomobject]@{Pages=$pagesMock}
    Assert (@(Get-UpdatePlan $documentMock $update).Count -eq 1) 'Native Visio-ID target resolves without managed metadata.'
    $update.updates += ('{"pageId":0,"visioId":99,"set":{"label":"Never"}}' | ConvertFrom-Json)
    Assert-Throws { Get-UpdatePlan $documentMock $update } '*resolved to 0 full shapes*'
    Assert ($shape.Text -eq 'Original') 'Bad batch preflight never edits any target.'
    $update='{"schemaVersion":1,"updates":[{"pageId":0,"visioId":7,"set":{"x":11}}]}' | ConvertFrom-Json
    Assert-Throws { Get-UpdatePlan $documentMock $update } '*outside its page*'
    $checks.Add('CRUD preflight requires page-qualified full-shape targets, rejects unsafe properties and cross-page endpoints, and makes no partial edits for bad batches.')

    function New-MockCollection($Entries) {
        $collection=[pscustomobject]@{Count=@($Entries).Count;Entries=@($Entries)}
        $collection | Add-Member ScriptMethod Item {param($index) return $this.Entries[$index-1]}
        return $collection
    }
    function New-MockShape([int]$Id, [bool]$OneD=$false) {
        $mock=[pscustomobject]@{
            ID=$Id;OneD=$OneD;Text="Shape $Id";NameU="Sheet.$Id";CellMap=@{}
            LayerCount=0;Shapes=(New-MockCollection @());Connects=(New-MockCollection @())
        }
        foreach ($pair in @{PinX=2.0;PinY=2.0;Width=2.0;Height=1.0;Angle=0.0;BeginX=1.0;BeginY=2.0;EndX=3.0;EndY=2.0;BeginArrow=0;EndArrow=13;LinePattern=2}.GetEnumerator()) {
            $mock.CellMap[$pair.Key]=[pscustomobject]@{ResultIU=$pair.Value}
        }
        $mock.CellMap['LineColor']=[pscustomobject]@{FormulaU='RGB(100,116,139)'}
        $mock | Add-Member ScriptMethod CellsU {param($name) return $this.CellMap[$name]}
        $mock | Add-Member ScriptMethod CellExistsU {param($name,$exists) return $this.CellMap.ContainsKey($name)}
        return $mock
    }
    $first=New-MockShape 7; $last=New-MockShape 9; $edge=New-MockShape 8 $true; $container=New-MockShape 10
    $containerCell=[pscustomobject]@{}
    $containerCell | Add-Member ScriptMethod ResultStr {param($units) return 'Container'}
    $container.CellMap['User.msvStructureType']=$containerCell
    $membership=[pscustomobject]@{}
    $membership | Add-Member ScriptMethod GetMemberShapes {param($flags) return @(7,9)}
    $container | Add-Member -NotePropertyName ContainerProperties -NotePropertyValue $membership
    $glue=@(
        [pscustomobject]@{FromSheet=$edge;FromCell=[pscustomobject]@{Name='BeginX'};ToSheet=$first;ToCell=[pscustomobject]@{Name='PinX'}},
        [pscustomobject]@{FromSheet=$edge;FromCell=[pscustomobject]@{Name='EndX'};ToSheet=$last;ToCell=[pscustomobject]@{Name='PinX'}}
    )
    $edge.Connects=New-MockCollection $glue
    $sheet | Add-Member ScriptMethod CellExistsU {param($name,$exists) return $false}
    $pageMock.Shapes=New-MockCollection @($first,$edge,$last,$container)
    $pageMock | Add-Member -NotePropertyName Name -NotePropertyValue 'Mock'
    $pageMock | Add-Member -NotePropertyName Layers -NotePropertyValue (New-MockCollection @())
    $pageMock | Add-Member -NotePropertyName Connects -NotePropertyValue (New-MockCollection $glue)
    $documentMock | Add-Member -NotePropertyName FullName -NotePropertyValue 'C:\Diagrams\mock.vsdx'
    $documentMock | Add-Member -NotePropertyName Saved -NotePropertyValue $true
    $documentMock | Add-Member -NotePropertyName DocumentSheet -NotePropertyValue $sheet
    $request='{"schemaVersion":1,"targets":[{"pageId":0,"visioId":7}]}' | ConvertFrom-Json
    $plan=Get-DeletePlan $documentMock $request
    Assert ($plan.report.requiredTargets.Count -eq 2 -and $plan.report.missingTargets[0].visioId -eq 8) 'Native endpoint preview includes its incoming connector but not the other endpoint.'
    Assert-Throws { Confirm-DeleteApproval $documentMock $request $plan } '*undeclared*'
    $approved=[pscustomobject]@{schemaVersion=1;targets=$plan.report.requiredTargets;approvedTargets=$plan.report.requiredTargets;previewToken=$plan.token}
    $complete=Get-DeletePlan $documentMock $approved
    Assert ($complete.token -eq $plan.token) 'Preview token binds the complete list regardless of which target exposed a dependency.'
    Confirm-DeleteApproval $documentMock $approved $complete
    $first.Text='Changed after preview'
    $stale=Get-DeletePlan $documentMock $approved
    Assert-Throws { Confirm-DeleteApproval $documentMock $approved $stale } '*stale*'
    $request='{"schemaVersion":1,"targets":[{"pageId":0,"visioId":10}]}' | ConvertFrom-Json
    $plan=Get-DeletePlan $documentMock $request
    Assert ($plan.report.requiredTargets.Count -eq 4 -and $plan.report.missingTargets.Count -eq 3) 'Container closure includes both native descendants and their connector.'
    $checks.Add('Native delete preflight computes complete dependency closure, requires exact approval, and invalidates stale previews without COM.')
    $documentMock | Add-Member -NotePropertyName Title -NotePropertyValue 'Native semantic fixture'
    $exportModel=Export-Model $documentMock
    Confirm-Model $exportModel
    Assert ($exportModel.pages[0].nodes.Count -eq 3 -and $exportModel.pages[0].edges.Count -eq 1) 'Native semantic export is reusable schemaVersion 1.'
    Assert ($exportModel.pages[0].edges[0].source -eq 'native-p0-s7' -and $exportModel.pages[0].edges[0].target -eq 'native-p0-s9') 'Export uses observed native glue, not invented target names.'
    Assert ($exportModel.pages[0].edges[0].kind -eq 'association' -and $exportModel.pages[0].edges[0].dashed -and $exportModel.pages[0].edges[0].direction -eq 'forward') 'Unknown native relationship remains association with real arrow/dash semantics.'
    $edge.Connects=New-MockCollection @()
    $exportModel=Export-Model $documentMock
    Confirm-Model $exportModel
    Assert ($null -eq $exportModel.pages[0].edges[0].source -and $null -eq $exportModel.pages[0].edges[0].target) 'Missing glue remains explicitly null with saved endpoint coordinates.'
    $checks.Add('COM-free native semantic roundtrip preserves glue, container membership, direction/dashes, and explicit unglued endpoint uncertainty.')
    function Set-MockProperty($Shape,[string]$Name,[string]$Value) {
        $cell=[pscustomobject]@{Value=$Value}
        $cell | Add-Member ScriptMethod ResultStr {param($units) return $this.Value}
        $Shape.CellMap['Prop.'+$Name]=$cell
    }
    $root=New-MockShape 20; $caption=New-MockShape 21; $glyph=New-MockShape 22
    foreach ($pair in @{AvId='icon-node';Kind='card';CardStyle='icon';FullLabel="Canonical service`nOriginal full description";Icon='storage';IconSize='0.72';IconAspect='1';Details='Native notes';SourceIdRef='source-icon';BaseFontSize='10'}.GetEnumerator()) { Set-MockProperty $root $pair.Key $pair.Value }
    $root | Add-Member -NotePropertyName Characters -NotePropertyValue ([pscustomobject]@{Text="Canonical service`nOriginal full description"})
    $root.Text='Visible service'
    $root.CellMap.Height.ResultIU=1.4
    $root.CellMap['HideText']=[pscustomobject]@{ResultIU=1}
    $caption.Text='Visible service'; $caption.NameU='Arbitrary-manual-name'
    Set-MockProperty $caption 'AvRole' 'caption'; Set-MockProperty $glyph 'AvRole' 'icon'
    $glyph.CellMap.Width.ResultIU=0.72; $glyph.CellMap.Height.ResultIU=0.72
    $root.Shapes=New-MockCollection @($caption,$glyph)
    $pageMock.Shapes=New-MockCollection @($root)
    $record=Inspect-Shape $root
    Assert ($record.label -ceq $root.Characters.Text -and $record.displayLabel -eq $caption.Text -and $record.sourceId -eq 'source-icon') 'Inspect reads the own-group canonical text even when Shape.Text aliases the caption.'
    $roundtrip=Export-Model $documentMock
    Assert ($roundtrip.pages[0].nodes[0].label -ceq $root.Characters.Text -and $record.canonicalLabelSource -ceq 'root.Characters.Text') 'Export preserves full canonical text rather than the caption proxy or stale metadata.'
    Assert ($roundtrip.pages[0].nodes[0].details -eq 'Native notes' -and $roundtrip.pages[0].nodes[0].iconSize -eq 0.72 -and $roundtrip.pages[0].nodes[0].displayLabel -eq 'Visible service') 'Semantic export retains caption, notes, and nominal icon size.'
    Confirm-Model $roundtrip
    $request='{"schemaVersion":1,"updates":[{"pageId":0,"visioId":20,"set":{"displayLabel":"Short caption","details":"New notes"}}]}' | ConvertFrom-Json
    Confirm-Changes $request 'Update'
    Assert (@(Get-UpdatePlan $documentMock $request).Count -eq 1) 'Caption/detail update preflights without changing canonical text.'
    $request.updates[0].set | Add-Member -NotePropertyName width -NotePropertyValue 2.1
    Assert-Throws { Get-UpdatePlan $documentMock $request } '*Width/height updates on icon groups are blocked*'
    $root.Shapes=New-MockCollection @($glyph)
    Assert ($null -eq (Inspect-Shape $root).displayLabel) 'Missing tagged caption is reported, never guessed from icon text.'
    Assert-Throws { Export-Model $documentMock } '*native caption child missing or ambiguous*'
    $checks.Add('Tagged native caption inspection/export and safe update preflight preserve canonical content; missing captions and brand-distorting resizes fail explicitly.')
    function New-MockPackDocument($Model) {
        $nativePages=@(); $nextId=100
        foreach ($spec in $Model.pages) {
            $pageSheet=New-MockShape 0
            $pageSheet.CellMap['PageWidth']=[pscustomobject]@{ResultIU=$spec.width}
            $pageSheet.CellMap['PageHeight']=[pscustomobject]@{ResultIU=$spec.height}
            Set-MockProperty $pageSheet 'AvFurniture' 'False'
            Set-MockProperty $pageSheet 'AvPageRole' $spec.role
            Set-MockProperty $pageSheet 'AvPageView' $spec.view
            $index=@{}; $nativeShapes=@(); $connections=@()
            foreach ($node in $spec.nodes) {
                $nextId++; $native=New-MockShape $nextId
                $native.Text=$node.label
                foreach ($pair in @(@('x','PinX'),@('y','PinY'),@('width','Width'),@('height','Height'))) { $native.CellMap[$pair[1]].ResultIU=$node.($pair[0]) }
                Set-MockProperty $native 'AvId' $node.id; Set-MockProperty $native 'Kind' $node.kind
                foreach ($pair in @(@('details','Details'),@('purpose','Purpose'),@('icon','Icon'),@('boundaryType','BoundaryType'))) {
                    if (Test-Field $node $pair[0]) { Set-MockProperty $native $pair[1] $node.($pair[0]) }
                }
                foreach ($pair in @(@('mainNodeIds','MainNodeIds'),@('mainEdgeIds','MainEdgeIds'))) {
                    if (Test-Field $node $pair[0]) { Set-MockProperty $native $pair[1] (ConvertTo-Json -InputObject $node.($pair[0]) -Compress) }
                }
                if ($node.kind -eq 'container') {
                    $containerCell=[pscustomobject]@{}
                    $containerCell | Add-Member ScriptMethod ResultStr {param($units) return 'Container'}
                    $native.CellMap['User.msvStructureType']=$containerCell
                    $membership=[pscustomobject]@{}
                    $membership | Add-Member ScriptMethod GetMemberShapes {param($flags) return @()}
                    $native | Add-Member -NotePropertyName ContainerProperties -NotePropertyValue $membership
                } elseif ($node.kind -eq 'card') {
                    $style=Get-CardStyle $node $Model.presentationProfile
                    Set-MockProperty $native 'CardStyle' $style; Set-MockProperty $native 'BaseFontSize' '10'
                    if ($style -in @('icon','label')) {
                        $native | Add-Member -NotePropertyName Characters -NotePropertyValue ([pscustomobject]@{Text=$node.label})
                        $nextId++; $child=New-MockShape $nextId; $child.Text=Get-DisplayLabel $node
                        Set-MockProperty $child 'AvRole' 'caption'
                        $children=@($child)
                        Set-MockProperty $native 'FullLabel' $node.label
                        $native.CellMap['HideText']=[pscustomobject]@{ResultIU=1}
                        if ($style -eq 'icon') {
                            $nextId++; $glyph=New-MockShape $nextId
                            $glyph.CellMap.Width.ResultIU=0.72; $glyph.CellMap.Height.ResultIU=0.72
                            Set-MockProperty $glyph 'AvRole' 'icon'; $children+=$glyph
                            Set-MockProperty $native 'IconSize' '0.72'; Set-MockProperty $native 'IconAspect' '1'
                        }
                        $native.Shapes=New-MockCollection $children
                    }
                }
                $index[$node.id]=$native; $nativeShapes+=$native
            }
            foreach ($edge in $spec.edges) {
                $nextId++; $native=New-MockShape $nextId $true; $native.Text=$edge.label
                foreach ($pair in @(@('id','AvId'),@('kind','Relationship'),@('source','SourceId'),@('target','TargetId'))) { Set-MockProperty $native $pair[1] $edge.($pair[0]) }
                $glue=@(
                    [pscustomobject]@{FromSheet=$native;FromCell=[pscustomobject]@{Name='BeginX'};ToSheet=$index[$edge.source];ToCell=[pscustomobject]@{Name='PinX'}},
                    [pscustomobject]@{FromSheet=$native;FromCell=[pscustomobject]@{Name='EndX'};ToSheet=$index[$edge.target];ToCell=[pscustomobject]@{Name='PinX'}}
                )
                $native.Connects=New-MockCollection $glue; $connections+=$glue; $nativeShapes+=$native
            }
            $nativePages+=[pscustomobject]@{ID=$nativePages.Count;Name=$spec.name;Shapes=(New-MockCollection $nativeShapes);PageSheet=$pageSheet;Layers=(New-MockCollection @());Connects=(New-MockCollection $connections)}
        }
        $documentSheet=New-MockShape 0
        Set-MockProperty $documentSheet 'OutputContract' $Model.outputContract
        Set-MockProperty $documentSheet 'Hardening' (ConvertTo-Json -InputObject $Model.hardening -Compress)
        Set-MockProperty $documentSheet 'PresentationProfile' $Model.presentationProfile
        return [pscustomobject]@{FullName='C:\Diagrams\architecture-pack.vsdx';Title=$Model.title;Saved=$true;DocumentSheet=$documentSheet;Pages=(New-MockCollection $nativePages)}
    }
    $pack=$packOriginal | ConvertFrom-Json
    $nativePack=New-MockPackDocument $pack
    $snapshot=Inspect-Document $nativePack
    Assert ($snapshot.outputContract -ceq $pack.outputContract -and $snapshot.hardening.status -ceq $pack.hardening.status -and
        ($snapshot.pages.view -join ',') -ceq 'main,hardening,flowchart') 'Native inspection preserves document contract, parsed hardening JSON and ordered AvPageView metadata.'
    $roundtrip=Export-Model $nativePack
    Assert ($roundtrip.outputContract -ceq $pack.outputContract -and $roundtrip.hardening.reason -ceq $pack.hardening.reason -and
        $roundtrip.pages[2].nodes[1].mainNodeIds.Count -eq 5 -and $roundtrip.pages[2].nodes[1].mainEdgeIds -is [array] -and
        $roundtrip.pages[2].nodes[1].mainEdgeIds.Count -eq 0 -and $roundtrip.pages[2].nodes[0].mainEdgeIds[0] -eq 'request') 'Export preserves multi-ID and empty-array MAIN coverage exactly, never guessing references.'
    Confirm-NativeArchitecturePack $nativePack
    Confirm-PackMerge $nativePack $pack
    Assert-Throws { Confirm-PackMerge $nativePack $crudModel } '*cannot remove or change*'
    $bad=$packOriginal | ConvertFrom-Json; $bad.pages[0].name='Renamed main'
    Assert-Throws { Confirm-PackMerge $nativePack $bad } '*cannot add, rename, reorder*'
    $bad=$packOriginal | ConvertFrom-Json; $bad.pages[0].view='hardening'
    Assert-Throws { Confirm-PackMerge $nativePack $bad } '*in that order*'
    $bad=$packOriginal | ConvertFrom-Json; $bad.hardening.status='not-applicable'
    Assert-Throws { Confirm-PackMerge $nativePack $bad } '*cannot change hardening*'
    $flow=$nativePack.Pages.Item(3)
    $flow.Shapes.Entries[0].CellMap.Remove('Prop.MainEdgeIds')
    Assert-Throws { Confirm-NativeArchitecturePack $nativePack } '*mainEdgeIds string arrays*'
    Set-MockProperty $flow.Shapes.Entries[0] 'MainEdgeIds' '["request"]'
    Set-MockProperty $flow.Shapes.Entries[0] 'MainNodeIds' '["hardening-only"]'
    Assert-Throws { Export-Model $nativePack } '*Unknown MAIN reference*'
    Set-MockProperty $flow.Shapes.Entries[0] 'MainNodeIds' '["function"]'
    $flow.Shapes.Entries[0].Text='Too short'
    Assert-Throws { Confirm-NativeArchitecturePack $nativePack } '*meaningful labels*'
    $flow.Shapes.Entries[0].Text=$pack.pages[2].nodes[0].label
    $flow.Shapes.Entries[2].CellMap['HideText']=[pscustomobject]@{ResultIU=1}
    Assert-Throws { Export-Model $nativePack } '*not a visible explanation*'
    $flow.Shapes.Entries[2].CellMap.Remove('HideText')
    $nativePack.Pages.Entries=@($nativePack.Pages.Entries[2],$nativePack.Pages.Entries[1],$nativePack.Pages.Entries[0])
    Assert-Throws { Export-Model $nativePack } '*in that order*'
    $nativePack=New-MockPackDocument $pack
    Set-MockProperty $nativePack.DocumentSheet 'OutputContract' 'unknown'
    Assert-Throws { Export-Model $nativePack } '*Unknown outputContract*'
    $nativePack=New-MockPackDocument $pack
    Set-MockProperty $nativePack.Pages.Item(3).Shapes.Entries[0] 'MainNodeIds' '"function"'
    Assert-Throws { Inspect-Document $nativePack } '*Invalid MainNodeIds*'
    $checks.Add('Mock-native pack inspection/export preserves exact metadata, validates ordering and coverage, rejects malformed Shape Data and blocks merge downgrade/page changes and invalid post-edit exports without COM.')
    $referencePack=$packOriginal | ConvertFrom-Json
    $referencePack.presentationProfile='reference'
    $nativeReference=New-MockPackDocument $referencePack
    $nativeEnterprise=New-MockPackDocument $pack
    $savedStylePath=$script:enterpriseStylePath
    $script:enterpriseStylePath={param($InputModel) throw 'Enterprise-only export style gate was invoked.'}
    try {
        $referenceRoundtrip=Export-Model $nativeReference
        Assert ($referenceRoundtrip.presentationProfile -ceq 'reference' -and
            $referenceRoundtrip.outputContract -ceq 'architecture-pack-v1.6' -and
            ($referenceRoundtrip.pages.view -join ',') -ceq 'main,hardening,flowchart') 'Native reference pack export preserves profile and contract without invoking enterprise-only presentation policy.'
        Confirm-NativeArchitecturePack $nativeReference
        Assert-Throws { Export-Model $nativeEnterprise } '*Enterprise-only export style gate was invoked*'
        $nativeReference.Pages.Item(3).Shapes.Entries[0].CellMap.Remove('Prop.MainEdgeIds')
        Assert-Throws { Confirm-NativeArchitecturePack $nativeReference } '*mainEdgeIds string arrays*'
    } finally { $script:enterpriseStylePath=$savedStylePath }
    $checks.Add('Reference-profile native pack export and CRUD guards skip enterprise-only presentation policy while retaining contract, ordering and MAIN coverage checks; enterprise exports still invoke their style gate.')
    $portableDocument=New-MockPackDocument $pack
    Set-MockProperty $portableDocument.DocumentSheet 'PortableRenderer' 'azure-visio-1.6'
    $portablePage=$portableDocument.Pages.Item(1)
    $boundary=@($portablePage.Shapes.Entries | Where-Object { (Read-Property $_ 'AvId') -eq 'system' })[0]
    $child=@($portablePage.Shapes.Entries | Where-Object { (Read-Property $_ 'AvId') -eq 'function' })[0]
    $boundary.CellMap.Remove('User.msvStructureType')
    Set-MockProperty $boundary 'ContainerStyle' 'boundary'
    $boundary.CellMap['FillPattern']=[pscustomobject]@{ResultIU=0}
    Set-MockProperty $child 'ParentId' 'system'
    foreach ($native in @($boundary,$child)) {
        $native.CellMap['LocPinX']=[pscustomobject]@{ResultIU=$native.CellMap.Width.ResultIU/2}
        $native.CellMap['LocPinY']=[pscustomobject]@{ResultIU=$native.CellMap.Height.ResultIU/2}
    }
    $boundaryRecord=Inspect-Shape $boundary
    Assert (-not $boundaryRecord.isContainer -and -not (Test-Field $boundaryRecord 'memberIds') -and
        (Test-PortableBoundary $boundaryRecord 'azure-visio-1.6')) 'Portable rectangles remain semantic boundaries, never fabricated native containers or native member lists.'
    $portableExport=Export-Model $portableDocument
    $exportedBoundary=@($portableExport.pages[0].nodes | Where-Object id -eq 'system')[0]
    $exportedChild=@($portableExport.pages[0].nodes | Where-Object id -eq 'function')[0]
    Assert ($exportedBoundary.kind -eq 'container' -and $exportedBoundary.containerStyle -eq 'boundary' -and
        $exportedChild.parent -eq 'system' -and @($portableExport.warnings | Where-Object { $_ -like '*semantic-boundary mode*' }).Count -eq 1) 'Portable boundary export recovers validated semantic parent metadata and explicitly warns that native membership is not claimed.'
    foreach ($parentId in @('missing-parent','SYSTEM','system;unsafe')) {
        Set-MockProperty $child 'ParentId' $parentId
        Assert-Throws { Export-Model $portableDocument } '*'
    }
    Set-MockProperty $child 'ParentId' 'system'
    $originalChildX=$child.CellMap.PinX.ResultIU
    $child.CellMap.PinX.ResultIU=0.8
    Assert-Throws { Export-Model $portableDocument } '*outside parent*refusing to drop or change ParentId*'
    $child.CellMap.PinX.ResultIU=$originalChildX
    Set-MockProperty $boundary 'ParentId' 'system'
    Assert-Throws { Export-Model $portableDocument } '*Container cycle*'
    Set-MockProperty $boundary 'ParentId' ''
    Set-MockProperty $portableDocument.DocumentSheet 'PortableRenderer' 'unknown-renderer'
    Assert-Throws { Export-Model $portableDocument } '*containerStyle must be boundary and is only supported on containers*'
    Set-MockProperty $portableDocument.DocumentSheet 'PortableRenderer' 'azure-visio-1.6'
    $request='{"schemaVersion":1,"updates":[{"pageId":0,"shapeId":"system","set":{"width":9}}]}' | ConvertFrom-Json
    Assert-Throws { Get-UpdatePlan $portableDocument $request } '*nonempty portable semantic boundaries*'
    $request='{"schemaVersion":1,"updates":[{"pageId":0,"shapeId":"function","set":{"x":1}}]}' | ConvertFrom-Json
    Assert-Throws { Get-UpdatePlan $portableDocument $request } '*no implicit semantic reparenting*'
    $request.updates[0].set.x=5.2
    Assert (@(Get-UpdatePlan $portableDocument $request).Count -eq 1 -and
        $child.CellMap.PinX.ResultIU -eq $originalChildX -and (Read-Property $child 'ParentId') -eq 'system') 'Safe child movement preflights without changing ownership; failed preflights leave native geometry untouched.'
    $request='{"schemaVersion":1,"updates":[{"pageId":0,"shapeId":"function","set":{"parent":"other"}}]}' | ConvertFrom-Json
    Assert-Throws { Confirm-Changes $request 'Update' } '*Unsupported set property*'
    $inner=New-MockShape 998
    foreach ($pair in @{AvId='inner';Kind='container';ContainerStyle='boundary';BoundaryType='system';ParentId='system'}.GetEnumerator()) { Set-MockProperty $inner $pair.Key $pair.Value }
    $inner.Text='Inner boundary'; $inner.CellMap.PinX.ResultIU=5; $inner.CellMap.PinY.ResultIU=3
    $inner.CellMap.Width.ResultIU=4; $inner.CellMap.Height.ResultIU=2
    $inner.CellMap['FillPattern']=[pscustomobject]@{ResultIU=0}
    $portablePage.Shapes=New-MockCollection (@($portablePage.Shapes.Entries)+@($inner))
    Set-MockProperty $child 'ParentId' 'inner'
    $nested=Export-Model $portableDocument
    Assert (@($nested.pages[0].nodes | Where-Object id -eq 'inner')[0].parent -eq 'system' -and
        @($nested.pages[0].nodes | Where-Object id -eq 'function')[0].parent -eq 'inner') 'Nested portable boundary hierarchy is preserved without flattening.'
    $nativeParent=New-MockShape 999
    $nativeParent.Text='Native ownership'
    $nativeParent.CellMap.PinX.ResultIU=5.5; $nativeParent.CellMap.PinY.ResultIU=4
    $nativeParent.CellMap.Width.ResultIU=10; $nativeParent.CellMap.Height.ResultIU=7
    Set-MockProperty $nativeParent 'AvId' 'native-parent'; Set-MockProperty $nativeParent 'Kind' 'container'
    Set-MockProperty $nativeParent 'BoundaryType' 'system'
    $containerCell=[pscustomobject]@{}
    $containerCell | Add-Member ScriptMethod ResultStr {param($units) return 'Container'}
    $nativeParent.CellMap['User.msvStructureType']=$containerCell
    $membership=[pscustomobject]@{Members=@($child.ID)}
    $membership | Add-Member ScriptMethod GetMemberShapes {param($flags) return $this.Members}
    $nativeParent | Add-Member -NotePropertyName ContainerProperties -NotePropertyValue $membership
    $portablePage.Shapes=New-MockCollection (@($portablePage.Shapes.Entries)+@($nativeParent))
    Set-MockProperty $child 'ParentId' 'stale-portable-parent'
    $authoritative=Export-Model $portableDocument
    Assert (@($authoritative.pages[0].nodes | Where-Object id -eq 'function')[0].parent -eq 'native-parent' -and
        @($authoritative.warnings | Where-Object { $_ -like '*native container membership overrides stored ParentId*' }).Count -eq 1) 'Real native membership overrides stale portable metadata.'
    $membership.Members=@()
    Set-MockProperty $child 'ParentId' 'native-parent'
    $removed=Export-Model $portableDocument
    Assert (-not (Test-Field @($removed.pages[0].nodes | Where-Object id -eq 'function')[0] 'parent') -and
        @($removed.warnings | Where-Object { $_ -like '*not supported by native membership*' }).Count -eq 1) 'Missing genuine native membership is not invented from ParentId metadata.'
    $checks.Add('Portable semantic boundaries preserve kinds and nested ParentId ownership with explicit warnings, enforce IDs/bounds/cycles, protect nonempty boundary geometry and child ownership, and defer to genuine native membership without fabricating containers.')
    function New-RouteMock([double[]]$Points) {
        $edge=New-MockShape 700 $true
        $edge | Add-Member -NotePropertyName Coordinates -NotePropertyValue $Points
        $edge | Add-Member -NotePropertyName GeometryCount -NotePropertyValue 1
        $edge | Add-Member -NotePropertyName RowTags -NotePropertyValue @{}
        $edge | Add-Member ScriptMethod RowCount {param($section) return $this.Coordinates.Count/2+1}
        $edge | Add-Member ScriptMethod RowType {param($section,$row) if ($this.RowTags.ContainsKey($row)) { return $this.RowTags[$row] }; if ($row -eq 0) { return 137 }; if ($row -eq 1) { return 138 }; return 139}
        $edge | Add-Member ScriptMethod CellsSRC {param($section,$row,$cell) if ($cell -eq 2) { return [pscustomobject]@{ResultIU=0.2} }; return [pscustomobject]@{ResultIU=$this.Coordinates[($row-1)*2+$cell]}}
        $edge | Add-Member -NotePropertyName TransformAngle -NotePropertyValue 0.0
        $edge | Add-Member ScriptMethod XYToPage {
            param($x,$y,$outX,$outY)
            $outX.Value=$x*[Math]::Cos($this.TransformAngle)-$y*[Math]::Sin($this.TransformAngle)
            $outY.Value=$x*[Math]::Sin($this.TransformAngle)+$y*[Math]::Cos($this.TransformAngle)
        }
        return $edge
    }
    $orthogonal=New-RouteMock @(1.5,5,3.5,5,3.5,1.5,5,1.5)
    $route=Get-NativeRouteGeometry $orthogonal
    Assert ($route.available -and $route.orthogonal -and $route.segments.Count -eq 3 -and
        ($route.segments.axis -join ',') -eq 'horizontal,vertical,horizontal') 'Actual nonaligned connector strokes are inspected as page-coordinate right-angle segments.'
    $diagonal=New-RouteMock @(1.5,5,5,1.5)
    $diagonal.CellMap['ShapeRouteStyle']=[pscustomobject]@{ResultIU=1}
    Assert ((Get-NativeRouteGeometry $diagonal).diagonalCount -eq 1) 'An orthogonal ShapeRouteStyle setting cannot conceal an actually diagonal stroke.'
    $curved=New-RouteMock @(0,0,1,0,2,0)
    $curved.RowTags[3]=140
    $curveReport=Get-NativeRouteGeometry $curved
    Assert ($curveReport.available -and -not $curveReport.orthogonal -and $curveReport.curvedCount -eq 1 -and
        $curveReport.segments[1].rowType -eq 140) 'ArcTo line jumps are flagged from actual Geometry rows rather than reduced to a misleading horizontal chord.'
    $rotated=New-RouteMock @(0,0,2,0); $rotated.TransformAngle=[Math]::PI/4
    Assert ((Get-NativeRouteGeometry $rotated).diagonalCount -eq 1) 'Local horizontal geometry rotated on the page is correctly reported as diagonal.'
    Assert (-not (Get-NativeRouteGeometry (New-MockShape 701 $true)).available) 'Unavailable path geometry is reported as unverified, never guessed from routing settings.'
    $service=[pscustomobject]@{visioId=1;id='producer';kind='card';serviceBounds=[pscustomobject]@{left=0.5;right=1.5;bottom=4.5;top=5.5}}
    $destination=[pscustomobject]@{visioId=2;id='blob';kind='card';serviceBounds=[pscustomobject]@{left=5;right=6;bottom=1;top=2}}
    $edgeRecord=[pscustomobject]@{
        visioId=3;id='producer-blob';kind='edge';source='producer';target='blob';sourceVisioIds=@(1);targetVisioIds=@(2)
        beginX=1.5;beginY=5;endX=5;endY=1.5;sourceSide='right';targetSide='left';sourcePosition=0.5;targetPosition=0.5
        relationship='telemetry';direction='forward';routeStyle='orthogonal';routeGeometry=$route
    }
    $routePage=[pscustomobject]@{shapes=@($service,$destination,$edgeRecord)}
    Assert (Get-NativeRoutingReport $routePage).valid 'Nonaligned telemetry icon-to-label endpoints with real full-shape glue and orthogonal bends are accepted.'
    $edgeRecord.routeGeometry=Get-NativeRouteGeometry $diagonal
    Assert (@((Get-NativeRoutingReport $routePage).issues | Where-Object code -eq 'DiagonalRoute').Count -eq 1) 'Page route validation inspects actual diagonal geometry rather than routeStyle.'
    $edgeRecord.routeGeometry=$route
    $edgeRecord.routeGeometry=Get-NativeRouteGeometry (New-RouteMock @(1.5,5,0.1,5,0.1,1.5,5,1.5))
    Assert (@((Get-NativeRoutingReport $routePage).issues | Where-Object code -eq 'EndpointDirection').Count -eq 1) 'A visually orthogonal connector leaving its right-side port to the left is invalid; inward Visio connection vectors must be the inverse of the outward route escape.'
    $edgeRecord.routeGeometry=$route
    $edgeRecord.target='invented-target'
    Assert (@((Get-NativeRoutingReport $routePage).issues | Where-Object code -eq 'EndpointIdentity').Count -eq 1) 'Glue to a different service cannot satisfy the declared relationship.'
    $edgeRecord.target='blob'; $edgeRecord.beginX=3
    Assert (@((Get-NativeRoutingReport $routePage).issues | Where-Object code -eq 'EndpointWhitespace').Count -eq 1) 'Endpoint on an oversized invisible group edge is rejected when it does not reach the glyph.'
    $edgeRecord.beginX=1.5
    $edgeRecord.sourceSide='left'
    Assert (@((Get-NativeRoutingReport $routePage).issues | Where-Object code -eq 'SelectedPortMismatch').Count -eq 1) 'A different visible side does not satisfy the selected port.'
    $edgeRecord.sourceSide='right'
    $edgeRecord.routeGeometry=Get-NativeRouteGeometry (New-RouteMock @(1.5,5,3.5,5,3.5,1.5,4.8,1.5))
    Assert (@((Get-NativeRoutingReport $routePage).issues | Where-Object code -eq 'StrokeEndpointMismatch').Count -eq 1) 'Visible stroke must actually reach the glued endpoint, not merely store the correct Begin/End cells.'
    $edgeRecord.routeGeometry=$route
    $obstacle=[pscustomobject]@{visioId=4;id='unrelated-caption';kind='card'
        serviceBounds=[pscustomobject]@{left=7;right=8;bottom=3;top=4}
        captionBounds=[pscustomobject]@{left=3;right=4;bottom=3;top=3.4}}
    $routePage.shapes+=@($obstacle)
    Assert (@((Get-NativeRoutingReport $routePage).issues | Where-Object code -eq 'RouteObstruction').Count -eq 1) 'Orthogonal shortcuts through another service caption are rejected.'
    $routePage.shapes=@($service,$destination,$edgeRecord)
    $service | Add-Member -NotePropertyName cardStyle -NotePropertyValue 'icon'
    $service | Add-Member -NotePropertyName captionBounds -NotePropertyValue ([pscustomobject]@{left=0.5;right=1.5;bottom=3.7;top=4.1})
    $edgeRecord.beginX=1; $edgeRecord.beginY=4.5; $edgeRecord.sourceSide='bottom'
    $edgeRecord.routeGeometry=Get-NativeRouteGeometry (New-RouteMock @(1,4.5,1,1.5,5,1.5))
    Assert (@((Get-NativeRoutingReport $routePage).issues | Where-Object code -eq 'RouteObstruction').Count -eq 1) 'A bottom glyph attachment may not exit through its own service caption.'
    $edgeRecord.routeGeometry=Get-NativeRouteGeometry (New-RouteMock @(1,4.5,2,4.5,2,1.5,5,1.5))
    $wrongBottom=Get-NativeRoutingReport $routePage
    Assert (@($wrongBottom.issues | Where-Object code -eq 'SelectedPortMismatch').Count -eq 1 -and
        $wrongBottom.attachments[0].region -eq 'glyph' -and $wrongBottom.attachments[0].expectedRegion -eq 'caption') 'A declared icon bottom port must use the approved caption-bottom convention, even if a glyph-bottom route avoids text; inspection does not mislabel the actual glyph attachment.'
    $edgeRecord.beginY=3.7
    $edgeRecord.routeGeometry=Get-NativeRouteGeometry (New-RouteMock @(1,3.7,1,1.5,5,1.5))
    $captionRoute=Get-NativeRoutingReport $routePage
    Assert ($captionRoute.valid -and $captionRoute.attachments[0].region -eq 'caption') 'Native and portable bottom ports on the actual caption lower edge are valid semantic service attachments, not falsely reported as misglued.'
    $edgeRecord.sourcePosition=0.25; $edgeRecord.beginX=0.75
    $edgeRecord.routeGeometry=Get-NativeRouteGeometry (New-RouteMock @(0.75,3.7,0.75,1.5,5,1.5))
    Assert (Get-NativeRoutingReport $routePage).valid 'Bottom position fractions use the actual caption width, independently of glyph and invisible group widths.'
    $edgeRecord.beginX=0
    $edgeRecord.routeGeometry=Get-NativeRouteGeometry (New-RouteMock @(0,3.7,0,1.5,5,1.5))
    Assert (@((Get-NativeRoutingReport $routePage).issues | Where-Object code -eq 'EndpointWhitespace').Count -eq 1) 'A point on broad group whitespace outside the actual caption is not excused by the bottom-port convention.'
    $edgeRecord.beginX=0.75
    $edgeRecord.routeGeometry=Get-NativeRouteGeometry (New-RouteMock @(0.75,3.7,0.75,4.3,3.5,4.3,3.5,1.5,5,1.5))
    Assert (@((Get-NativeRoutingReport $routePage).issues | Where-Object code -eq 'RouteObstruction').Count -eq 1) 'Even a correctly attached caption-bottom port cannot send its stroke upward through its own text.'
    $edgeRecord.sourcePosition=0.5
    $edgeRecord.beginY=5
    $edgeRecord.beginX=0.5; $edgeRecord.endX=6; $edgeRecord.sourceSide='left'; $edgeRecord.targetSide='right'
    $edgeRecord.direction='backward'
    $edgeRecord.routeGeometry=Get-NativeRouteGeometry (New-RouteMock @(0.5,5,0.1,5,0.1,0.3,6.4,0.3,6.4,1.5,6,1.5))
    Assert ((Get-NativeRoutingReport $routePage).valid -and $edgeRecord.source -eq 'producer' -and $edgeRecord.target -eq 'blob' -and
        $edgeRecord.direction -eq 'backward') 'Return/back edges route around services without inventing endpoints or reversing relationship semantics.'
    $checks.Add('Actual stroke routing checks detect diagonals, curved/rotated paths, wrong glue/ports, whitespace, disconnected strokes, and own/unrelated caption crossings; fitted-caption bottom ports, nonaligned telemetry and back edges retain semantics.')

    $portRoot=New-MockShape 710
    $portRoot.CellMap.Width.ResultIU=4; $portRoot.CellMap.Height.ResultIU=2
    Set-MockProperty $portRoot 'AvId' 'wide-service'; Set-MockProperty $portRoot 'CardStyle' 'icon'
    $portGlyph=New-MockShape 711
    $portGlyph.CellMap.PinX.ResultIU=2; $portGlyph.CellMap.PinY.ResultIU=1.5
    $portGlyph.CellMap.Width.ResultIU=0.6; $portGlyph.CellMap.Height.ResultIU=0.4
    Set-MockProperty $portGlyph 'AvRole' 'icon'
    $portCaption=New-MockShape 713
    $portCaption.CellMap.PinX.ResultIU=2; $portCaption.CellMap.PinY.ResultIU=0.4
    $portCaption.CellMap.Width.ResultIU=1.2; $portCaption.CellMap.Height.ResultIU=0.4
    $portCaption | Add-Member ScriptMethod BoundingBox {
        param($flags,$outLeft,$outBottom,$outRight,$outTop)
        $outLeft.Value=1.44; $outRight.Value=2.56; $outBottom.Value=0.24; $outTop.Value=0.56
    }
    Set-MockProperty $portCaption 'AvRole' 'caption'
    $portRoot.Shapes=New-MockCollection @($portGlyph,$portCaption)
    $portRoot | Add-Member ScriptMethod AddNamedRow {
        param($section,$row,$tag)
        foreach ($field in @('X','Y','DirX','DirY','Type')) { $this.CellMap["Connections.$row.$field"]=[pscustomobject]@{FormulaU='';OwnerId=$this.ID} }
        return 0
    }
    $glueCell=[pscustomobject]@{GluedTo=$null}
    $glueCell | Add-Member ScriptMethod GlueTo {param($target) $this.GluedTo=$target}
    Glue-End $glueCell $portRoot 'bottom' 0.5
    $bottomPort=Get-ServicePort $portRoot 'bottom' 0.5
    Assert ([Math]::Abs($bottomPort.u-0.5) -lt 0.0001 -and [Math]::Abs($bottomPort.v-0.1) -lt 0.0001 -and
        $bottomPort.dirY -eq -1 -and $bottomPort.region -eq 'caption' -and $glueCell.GluedTo.OwnerId -eq 710) 'Bottom icon port is on the actual fitted-caption lower edge, not broad group whitespace; glue remains on the full service group.'
    Assert ([Math]::Abs((Get-ServicePort $portRoot 'bottom' 0.25).u-0.425) -lt 0.0001) 'Bottom port fractions track caption width rather than glyph or invisible anchor width.'
    Assert ($portRoot.CellMap['Connections.AvPortbottom0_5.X'].FormulaU -like '*Sheet.713!Width*' -and
        $portRoot.CellMap['Connections.AvPortbottom0_5.Y'].FormulaU -like '*Sheet.713!PinY*') 'Native caption ports dynamically follow caption text-width/height edits rather than retaining stale normalized coordinates.'
    Assert ($portRoot.CellMap['Connections.AvPortbottom0_5.DirY'].FormulaU -eq '1') 'Type=0 native bottom connection has an inward positive-Y vector so Visio routes its stroke outward/down, not upward through the caption.'
    $rightPort=Get-ServicePort $portRoot 'right' 0.5
    Assert ([Math]::Abs($rightPort.u-0.575) -lt 0.0001 -and $rightPort.dirX -eq 1) 'Right icon port ignores invisible group whitespace.'
    Glue-End $glueCell $portRoot 'bottom' 0.5
    Assert (@($portRoot.CellMap.Keys | Where-Object { $_ -like 'Connections.*.X' }).Count -eq 1) 'Stable named ports are reused rather than adding drifting connection points.'
    $labelRoot=New-MockShape 712
    Set-MockProperty $labelRoot 'CardStyle' 'label'
    $labelCaption=New-MockShape 714
    $labelCaption.CellMap.PinX.ResultIU=1; $labelCaption.CellMap.PinY.ResultIU=0.5
    $labelCaption.CellMap.Width.ResultIU=1; $labelCaption.CellMap.Height.ResultIU=0.4
    $labelCaption | Add-Member ScriptMethod BoundingBox {
        param($flags,$outLeft,$outBottom,$outRight,$outTop)
        $outLeft.Value=0.54; $outRight.Value=1.46; $outBottom.Value=0.34; $outTop.Value=0.66
    }
    Set-MockProperty $labelCaption 'AvRole' 'caption'
    $labelRoot.Shapes=New-MockCollection @($labelCaption)
    Assert ((Get-ServicePort $labelRoot 'left' 0.25).u -eq 0.25 -and
        [Math]::Abs((Get-ServicePort $labelRoot 'left' 0.25).v-0.4) -lt 0.0001) 'Label-only services retain explicit positions on the fitted text perimeter, not the invisible group.'
    Assert ((Get-ServicePort $labelRoot 'right').u -eq 0.75 -and
        [Math]::Abs((Get-ServicePort $labelRoot 'top').v-0.7) -lt 0.0001 -and
        [Math]::Abs((Get-ServicePort $labelRoot 'bottom').v-0.3) -lt 0.0001) 'All four label-only sides use actual caption bounds.'
    $labelCaption.CellMap.Width.ResultIU=2
    $measured=Get-NativeCaptionBounds $labelCaption -Local
    Assert ([Math]::Abs($measured.left-0.5) -lt 0.0001 -and [Math]::Abs($measured.right-1.5) -lt 0.0001 -and
        [Math]::Abs((Get-ServicePort $labelRoot 'left').u-0.25) -lt 0.0001) 'Older portable caption children spanning the whole node are measured from native text extent, not mistaken for visible text or broad whitespace attachment.'
    $measuredPort=Get-ServicePort $labelRoot 'bottom' 0.15
    Assert ($measuredPort.xFormula -notmatch 'TEXTWIDTH' -and $measuredPort.xFormula -like '*Width*0.5*' -and
        [Math]::Abs($measuredPort.u-0.325) -lt 0.0001) 'Caption port formulas encode exactly the inspected BoundingBox extent rather than a different TEXTWIDTH measurement; noncentral .15 positions agree.'
    $setLabelFunction=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Set-NodeLabel'},$true))[0]
    Assert ($setLabelFunction.Extent.Text -match 'TEXTWIDTH\(TheText' -and
        $setLabelFunction.Extent.Text -match "CellsU\('Height'\).ResultIUForce=\`$layout.captionHeight") 'Caption rectangles fit native measured text width and actual caption lines, including label-only cards.'
    $addEdgeFunction=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Add-Edge'},$true))[0]
    Assert ($addEdgeFunction.Extent.Text -match "ShapeRouteStyle'\)\.FormulaU = '1'" -and
        $addEdgeFunction.Extent.Text -notmatch '\$route=''straight''') 'New connectors always normalize authored routing to right angles; no icon/label straight shortcut remains.'
    $checks.Add('Full-service named cardinal ports use glyph left/right/top and fitted-caption bottom, preserve label positions, follow caption edits dynamically, and normalize new routing instead of diagonal straight shortcuts.')

    $boundaryStart=[pscustomobject]@{x=11.14;y=2.24}
    $boundaryEnd=[pscustomobject]@{x=6.1;y=1.75}
    $boundaryObstacles=@(
        [pscustomobject]@{left=2;right=10.2;bottom=1.75;top=9.25},
        [pscustomobject]@{left=11.14;right=11.86;bottom=1.88;top=2.60},
        [pscustomobject]@{left=10.788;right=12.212;bottom=1.25;top=1.747}
    )
    $boundaryRoute=Find-BoundaryRoute $boundaryStart $boundaryEnd 'left' 'bottom' $boundaryObstacles 18 11
    Assert ($boundaryRoute[1].x -lt $boundaryStart.x -and $boundaryRoute[-2].y -lt $boundaryEnd.y) 'Native container fallback leaves the glyph and approaches the boundary from the selected exterior sides.'
    for ($i=1; $i -lt $boundaryRoute.Count; $i++) {
        $a=$boundaryRoute[$i-1]; $b=$boundaryRoute[$i]
        Assert ([Math]::Abs($a.x-$b.x) -lt 0.000001 -or [Math]::Abs($a.y-$b.y) -lt 0.000001) 'Boundary gutter route is genuinely orthogonal.'
        $segment=[pscustomobject]@{fromX=$a.x;fromY=$a.y;toX=$b.x;toY=$b.y;axis=$(if ($a.y -eq $b.y) {'horizontal'} else {'vertical'})}
        foreach ($bounds in $boundaryObstacles) { Assert (-not (Test-RouteCrossing $segment $bounds)) 'Boundary gutter avoids the container interior and all glyph/caption obstacles.' }
    }
    Assert-Throws { Find-BoundaryRoute $boundaryStart $boundaryEnd 'left' 'bottom' @([pscustomobject]@{left=0;right=18;bottom=0;top=11}) 18 11 } '*No clear orthogonal boundary approach*'
    $checks.Add('Native-container exterior gutter routes preserve selected endpoint directions and avoid glyphs/captions; blocked layouts fail rather than creating shortcut strokes.')

    $rerouteEdge=New-MockShape 720 $true; Set-MockProperty $rerouteEdge 'Kind' 'edge'
    foreach ($cell in @('ShapeRouteStyle','ConFixedCode','ConLineRouteExt','ConLineJumpCode','Rounding')) { $rerouteEdge.CellMap[$cell]=[pscustomobject]@{FormulaU='2';ResultIU=2} }
    $rerouteNode=New-MockShape 721; Set-MockProperty $rerouteNode 'Kind' 'card'
    $routeSelection=[pscustomobject]@{Selected=[Collections.Generic.List[int]]::new();LayoutCalls=0}
    $routeSelection | Add-Member ScriptMethod Select {param($shape,$flags) $this.Selected.Add([int]$shape.ID)}
    $routeSelection | Add-Member ScriptMethod Layout {$this.LayoutCalls++}
    $reroutePage=[pscustomobject]@{Shapes=(New-MockCollection @($rerouteNode,$rerouteEdge));Selection=$routeSelection}
    $reroutePage | Add-Member ScriptMethod CreateSelection {param($type,$mode,$data) return $this.Selection}
    $rerouteDocument=[pscustomobject]@{Pages=(New-MockCollection @($reroutePage))}
    Complete-NativeRouting $rerouteDocument -Force
    Assert ($routeSelection.LayoutCalls -eq 1 -and $routeSelection.Selected.Count -eq 1 -and $routeSelection.Selected[0] -eq 720 -and
        $rerouteEdge.CellMap.ShapeRouteStyle.FormulaU -eq '1' -and $rerouteEdge.CellMap.ConFixedCode.FormulaU -eq '0' -and
        $rerouteEdge.CellMap.ConLineRouteExt.FormulaU -eq '1' -and $rerouteNode.CellMap.PinX.ResultIU -eq 2) 'Final layout reroutes connectors only with freely routed straight orthogonal segments, without moving nodes.'
    $reroutePage.Shapes=New-MockCollection @($rerouteNode)
    Complete-NativeRouting $rerouteDocument -Force
    Assert ($routeSelection.LayoutCalls -eq 1) 'Empty connector selections never call Layout, which would otherwise relayout every shape.'
    Set-MockProperty $portRoot 'Kind' 'card'
    Set-MockProperty $rerouteEdge 'AvId' 'refresh-edge'
    Set-MockProperty $rerouteEdge 'SourceId' 'wide-service'
    Set-MockProperty $rerouteEdge 'SourceSide' 'bottom'
    Set-MockProperty $rerouteEdge 'SourcePosition' '0.15'
    $rerouteEdge.CellMap.BeginX | Add-Member NoteProperty GluedTo $null
    $rerouteEdge.CellMap.BeginX | Add-Member ScriptMethod GlueTo {param($target) $this.GluedTo=$target}
    $rerouteEdge.Connects=New-MockCollection @([pscustomobject]@{FromCell=[pscustomobject]@{Name='BeginX'};ToSheet=$portRoot})
    $reroutePage.Shapes=New-MockCollection @($portRoot,$rerouteEdge)
    Complete-NativeRouting $rerouteDocument -Force
    Assert ($rerouteEdge.CellMap.BeginX.GluedTo.OwnerId -eq $portRoot.ID -and
        $portRoot.CellMap['Connections.AvPortbottom0_15.X'].FormulaU -like '*0.15*' -and
        $portRoot.CellMap['Connections.AvPortbottom0_15.DirY'].FormulaU -eq '1') 'Final rerouting refreshes measured noncentral caption ports and inward vectors after caption edits without changing the full service glue.'
    $rerouteEdge.Connects=New-MockCollection @([pscustomobject]@{FromCell=[pscustomobject]@{Name='BeginX'};ToSheet=$rerouteNode})
    Assert-Throws { Complete-NativeRouting $rerouteDocument -Force } '*actual full-service glue differs from its declared endpoint*'
    $checks.Add('Caption-port refresh preserves noncentral positions/full-group identity and rejects stale or incorrect endpoint glue before rerouting.')
    $policyDocument=New-MockPackDocument $pack
    Set-MockProperty $policyDocument.DocumentSheet 'RoutingPolicy' 'orthogonal-v1.6'
    Assert-Throws { Export-Model $policyDocument } '*Native architecture routing validation failed*RouteGeometryUnavailable*'
    $flowPolicy=New-MockPackDocument $pack
    Set-MockProperty $flowPolicy.DocumentSheet 'RoutingPolicy' 'orthogonal-v1.6'
    $flowPolicy.Pages.Item(1).Shapes=New-MockCollection @($flowPolicy.Pages.Item(1).Shapes.Entries | Where-Object {-not $_.OneD})
    Assert-Throws { Export-Model $flowPolicy } '*Main flowchart*RouteGeometryUnavailable*'
    $policyDocument.DocumentSheet.CellMap.Remove('Prop.RoutingPolicy')
    Assert ((Export-Model $policyDocument).pages.Count -eq 3) 'Existing documents without the new routing policy remain readable/exportable; new-policy documents cannot silently pass unknown actual geometry.'
    $checks.Add('Connector-only final rerouting preserves node placement, clears curved/diagonal routing settings, never invokes page-wide layout from an empty selection, and enforces actual routing on flowchart/notes pages too.')
} finally {
    Remove-Item -LiteralPath $fixtures -Recurse -Force
}
$draftTests = & (Join-Path $PSScriptRoot 'Draft-Import.Tests.ps1') -OutputDirectory $OutputDirectory
Assert ($draftTests.passed -and $draftTests.fixtureDirectoryRemoved -and $draftTests.checkGroups -gt 0 -and
    $draftTests.checkGroups -eq $draftTests.checks.Count) 'Draft importer filesystem tests pass and clean their own fixtures.'
foreach ($check in $draftTests.checks) { $checks.Add('Draft importer: '+$check) }
$referenceTests = & (Join-Path $PSScriptRoot 'Reference-Fidelity.Tests.ps1') -OutputDirectory $OutputDirectory
Assert ($referenceTests.passed -and $referenceTests.fixtureDirectoryRemoved -and $referenceTests.checkGroups -gt 0) 'Reference fidelity tests pass and clean fixtures.'
foreach ($check in $referenceTests.checks) { $checks.Add('Reference fidelity: '+$check) }
if ($ReferenceModelPath) {
    $fidelityHelper=Join-Path $PSScriptRoot 'Reference-Fidelity.ps1'
    $positive=& $fidelityHelper -ModelPath $ReferenceModelPath -ReferencePath $ReferencePath | ConvertFrom-Json
    Assert $positive.valid 'Actual user-reference model satisfies its independently extracted source contract.'
    $sourceContract=Get-Content -LiteralPath $ReferencePath -Raw | ConvertFrom-Json
    $bad=Get-Content -LiteralPath $ReferenceModelPath -Raw | ConvertFrom-Json
    $page=@($bad.pages | Where-Object {$_.name -eq $sourceContract.referencePage})[0]
    $victim=@($page.nodes | Where-Object {$_.kind -eq 'card'})[0]
    $page.nodes=@($page.nodes | Where-Object {$_.id -ne $victim.id})
    $page.edges=@($page.edges | Where-Object {$_.source -ne $victim.id -and $_.target -ne $victim.id})
    $page.nodes += [pscustomobject]@{id='unapproved-media-bridge';sourceId='unapproved-media-bridge';kind='card';label='Media bridge [CUSTOM]';x=1;y=1;width=1;height=1}
    $negativePath=Join-Path $OutputDirectory 'deliberate-reference-drift.json'
    $bad | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $negativePath -Encoding UTF8
    $negative=& $fidelityHelper -ModelPath $negativePath -ReferencePath $ReferencePath -ReportOnly | ConvertFrom-Json
    Assert (-not $negative.valid) 'Reference gate rejects an omitted source component and an unapproved custom adapter.'
    $bad=Get-Content -LiteralPath $ReferenceModelPath -Raw | ConvertFrom-Json
    @($bad.pages | Where-Object {$_.name -eq $sourceContract.referencePage})[0].name='Redesign replacing the reference'
    $bad | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $negativePath -Encoding UTF8
    $negative=& $fidelityHelper -ModelPath $negativePath -ReferencePath $ReferencePath -ReportOnly | ConvertFrom-Json
    Assert (-not $negative.valid) 'A differently named proposal cannot replace the required source-faithful page.'
    Remove-Item -LiteralPath $negativePath
    $checks.Add('Actual reference passes; deliberate source omission/custom-adapter drift and replacement by a proposal are both rejected.')
}
if ($PortableOnly) {
    [pscustomobject]@{passed=$checks.Count; checks=$checks} | ConvertTo-Json -Depth 4
    return
}

$IconDirectory=$requestedIconDirectory
$script:catalog=$null
if ($IconDirectory) { $IconDirectory=Get-AbsolutePath $IconDirectory; $script:helperIconDirectory=$IconDirectory }
$svgEntry=$null; $svgPath=''
if ($CrudOnly -or $EnterpriseOnly) {
    if (-not $IconDirectory) { throw 'Native icon tests require -IconDirectory with at least one audited usable catalog SVG.' }
    $svgEntries=@((Get-IconCatalog).Values | Where-Object {
        (Get-Value $_ 'usable' $false) -is [bool] -and (Get-Value $_ 'usable' $false) -and
            [IO.Path]::GetExtension((Get-Value $_ 'path' '')) -ieq '.svg'
    } | Sort-Object id)
    if (-not $svgEntries.Count) { throw 'Native icon tests require at least one catalog SVG with usable:true.' }
    $preferred=@($svgEntries | Where-Object { $_.id -eq 'azure-440773d3fa75' })
    $svgEntry=$(if ($preferred.Count) { $preferred[0] } else { $svgEntries[0] })
    $svgPath=Resolve-IconRef $svgEntry.id
    $crudModel.pages[0].nodes += [pscustomobject]@{
        id='svg-catalog';kind='card';label=(Get-Value $svgEntry 'name' $svgEntry.id);iconRef=$svgEntry.id
        x=3.0;y=3.0;width=2.2;height=1.3;state='Fixture';sourceRef=(Get-Value $svgEntry 'sourcePage' '')
        requirementIds=@('ICON-CATALOG');confidence='High'
    }
    $crudModel | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $crudModelPath -Encoding UTF8
    [void](Run-Helper @{Action='Validate';ModelPath=$crudModelPath})
}
$app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
$script:syncRoots = @()
$environment = Join-Path $PSScriptRoot 'environment.json'
if (Test-Path -LiteralPath $environment) { $script:syncRoots = @(Read-Environment $environment) }
$beforeDocuments = @{}
for ($i=1; $i -le $app.Documents.Count; $i++) {
    $open = $app.Documents.Item($i)
    $beforeDocuments[$open.FullName] = $open.Saved
}
if ($EnterpriseOnly) {
    $enterprisePath=Join-Path $OutputDirectory 'enterprise-model.json'
    $drawing=Join-Path $OutputDirectory ('Enterprise-Icons-'+[guid]::NewGuid().ToString('N')+'.vsdx')
    $enterprise=[pscustomobject]@{schemaVersion=1;title='Icon-first native regression';presentationProfile='enterprise';pages=@(
        [pscustomobject]@{name='Architecture';role='diagram';width=12;height=8;furniture=$false;nodes=@(
            [pscustomobject]@{id='system';sourceId='source-system';kind='container';containerStyle='boundary';boundaryType='system';label='Application system';x=6;y=4;width=11;height=7},
            [pscustomobject]@{id='environment';sourceId='source-environment';kind='container';containerStyle='boundary';boundaryType='environment';label='Service environment';details='Full environment ownership detail retained outside the visible heading.';fill='RGB(242,248,254)';color='RGB(0,120,212)';x=7.5;y=4;width=7;height=5.8;parent='system'},
            [pscustomobject]@{id='select';sourceId='source-select';kind='card';cardStyle='label';label="Select destination`nFull functional semantics remain native and editable.";displayLabel='Select destination';details='Generic function; no fabricated product icon.';x=2.2;y=4;width=2.5;height=0.65;parent='system'}
        );edges=@()}
    )}
    for ($i=0; $i -lt 5; $i++) {
        $letter=[string][char](65+$i)
        $enterprise.pages[0].nodes += [pscustomobject]@{
            id="service-$letter";sourceId="source-$letter";kind='card';cardStyle='icon';iconRef=$svgEntry.id;iconSize=0.72
            label=((Get-Value $svgEntry 'name' $svgEntry.id)+" instance $letter`nCanonical source narrative remains intact without filling the diagram with large text cards.")
            displayLabel="Instance $letter";details="Native notes for instance $letter";fontSize=10
            x=(5.3+($i%3)*2.2);y=$(if ($i -lt 3) {5.3} else {2.8});width=1.8;height=1.5;parent='environment'
        }
        $enterprise.pages[0].edges += [pscustomobject]@{id="flow-$letter";source=$(if ($i -eq 0) {'select'} else {"service-$([char](64+$i))"});target="service-$letter";kind='logical';label='Request';direction='forward'}
    }
    $enterprise | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $enterprisePath -Encoding UTF8
    Assert (Run-Helper @{Action='Validate';ModelPath=$enterprisePath}).valid 'Enterprise style and model validate before creation.'
    [void](Run-Helper @{Action='New';LegacyModel=$true;ModelPath=$enterprisePath;DocumentPath=$drawing})
    $document=Get-OpenDocument $drawing; $page=$document.Pages.Item(1)
    $snapshot=Run-Helper @{Action='Inspect';DocumentPath=$drawing}; $pageId=$snapshot.pages[0].pageId
    $otherProfile=$enterprise | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $otherProfile.presentationProfile='legacy'
    $otherPath=Join-Path $OutputDirectory 'profile-change.json'
    $otherProfile | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $otherPath -Encoding UTF8
    Assert-Throws { Run-Helper @{Action='Merge';ModelPath=$otherPath;DocumentPath=$drawing} } '*Merge cannot switch presentation profile*'
    Assert ((Run-Helper @{Action='Inspect';DocumentPath=$drawing}).presentationProfile -eq 'enterprise') 'Rejected Merge does not relabel an existing presentation profile.'
    $node=$page.Shapes.ItemU('av-service-A'); $caption=Get-RoleShape $node 'caption'; $anchor=Get-RoleShape $node 'anchor'; $glyph=Get-RoleShape $node 'icon'
    Assert ($node.Shapes.Count -eq 3 -and $node.CellsU('HideText').ResultIU -eq 1) 'Icon node contains anchor, native caption, and original artwork with hidden canonical text.'
    foreach ($part in @($anchor,$caption)) { Assert ($part.CellsU('LinePattern').ResultIU -eq 0 -and $part.CellsU('FillPattern').ResultIU -eq 0) 'Anchor/caption have no visible card border or background.' }
    $probe=$page.Import($svgPath); $nativeAspect=$probe.CellsU('Width').ResultIU/$probe.CellsU('Height').ResultIU; [void]$probe.Delete()
    Assert ([Math]::Abs($glyph.CellsU('Width').ResultIU/$glyph.CellsU('Height').ResultIU-$nativeAspect) -lt 0.001) 'Icon aspect matches original native SVG import.'
    $iconLink=$page.Shapes.ItemU('av-flow-B')
    $expectedX=$node.CellsU('PinX').ResultIU-$node.CellsU('Width').ResultIU/2+$glyph.CellsU('PinX').ResultIU+$glyph.CellsU('Width').ResultIU/2
    Assert ([Math]::Abs($iconLink.CellsU('BeginX').ResultIU-$expectedX) -lt 0.001) 'Horizontal links terminate at visible glyph bounds while remaining attached to the full semantic node.'
    Assert ($iconLink.CellsU('ShapeRouteStyle').ResultIU -eq 1) 'Icon interactions use the v1.6 right-angle route policy.'
    Assert ((Get-FileHash -LiteralPath $svgPath -Algorithm SHA256).Hash -ieq $svgEntry.sha256) 'Source SVG bytes remain unchanged.'
    $caption.NameU='caption-renamed-manually'; [void]$document.Save()
    foreach ($edge in @($snapshot.pages[0].shapes | Where-Object {$_.kind -eq 'edge'})) { Assert ($edge.connections -ge 2) 'Enterprise connectors glue to full native semantic nodes.' }
    foreach ($member in @($snapshot.pages[0].shapes | Where-Object {$_.parent})) {
        $boundary=@($snapshot.pages[0].shapes | Where-Object {$_.id -eq $member.parent})[0]
        Assert ($boundary.memberIds -contains $member.visioId) 'Boundary membership remains native.'
    }
    $changesFile=Join-Path $OutputDirectory 'enterprise-changes.json'
    function Enterprise-Update([string]$Id,$Set) {
        @{schemaVersion=1;updates=@(@{pageId=$pageId;shapeId=$Id;set=$Set})} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $changesFile -Encoding UTF8
        [void](Run-Helper @{Action='Update';DocumentPath=$drawing;ChangesPath=$changesFile})
    }
    $canonical=[string]$node.Characters.Text
    Enterprise-Update 'service-A' @{displayLabel="Visible`ninstance A";details='Updated native-only notes'}
    Assert ($node.Characters.Text -ceq $canonical -and $caption.Text -eq "Visible`ninstance A") 'Caption-only update preserves all canonical text despite manual caption renaming.'
    Enterprise-Update 'service-A' @{label="Service A`nUpdated full source semantics."}
    Assert ($caption.Text -eq 'Service A' -and (Read-Property $node 'FullLabel') -eq $node.Characters.Text) 'Label update resets caption to its first line and refreshes canonical metadata.'
    Enterprise-Update 'service-A' @{label="Canonical A`nFull retained narrative.";displayLabel='Explicit caption'}
    Assert ($caption.Text -eq 'Explicit caption') 'Explicit displayLabel wins over first-line default in a combined update.'
    Enterprise-Update 'service-A' @{x=5.4}
    Assert ([Math]::Abs($glyph.CellsU('Width').ResultIU/$glyph.CellsU('Height').ResultIU-$nativeAspect) -lt 0.001) 'Full-shape moves preserve icon aspect.'
    Assert-Throws { Enterprise-Update 'service-A' @{width=2.0} } '*Width/height updates on icon groups are blocked*'
    [void](Run-Helper @{Action='Rename';DocumentPath=$drawing;ComponentId='service-A';Label="Renamed A`nCanonical rename detail."})
    Assert ($caption.Text -eq 'Renamed A') 'Rename resets visible caption instead of silently editing only hidden semantics.'
    Enterprise-Update 'select' @{displayLabel='Select target'}
    Assert ($page.Shapes.ItemU('av-select').Characters.Text -like "Select destination*Full functional*") 'Borderless label card retains source semantics during caption-only updates.'
    $location=$document.FullName; [void]$document.Close(); $document=$app.Documents.Open($location)
    $exportPath=Join-Path $OutputDirectory 'enterprise-roundtrip.json'
    [void](Run-Helper @{Action='ExportModel';DocumentPath=$drawing;ModelPath=$exportPath})
    $exported=Get-Content -LiteralPath $exportPath -Raw | ConvertFrom-Json
    Assert ($exported.presentationProfile -eq 'enterprise' -and $exported.pages[0].role -eq 'diagram') 'Profile and page role survive native save/reopen/export.'
    $exportNode=@($exported.pages[0].nodes | Where-Object {$_.id -eq 'service-A'})[0]
    Assert ($exportNode.displayLabel -eq 'Renamed A' -and $exportNode.details -eq 'Updated native-only notes' -and $exportNode.sourceId -eq 'source-A' -and $exportNode.iconSize -eq 0.72) 'Export retains actual caption, notes, source identity, and iconSize.'
    $clonePath=Join-Path $OutputDirectory 'Enterprise-Clone.vsdx'
    [void](Run-Helper @{Action='New';LegacyModel=$true;DocumentPath=$clonePath;ModelPath=$exportPath})
    $clone=Run-Helper @{Action='Inspect';DocumentPath=$clonePath}
    $cloneNode=@($clone.pages[0].shapes | Where-Object {$_.id -eq 'service-A'})[0]
    Assert ($cloneNode.label -eq $exportNode.label -and $cloneNode.displayLabel -eq $exportNode.displayLabel -and $cloneNode.details -eq $exportNode.details -and $cloneNode.sourceId -eq 'source-A' -and $clone.presentationProfile -eq 'enterprise') 'New clone preserves canonical and visual semantics.'
    Assert (@($clone.pages[0].shapes | Where-Object {$_.containerStyle -eq 'boundary' -and $_.boundaryType}).Count -eq 2) 'Boundary styles and semantic boundary types survive cloning.'
    $environment=@($clone.pages[0].shapes | Where-Object {$_.id -eq 'environment'})[0]
    Assert ($environment.boundaryFill -eq 'RGB(242,248,254)' -and $environment.boundaryColor -eq 'RGB(0,120,212)' -and
        $environment.details -eq 'Full environment ownership detail retained outside the visible heading.') 'Boundary tint, color, and hidden scope detail survive native cloning.'
    foreach ($path in $beforeDocuments.Keys) {
        $open=@(for ($i=1; $i -le $app.Documents.Count; $i++) { if ($app.Documents.Item($i).FullName -eq $path) { $app.Documents.Item($i) } })
        Assert ($open.Count -eq 1 -and $open[0].Saved -eq $beforeDocuments[$path]) 'Enterprise regression preserves unrelated open documents and saved state.'
    }
    [void](Run-Helper @{Action='Export';DocumentPath=$drawing;OutputDirectory=(Join-Path $OutputDirectory 'Enterprise-Exports')})
    $checks.Add('Native enterprise icons, invisible anchors, tagged captions, glue, aspect, CRUD, save/reopen/export/clone, boundaries, source IDs, and profile are verified.')
    [pscustomobject]@{passed=$checks.Count;checks=$checks;drawing=$drawing;clone=$clonePath;model=$exportPath} | ConvertTo-Json -Depth 8
    return
}
if ($CrudOnly) {
    $drawing=Join-Path $OutputDirectory ('Enterprise-AI-CRUD-'+[guid]::NewGuid().ToString('N')+'.vsdx')
    $changesFile=Join-Path $OutputDirectory 'changes.json'
    function Write-Changes($Value) { $Value | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $changesFile -Encoding UTF8 }
    function Crud-Snapshot { return Run-Helper @{Action='Inspect';DocumentPath=$drawing} }
    $preflight=Run-Helper @{Action='Check';ModelPath=$crudModelPath}
    Assert ($preflight.ready -and $preflight.requiredOnly -and $preflight.azureMastersVerified -eq 0 -and $preflight.iconRefsVerified -contains $svgEntry.id) 'Model Check requires only the container master and selected SVG, not all Azure stencils.'
    [void](Run-Helper @{Action='New';LegacyModel=$true;ModelPath=$crudModelPath;DocumentPath=$drawing})
    $document=Get-OpenDocument $drawing
    Assert ($null -ne $document) 'Own CRUD drawing resolves by exact path.'
    $initial=Crud-Snapshot
    $pageId=$initial.pages[0].pageId
    Assert ($initial.pages.Count -eq 1 -and $initial.pages[0].shapes.Count -eq 9) 'Non-hub graph has six semantic nodes and three edges, without default furniture.'
    $svgCard=$document.Pages.Item(1).Shapes.ItemU('av-svg-catalog')
    Assert ($svgCard.Shapes.Count -eq 2) 'Page.Import SVG is grouped with a native semantic card.'
    $probe=$document.Pages.Item(1).Import($svgPath)
    $nativeAspect=$probe.CellsU('Width').ResultIU/$probe.CellsU('Height').ResultIU
    [void]$probe.Delete()
    [void]$document.Save()
    $iconChildren=@()
    for ($i=1; $i -le $svgCard.Shapes.Count; $i++) {
        $child=$svgCard.Shapes.Item($i)
        if ($child.CellsU('Width').ResultIU -le 0.551 -and $child.CellsU('Height').ResultIU -le 0.551) { $iconChildren += $child }
    }
    Assert ($iconChildren.Count -eq 1) 'Exactly one aspect-fitted SVG child exists beside the card background.'
    $iconAspect=$iconChildren[0].CellsU('Width').ResultIU/$iconChildren[0].CellsU('Height').ResultIU
    Assert ([Math]::Abs($iconAspect-$nativeAspect) -le [Math]::Max(0.001,$nativeAspect*0.001)) 'Imported SVG preserves the native import aspect ratio.'
    Assert ((Get-FileHash -LiteralPath $svgPath -Algorithm SHA256).Hash -ieq $svgEntry.sha256) 'Original catalog asset remains byte-for-byte unchanged.'
    $checks.Add("Imported audited SVG '$($svgEntry.id)' with native Page.Import, card grouping, and preserved aspect ratio.")
    $appBefore=@($initial.pages[0].shapes | Where-Object { $_.id -eq 'copilot' })[0]
    $experience=@($initial.pages[0].shapes | Where-Object { $_.id -eq 'experience' })[0]
    Assert ($experience.relationship -eq 'logical' -and $experience.requirementIds -is [array] -and
        $experience.requirementIds[0] -is [string] -and $experience.requirementIds[0] -eq 'R1' -and
        $experience.confidence -eq 0.7) 'Edge provenance and semantics persist as a flat string array in native Shape Data.'
    $checks.Add('Created non-hub AI experience graph with logical/query/ingestion semantics and provenance.')

    Write-Changes @{schemaVersion=1;updates=@(
        @{pageId=$pageId;shapeId='copilot';set=@{label='Copilot / Teams / Web';state='Confirmed'}},
        @{pageId=$pageId;shapeId='users';set=@{x=1.9}},
        @{pageId=$pageId;shapeId='query';set=@{target=@{shapeId='data'};relationship='dependency';state='Confirmed'}}
    )}
    [void](Run-Helper @{Action='Update';ChangesPath=$changesFile;DocumentPath=$drawing})
    $updated=Crud-Snapshot
    $appAfter=@($updated.pages[0].shapes | Where-Object { $_.id -eq 'copilot' })[0]
    Assert ($appAfter.label.Trim() -eq 'Copilot / Teams / Web' -and $appAfter.state -eq 'Confirmed') 'Explicit label/state changes apply.'
    Assert ($appAfter.visioId -eq $appBefore.visioId -and $appAfter.x -eq $appBefore.x -and $appAfter.width -eq $appBefore.width) 'Update preserves unrelated identity and geometry.'
    Assert ($appAfter.sourceRef -eq $appBefore.sourceRef -and $appAfter.requirementIds[0] -eq 'R2' -and $appAfter.confidence -eq 0.8 -and $appAfter.purpose -eq 'User experience') 'Update preserves unspecified provenance.'
    Assert ($appAfter.sourceId -eq 'ref-copilot' -and $appAfter.cardStyle -eq 'detail') 'Source IDs and detail layout survive updates.'
    $detail=$document.Pages.Item(1).Shapes.ItemU('av-copilot')
    Assert ($detail.CellsU('Para.HorzAlign').ResultIU -eq 0 -and $detail.CellsU('VerticalAlign').ResultIU -eq 0 -and
        $detail.CellsU('Char.Style').ResultIU -eq 1) 'Detail cards retain left/top alignment and a native bold heading after Update.'
    $query=@($updated.pages[0].shapes | Where-Object { $_.id -eq 'query' })[0]
    $data=@($updated.pages[0].shapes | Where-Object { $_.id -eq 'data' })[0]
    Assert ($query.targetVisioIds -contains $data.visioId -and $query.target -eq 'data' -and $query.direction -eq 'both' -and $query.relationship -eq 'dependency') 'Explicit endpoint edit rewires native glue without changing direction.'
    $checks.Add('Update preserves unspecified fields and shape identity while changing labels, state, geometry, and native endpoints.')

    $beforeBad=($updated.pages | ConvertTo-Json -Depth 60 -Compress)
    Write-Changes @{schemaVersion=1;updates=@(
        @{pageId=$pageId;shapeId='copilot';set=@{label='MUST NOT APPLY'}},
        @{pageId=$pageId;visioId=2147483647;set=@{label='Missing'}}
    )}
    Assert-Throws { Run-Helper @{Action='Update';ChangesPath=$changesFile;DocumentPath=$drawing} } '*resolved to 0 full shapes*'
    Assert (((Crud-Snapshot).pages | ConvertTo-Json -Depth 60 -Compress) -ceq $beforeBad) 'Bad batch leaves the entire drawing unchanged.'
    Assert $document.Saved 'Bad preflight leaves saved state unchanged.'
    Write-Changes @{schemaVersion=1;updates=@(@{pageId=$pageId;shapeId='users';set=@{x=-5}})}
    Assert-Throws { Run-Helper @{Action='Update';ChangesPath=$changesFile;DocumentPath=$drawing} } '*outside its page*'
    $checks.Add('Bad batches and out-of-page geometry fail without any earlier partial edit.')

    Write-Changes @{schemaVersion=1;targets=@(@{pageId=$pageId;shapeId='platform'})}
    $containerPreview=Run-Helper @{Action='Delete';ChangesPath=$changesFile;DocumentPath=$drawing}
    Assert ($containerPreview.missingTargets.Count -ge 4) 'Container preview includes descendants and their connected edges.'
    Assert-Throws { Run-Helper @{Action='Delete';ChangesPath=$changesFile;DocumentPath=$drawing;ApplyDelete=$true} } '*undeclared*'
    Write-Changes @{schemaVersion=1;targets=@(@{pageId=$pageId;shapeId='users'})}
    $preview=Run-Helper @{Action='Delete';ChangesPath=$changesFile;DocumentPath=$drawing}
    Assert ($preview.preview -and -not $preview.applied -and $preview.requiredTargets.Count -eq 2) 'Connected node preview names exactly node and connected edge.'
    Assert (((Crud-Snapshot).pages | ConvertTo-Json -Depth 60 -Compress) -ceq $beforeBad) 'Delete preview never changes drawing contents.'
    Assert-Throws { Run-Helper @{Action='Delete';ChangesPath=$changesFile;DocumentPath=$drawing;ApplyDelete=$true} } '*undeclared*'
    Write-Changes @{schemaVersion=1;targets=$preview.requiredTargets;approvedTargets=@($preview.requiredTargets[0]);previewToken=$preview.previewToken}
    Assert-Throws { Run-Helper @{Action='Delete';ChangesPath=$changesFile;DocumentPath=$drawing;ApplyDelete=$true} } '*match the complete deletion list*'
    Write-Changes @{schemaVersion=1;targets=$preview.requiredTargets;approvedTargets=$preview.requiredTargets;previewToken=('0'*64)}
    Assert-Throws { Run-Helper @{Action='Delete';ChangesPath=$changesFile;DocumentPath=$drawing;ApplyDelete=$true} } '*stale*'
    Write-Changes @{schemaVersion=1;targets=$preview.requiredTargets;approvedTargets=$preview.requiredTargets;previewToken=$preview.previewToken}
    $deleted=Run-Helper @{Action='Delete';ChangesPath=$changesFile;DocumentPath=$drawing;ApplyDelete=$true}
    Assert ($deleted.applied -and $deleted.deletedTargets.Count -eq 2) 'Explicitly approved complete list is deleted.'
    Assert (@((Crud-Snapshot).pages[0].shapes | Where-Object { $_.id -in @('users','experience') }).Count -eq 0) 'Only approved connected node and edge disappear.'
    $checks.Add('Delete preview is read-only; undeclared connections/descendants, incomplete approval, and stale tokens are blocked before explicit approved deletion.')

    $location=$document.FullName
    [void]$document.Close()
    $document=$app.Documents.Open($location)
    $persisted=Crud-Snapshot
    Assert ($persisted.pages[0].shapes.Count -eq 7 -and $persisted.saved) 'Approved edits and deletion persist after reopen.'
    $modelExport=Join-Path $OutputDirectory 'roundtrip.json'
    [void](Run-Helper @{Action='ExportModel';DocumentPath=$drawing;ModelPath=$modelExport})
    Assert-Throws { Run-Helper @{Action='ExportModel';DocumentPath=$drawing;ModelPath=$modelExport} } '*Refusing to overwrite*'
    Assert (Run-Helper @{Action='Validate';ModelPath=$modelExport}).valid 'Live exported model validates without COM.'
    $exported=Get-Content -LiteralPath $modelExport -Raw | ConvertFrom-Json
    $clonePath=Join-Path $OutputDirectory ('AI-Roundtrip-'+[guid]::NewGuid().ToString('N')+'.vsdx')
    [void](Run-Helper @{Action='New';LegacyModel=$true;ModelPath=$modelExport;DocumentPath=$clonePath})
    $clone=Run-Helper @{Action='Inspect';DocumentPath=$clonePath}
    Assert ($clone.pages[0].shapes.Count -eq $persisted.pages[0].shapes.Count) 'Clone preserves graph size.'
    foreach ($shape in $persisted.pages[0].shapes) {
        $copy=@($clone.pages[0].shapes | Where-Object { $_.id -eq $shape.id })[0]
        Assert ($copy.label -eq $shape.label -and $copy.state -eq $shape.state -and $copy.sourceRef -eq $shape.sourceRef -and $copy.confidence -eq $shape.confidence) 'Clone preserves actual labels, state, source references, and numeric/qualitative confidence.'
        Assert ($copy.sourceId -eq $shape.sourceId -and $copy.cardStyle -eq $shape.cardStyle) 'Source IDs and card layout persist through native model roundtrip.'
        if ($shape.kind -eq 'edge') {
            Assert ($copy.source -eq $shape.source -and $copy.target -eq $shape.target -and $copy.relationship -eq $shape.relationship -and $copy.direction -eq $shape.direction -and $copy.dashed -eq $shape.dashed) 'Clone preserves actual graph and edge semantics.'
        } else {
            Assert ([Math]::Abs($copy.x-$shape.x) -lt 0.001 -and [Math]::Abs($copy.y-$shape.y) -lt 0.001) 'Clone preserves live node positions.'
            if ($shape.parent) { Assert ($copy.parent -eq $shape.parent) 'Clone preserves native container membership.' }
            if ($shape.iconRef) { Assert ($copy.iconRef -eq $shape.iconRef -and $copy.internalShapeCount -eq 2) 'SVG iconRef survives export and is imported into the clone.' }
        }
    }
    $checks.Add('Saved/reopened semantic export validates and creates a new editable graph retaining labels, live geometry, state, provenance, membership, and edge semantics.')

    $page=$document.Pages.Item(1)
    $native=$page.DrawRectangle(0.5,0.6,2.5,1.4); $native.Text='Native operations'
    $primitive=$page.DrawRectangle(3,0.6,3.8,1.4); $primitive.Text='Group child'
    $primitive2=$page.DrawRectangle(4,0.6,4.8,1.4)
    $selection=$page.CreateSelection(0,0,$null)
    [void]$selection.Select($primitive,2); [void]$selection.Select($primitive2,2)
    $group=$selection.Group()
    $loose=$page.Drop($app.ConnectorToolDataObject,0,0)
    $loose.CellsU('BeginX').ResultIU=1.0; $loose.CellsU('BeginY').ResultIU=0.3
    $loose.CellsU('EndX').ResultIU=3.0; $loose.CellsU('EndY').ResultIU=0.3
    [void]$document.Save()
    $nativeSnapshot=Crud-Snapshot
    $nativeRecord=@($nativeSnapshot.pages[0].shapes | Where-Object { $_.visioId -eq $native.ID })[0]
    $groupRecord=@($nativeSnapshot.pages[0].shapes | Where-Object { $_.visioId -eq $group.ID })[0]
    Assert (-not $nativeRecord.id -and $nativeRecord.label.Trim() -eq 'Native operations') 'Inspect reports untagged full-shape native IDs.'
    Assert ($groupRecord.children.Count -eq 2) 'Inspect nests native group primitives instead of counting separate architectural nodes.'
    Write-Changes @{schemaVersion=1;updates=@(@{pageId=$pageId;visioId=$native.ID;set=@{label='Native editable card';state='Observed'}})}
    [void](Run-Helper @{Action='Update';ChangesPath=$changesFile;DocumentPath=$drawing})
    Assert ($native.Text.Trim() -eq 'Native editable card') 'Native ID update works without AvId.'
    $nativeModel=Run-Helper @{Action='ExportModel';DocumentPath=$drawing}
    $nativeEdge=@($nativeModel.pages[0].edges | Where-Object { $_.id -eq "native-p$pageId-s$($loose.ID)" })[0]
    Assert ($null -eq $nativeEdge.source -and $null -eq $nativeEdge.target) 'Unglued native endpoints remain explicitly unglued.'
    Assert (@($nativeModel.warnings | Where-Object { $_ -like '*generic editable semantic*' }).Count -ge 2) 'Native export warns that generic cards do not reproduce artwork.'
    $nativeModelPath=Join-Path $OutputDirectory 'native-roundtrip.json'
    $nativeModel | ConvertTo-Json -Depth 60 | Set-Content -LiteralPath $nativeModelPath -Encoding UTF8
    Assert (Run-Helper @{Action='Validate';ModelPath=$nativeModelPath}).valid 'Arbitrary native drawing produces reusable schema with explicit uncertainty.'
    $checks.Add('Native ID updates, recursive group inspection, and uncertainty-preserving generic export work without claiming visual fidelity.')

    foreach ($path in $beforeDocuments.Keys) {
        $matches=@()
        for ($i=1; $i -le $app.Documents.Count; $i++) {
            if ($app.Documents.Item($i).FullName -eq $path) { $matches += $app.Documents.Item($i) }
        }
        Assert ($matches.Count -eq 1 -and $matches[0].Saved -eq $beforeDocuments[$path]) 'CRUD suite preserves every unrelated open document and saved/unsaved state.'
    }
    $checks.Add('All preexisting user drawings remain open with their original saved/unsaved state.')
    [pscustomobject]@{passed=$checks.Count;checks=$checks;drawing=$drawing;clone=$clonePath;model=$modelExport} | ConvertTo-Json -Depth 8
    return
}
$preflight = Run-Helper @{Action='Check'}
Assert ($preflight.ready -and $preflight.azureMastersVerified -eq 13) 'Installed Azure masters pass preflight.'
Assert ($app.Documents.Count -eq $beforeDocuments.Count) 'Preflight leaves the open document count unchanged.'
for ($i=1; $i -le $app.Documents.Count; $i++) {
    $open = $app.Documents.Item($i)
    Assert ($beforeDocuments.ContainsKey($open.FullName) -and $beforeDocuments[$open.FullName] -eq $open.Saved) 'Preflight preserves existing document state.'
}
$checks.Add('Prerequisite check verifies native masters without changing existing drawings or leaving stencils open.')

function Find-Drawing {
    $found = Get-OpenDocument $drawing
    Assert ($null -ne $found) 'Test drawing must resolve by its full path.'
    return $found
}
function Snapshot {
    return Run-Helper @{Action='Inspect'; DocumentPath=$drawing}
}
function Assert-Graph($Snapshot, $Model) {
    Assert ($Snapshot.pages.Count -eq 5) 'Five pages must persist.'
    foreach ($spec in $Model.pages) {
        $page = @($Snapshot.pages | Where-Object { $_.name -eq $spec.name })[0]
        $managed = @($page.shapes | Where-Object { $_.id })
        Assert ($managed.Count -eq $spec.nodes.Count+$spec.edges.Count) "Managed shape count on $($spec.name)"
        Assert (@($managed.id | Select-Object -Unique).Count -eq $managed.Count) "Unique ids on $($spec.name)"
        foreach ($edge in @($managed | Where-Object { $_.kind -eq 'edge' })) {
            Assert ($edge.connections -ge 2) "Connector $($edge.id) remains glued."
        }
        foreach ($node in @($managed | Where-Object { $_.parent })) {
            $container = @($managed | Where-Object { $_.id -eq $node.parent })[0]
            Assert ($container.memberIds -contains $node.visioId) "Native membership: $($node.id)"
        }
    }
}

[void](Run-Helper @{Action='New'; LegacyModel=$true; ModelPath=$baseline; DocumentPath=$drawing})
$model2 = Get-Content -LiteralPath $baseline -Raw | ConvertFrom-Json
Assert-Graph (Snapshot) $model2
$checks.Add('Created five pages with native containers, metadata, layers and glued connectors.')
$document = Find-Drawing
Assert-Throws { Run-Helper @{Action='Export'; OutputDirectory='relative-export'; DocumentPath=$drawing} } '*fully qualified Windows path*'
$page = $document.Pages.Item('03 Network topology')
$moved = $page.Shapes.ItemU('av-spoke-01-vnet')
$connector = $page.Shapes.ItemU('av-peer-spoke-01')
$beforeY = $connector.CellsU('EndY').ResultIU
$moved.CellsU('PinY').ResultIU = $moved.CellsU('PinY').ResultIU - 0.10
$manualY = $moved.CellsU('PinY').ResultIU
$manualId = $moved.ID
Assert ([Math]::Abs($connector.CellsU('EndY').ResultIU-$beforeY) -gt 0.05) 'Connector follows the moved native shape.'
$manualNote = "Platform-owned hub services.`nReview regional firewall capacity."
$page.Shapes.ItemU('av-hub-note').Text = $manualNote
[void]$document.Save()
$checks.Add('Connector endpoints follow a manual shape move.')

[void](& (Join-Path $PSScriptRoot 'New-ReferenceModel.ps1') -OutputPath $expanded -SpokeCount 3)
[void](Run-Helper @{Action='Merge'; ModelPath=$expanded; DocumentPath=$drawing})
$model3 = Get-Content -LiteralPath $expanded -Raw | ConvertFrom-Json
Assert-Graph (Snapshot) $model3
Assert ($page.Shapes.ItemU('av-spoke-01-vnet').ID -eq $manualId) 'Merge preserves shape identity.'
Assert ([Math]::Abs($page.Shapes.ItemU('av-spoke-01-vnet').CellsU('PinY').ResultIU-$manualY) -lt 0.001) 'Merge preserves manual position.'
Assert ($page.Shapes.ItemU('av-hub-note').Text.Trim() -eq $manualNote) 'Merge preserves manual note.'
$checks.Add('Added third spoke across overview, governance and topology while preserving manual work.')

$beforeCounts = @((Snapshot).pages | ForEach-Object { $_.shapes.Count })
[void](Run-Helper @{Action='Merge'; ModelPath=$expanded; DocumentPath=$drawing})
$afterCounts = @((Snapshot).pages | ForEach-Object { $_.shapes.Count })
Assert (($beforeCounts -join ',') -eq ($afterCounts -join ',')) 'Repeated merge is idempotent.'
$checks.Add('Repeated merge creates no duplicate shapes or connectors.')

$renamed = "Spoke 3 virtual network`nExpansion example"
[void](Run-Helper @{Action='Rename'; ComponentId='spoke-03-vnet'; Label=$renamed; DocumentPath=$drawing})
$views = @((Snapshot).pages.shapes | Where-Object { $_.component -eq 'spoke-03-vnet' })
Assert ($views.Count -eq 2) 'Third spoke appears in two named component views.'
foreach ($view in $views) { Assert ($view.label.Trim() -eq $renamed) 'Cross-page rename is consistent.' }
$checks.Add('Renamed the shared component consistently across views.')

$flow = $document.Pages.Item('04 Traffic and DNS').Shapes.ItemU('av-flow-egress-to-firewall')
$originalWeight = $flow.CellsU('LineWeight').ResultIU
[void](Run-Helper @{Action='Highlight'; EdgeIds=@('flow-egress-to-firewall','flow-firewall-to-internet'); DocumentPath=$drawing})
Assert ($flow.CellsU('LineWeight').ResultIU -gt $originalWeight) 'Highlight thickens selected connectors.'
[void](Run-Helper @{Action='Highlight'; EdgeIds=@('flow-egress-to-firewall','flow-firewall-to-internet'); ClearHighlight=$true; DocumentPath=$drawing})
Assert ([Math]::Abs($flow.CellsU('LineWeight').ResultIU-$originalWeight) -lt 0.0001) 'Clear highlight restores original styling.'
$checks.Add('Highlight and clear preserve original connector styling.')

$blocked = $false
try { [void](Run-Helper @{Action='Highlight'; EdgeIds=@('flow-egress-to-firewall','missing-edge'); DocumentPath=$drawing}) }
catch [System.Management.Automation.RuntimeException] {
    if ($_.Exception.Message -notlike "*Connector 'missing-edge' not found*") { throw }
    $blocked = $true
}
Assert $blocked 'Missing edge must fail.'
Assert ([Math]::Abs($flow.CellsU('LineWeight').ResultIU-$originalWeight) -lt 0.0001) 'Failed edit rolls back the earlier mutation.'
if (-not $document.Saved) { [void]$document.Save() }
$checks.Add('A partially applied invalid request rolls back rather than leaving changes.')

$blocked = $false
try { [void](Run-Helper @{Action='New'; LegacyModel=$true; ModelPath=$baseline; DocumentPath=$drawing}) }
catch [System.Management.Automation.RuntimeException] {
    if ($_.Exception.Message -notlike '*Refusing to overwrite*') { throw }
    $blocked = $true
}
Assert $blocked 'Existing files must not be overwritten.'
$checks.Add('New refuses an existing destination.')

$template = Join-Path $OutputDirectory 'Editing-Example.vstx'
[void](Run-Helper @{Action='Template'; TemplatePath=$template; DocumentPath=$drawing})
$fromTemplate = $app.Documents.Add($template)
try { Assert ($fromTemplate.Pages.Count -eq 5) 'Template creates a five-page drawing.' }
finally { $fromTemplate.Saved=$true; [void]$fromTemplate.Close() }
$checks.Add('Saved template opens as a new editable five-page drawing.')

$cloudPath = $document.FullName
[void]$document.Close()
$document = $app.Documents.Open($cloudPath)
Assert-Graph (Snapshot) $model3
Assert ($document.Pages.Item('03 Network topology').Shapes.ItemU('av-hub-note').Text.Trim() -eq $manualNote) 'Manual edits persist after reopening.'
$checks.Add('Saved graph and manual edits persist after closing and reopening.')
[void](Run-Helper @{Action='Export'; OutputDirectory=(Join-Path $OutputDirectory 'Exports'); DocumentPath=$drawing})
$checks.Add('Created PDF and five page PNG exports.')
for ($i=1; $i -le $app.Documents.Count; $i++) {
    $open = $app.Documents.Item($i)
    if ($beforeDocuments.ContainsKey($open.FullName)) {
        Assert ($beforeDocuments[$open.FullName] -eq $open.Saved) 'Integration preserves unrelated saved/unsaved state.'
    }
}
foreach ($path in $beforeDocuments.Keys) {
    $stillOpen = $false
    for ($i=1; $i -le $app.Documents.Count; $i++) {
        if ($app.Documents.Item($i).FullName -eq $path) { $stillOpen = $true; break }
    }
    Assert $stillOpen 'Integration does not close unrelated documents.'
}
$checks.Add('Existing documents remain open with their original saved or unsaved state.')
[pscustomobject]@{ passed=$checks.Count; checks=$checks; drawing=$drawing; template=$template } | ConvertTo-Json -Depth 4

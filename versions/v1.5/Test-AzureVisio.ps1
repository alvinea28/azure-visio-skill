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
    'Confirm-EdgeSemantics','Confirm-TargetRef','Confirm-Changes','Resolve-Target','Get-UpdatePlan','Confirm-Unlocked',
    'Read-Property','Test-Container','Get-IconCatalog','Resolve-IconRef','Get-TextHash',
    'Get-CardStyle','Get-DisplayLabel','Get-CaptionLayout','Get-RoleShape',
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
    Assert (Run-Helper @{Action='Validate';ModelPath=$stylePath}).valid 'Supporting notes pages are exempt while a primary diagram remains mandatory.'
    $inferred=$styleOriginal | ConvertFrom-Json
    foreach ($node in @($inferred.pages[0].nodes | Where-Object {$_.kind -eq 'card' -and (Get-Value $_ 'icon' '')})) {
        $node.PSObject.Properties.Remove('cardStyle'); $node.PSObject.Properties.Remove('displayLabel')
    }
    $inferred | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $stylePath -Encoding UTF8
    $inferredReport=& $styleHelper -ModelPath $stylePath -ReportOnly | ConvertFrom-Json
    Assert $inferredReport.valid 'Enterprise gate must infer icon style from icon/iconRef under enterprise profile, matching the renderer, and use the first canonical line as its default caption.'
    Assert (Run-Helper @{Action='Validate';ModelPath=$stylePath}).valid 'Controller and style helper agree on inferred icon styling.'
    $checks.Add('Enterprise style: inferred/default captions, hidden semantics, notes-page exemption, and 14 negative ratio/text/boundary/profile/word-budget cases pass without COM.')

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
    $root.Text="Canonical service`nOriginal full description"
    $root.CellMap.Height.ResultIU=1.4
    $root.CellMap['HideText']=[pscustomobject]@{ResultIU=1}
    $caption.Text='Visible service'; $caption.NameU='Arbitrary-manual-name'
    Set-MockProperty $caption 'AvRole' 'caption'; Set-MockProperty $glyph 'AvRole' 'icon'
    $glyph.CellMap.Width.ResultIU=0.72; $glyph.CellMap.Height.ResultIU=0.72
    $root.Shapes=New-MockCollection @($caption,$glyph)
    $pageMock.Shapes=New-MockCollection @($root)
    $record=Inspect-Shape $root
    Assert ($record.label -ceq $root.Text -and $record.displayLabel -eq $caption.Text -and $record.sourceId -eq 'source-icon') 'Inspect reads actual canonical and actual tagged caption text independently of NameU.'
    $roundtrip=Export-Model $documentMock
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
    [void](Run-Helper @{Action='New';ModelPath=$enterprisePath;DocumentPath=$drawing})
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
    Assert ($iconLink.CellsU('ShapeRouteStyle').ResultIU -eq 2) 'Aligned icon interactions use straight routes, avoiding invisible-anchor detours.'
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
    $canonical=[string]$node.Text
    Enterprise-Update 'service-A' @{displayLabel="Visible`ninstance A";details='Updated native-only notes'}
    Assert ($node.Text -ceq $canonical -and $caption.Text -eq "Visible`ninstance A") 'Caption-only update preserves all canonical text despite manual caption renaming.'
    Enterprise-Update 'service-A' @{label="Service A`nUpdated full source semantics."}
    Assert ($caption.Text -eq 'Service A' -and (Read-Property $node 'FullLabel') -eq $node.Text) 'Label update resets caption to its first line and refreshes canonical metadata.'
    Enterprise-Update 'service-A' @{label="Canonical A`nFull retained narrative.";displayLabel='Explicit caption'}
    Assert ($caption.Text -eq 'Explicit caption') 'Explicit displayLabel wins over first-line default in a combined update.'
    Enterprise-Update 'service-A' @{x=5.4}
    Assert ([Math]::Abs($glyph.CellsU('Width').ResultIU/$glyph.CellsU('Height').ResultIU-$nativeAspect) -lt 0.001) 'Full-shape moves preserve icon aspect.'
    Assert-Throws { Enterprise-Update 'service-A' @{width=2.0} } '*Width/height updates on icon groups are blocked*'
    [void](Run-Helper @{Action='Rename';DocumentPath=$drawing;ComponentId='service-A';Label="Renamed A`nCanonical rename detail."})
    Assert ($caption.Text -eq 'Renamed A') 'Rename resets visible caption instead of silently editing only hidden semantics.'
    Enterprise-Update 'select' @{displayLabel='Select target'}
    Assert ($page.Shapes.ItemU('av-select').Text -like "Select destination*Full functional*") 'Borderless label card retains source semantics during caption-only updates.'
    $location=$document.FullName; [void]$document.Close(); $document=$app.Documents.Open($location)
    $exportPath=Join-Path $OutputDirectory 'enterprise-roundtrip.json'
    [void](Run-Helper @{Action='ExportModel';DocumentPath=$drawing;ModelPath=$exportPath})
    $exported=Get-Content -LiteralPath $exportPath -Raw | ConvertFrom-Json
    Assert ($exported.presentationProfile -eq 'enterprise' -and $exported.pages[0].role -eq 'diagram') 'Profile and page role survive native save/reopen/export.'
    $exportNode=@($exported.pages[0].nodes | Where-Object {$_.id -eq 'service-A'})[0]
    Assert ($exportNode.displayLabel -eq 'Renamed A' -and $exportNode.details -eq 'Updated native-only notes' -and $exportNode.sourceId -eq 'source-A' -and $exportNode.iconSize -eq 0.72) 'Export retains actual caption, notes, source identity, and iconSize.'
    $clonePath=Join-Path $OutputDirectory 'Enterprise-Clone.vsdx'
    [void](Run-Helper @{Action='New';DocumentPath=$clonePath;ModelPath=$exportPath})
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
    [void](Run-Helper @{Action='New';ModelPath=$crudModelPath;DocumentPath=$drawing})
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
    [void](Run-Helper @{Action='New';ModelPath=$modelExport;DocumentPath=$clonePath})
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

[void](Run-Helper @{Action='New'; ModelPath=$baseline; DocumentPath=$drawing})
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
try { [void](Run-Helper @{Action='New'; ModelPath=$baseline; DocumentPath=$drawing}) }
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

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Check', 'Catalog', 'Validate', 'New', 'Merge', 'Inspect', 'ExportModel', 'Update', 'Delete', 'Rename', 'Highlight', 'Export', 'Template')]
    [string]$Action,
    [string]$DocumentPath,
    [string]$ModelPath,
    [string]$ReferencePath,
    [string]$ChangesPath,
    [string]$OutputDirectory,
    [string]$TemplatePath,
    [string]$ComponentId,
    [string]$Label,
    [string[]]$EdgeIds,
    [switch]$ClearHighlight,
    [switch]$OverwriteExports,
    [switch]$LaunchVisio = $true,
    [switch]$NoLaunchVisio,
    [switch]$ApplyDelete,
    [switch]$LegacyModel,
    [string]$StencilDirectory,
    [string]$IconDirectory,
    [string]$EnvironmentPath
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion -lt [version]'5.1') {
    throw 'Use Windows PowerShell 5.1 (powershell.exe) in the interactive Windows desktop session, not pwsh or a cloud shell.'
}
$script:app = $null
$script:stencils = @{}
$script:stencilPaths = @{}
$script:syncRoots = @()
$script:ownedStencils = [System.Collections.Generic.List[object]]::new()
$script:catalog = $null
$script:enterpriseStylePath = Join-Path $PSScriptRoot 'Enterprise-Style.ps1'
$script:icons = @{
    firewall = @('AZURENETWORKING_U.VSSX', 'Firewalls')
    vnet = @('AZURENETWORKING_U.VSSX', 'Virtual Networks')
    privateLink = @('AZURENETWORKING_U.VSSX', 'Private Link')
    dns = @('AZURENETWORKING_U.VSSX', 'DNS Zones')
    watcher = @('AZURENETWORKING_U.VSSX', 'Network Watcher')
    subscription = @('AZUREGENERAL_U.VSSX', 'Subscriptions')
    managementGroup = @('AZUREGENERAL_U.VSSX', 'Management Groups')
    storage = @('AZURESTORAGE_U.VSSX', 'Storage Accounts')
    monitor = @('AZUREMANAGEMENTGOVERNANCE_U.VSSX', 'Monitor')
    logs = @('AZUREMANAGEMENTGOVERNANCE_U.VSSX', 'Log Analytics Workspaces')
    policy = @('AZUREMANAGEMENTGOVERNANCE_U.VSSX', 'Policy')
    entra = @('AZUREIDENTITY_U.VSSX', 'Azure Active Directory')
    internet = @('AZUREENTERPRISE_U.VSSX', 'Internet')
}
$script:edgeStyles = @{
    peering = @('RGB(37,99,235)', 2, 0, 'Connectivity')
    traffic = @('RGB(194,99,13)', 1, 13, 'Traffic')
    dns = @('RGB(124,58,237)', 2, 0, 'DNS')
    governance = @('RGB(100,116,139)', 1, 0, 'Governance')
    telemetry = @('RGB(5,130,94)', 2, 13, 'Operations')
    logical = @('RGB(71,85,105)', 1, 13, 'Logical')
    ingestion = @('RGB(8,145,178)', 1, 13, 'Data')
    query = @('RGB(109,40,217)', 1, 13, 'Data')
    dependency = @('RGB(100,116,139)', 2, 13, 'Dependencies')
    association = @('RGB(100,116,139)', 2, 0, 'Associations')
}

function Get-Value($Object, [string]$Name, $Default = $null) {
    if ($null -eq $Object) { return $Default }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $Default }
    if ($property.Value -is [array]) { return ,$property.Value }
    return $property.Value
}

function Get-RunningVisio {
    return [Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application')
}

function Start-InstalledVisio {
    if ($null -eq [type]::GetTypeFromProgID('Visio.Application')) {
        throw 'Desktop Visio is not installed or registered. Visio for the web is not sufficient.'
    }
    $application = New-Object -ComObject Visio.Application
    $application.Visible = $true
    return $application
}

function Connect-Visio([bool]$AllowLaunch = $true) {
    try { return Get-RunningVisio }
    catch [Runtime.InteropServices.COMException] {
        if ($_.Exception.HResult -ne -2147221021) { throw }
        if (-not $AllowLaunch) { throw 'No running desktop Visio session; automatic launch was explicitly disabled.' }
        return Start-InstalledVisio
    }
}

function Test-Field($Object, [string]$Name) {
    return $null -ne $Object -and $null -ne $Object.PSObject.Properties[$Name]
}

function Confirm-Number($Value, [string]$Name, [double]$Minimum = -1000000, [double]$Maximum = 1000000) {
    if ($null -eq $Value -or $Value -is [bool] -or $Value -is [string] -or $Value -isnot [ValueType]) {
        throw "$Name must be a finite JSON number."
    }
    $number = [double]$Value
    if ([double]::IsNaN($number) -or [double]::IsInfinity($number) -or $number -lt $Minimum -or $number -gt $Maximum) {
        throw "$Name must be finite and between $Minimum and $Maximum."
    }
}

function Confirm-Identifier($Value) {
    if ($Value -isnot [string] -or $Value -notmatch '^[A-Za-z][A-Za-z0-9_-]{0,100}$') {
        throw "Invalid shape identifier: $Value"
    }
}

function Confirm-Fields($Object, [string[]]$Allowed, [string]$Context) {
    if ($null -eq $Object -or $Object -isnot [pscustomobject]) { throw "$Context must be a JSON object." }
    foreach ($property in $Object.PSObject.Properties) {
        if ($property.Name -cnotin $Allowed) { throw "Unsupported $Context property '$($property.Name)'." }
    }
}

function Confirm-Metadata($Object) {
    foreach ($name in @('label','displayLabel','details','boundaryType','containerStyle','state','purpose','sourceRef','sourceId','component','url','layer')) {
        if ((Test-Field $Object $name) -and (Get-Value $Object $name) -isnot [string]) { throw "$name must be a string." }
    }
    foreach ($name in @('requirementIds','mainNodeIds','mainEdgeIds')) {
        if (Test-Field $Object $name) {
            $ids = Get-Value $Object $name
            if ($ids -isnot [array]) { throw "$name must be an array of nonempty strings." }
            foreach ($id in $ids) {
                if ($id -isnot [string] -or [string]::IsNullOrWhiteSpace($id)) { throw "$name must contain nonempty strings." }
            }
        }
    }
    if (Test-Field $Object 'confidence') {
        if ($Object.confidence -is [string]) {
            if ($Object.confidence -cnotin @('High','Medium','Low','Unknown')) { throw 'confidence must be High, Medium, Low, Unknown, or a number between 0 and 1.' }
        } else { Confirm-Number $Object.confidence 'confidence' 0 1 }
    }
}

function Set-Metadata($Shape, $Object) {
    $fields = @{state='State'; purpose='Purpose'; sourceRef='SourceRef'; sourceId='SourceIdRef'; component='ComponentId'; url='Source'; details='Details'; boundaryType='BoundaryType'; containerStyle='ContainerStyle'}
    foreach ($name in $fields.Keys) {
        if (Test-Field $Object $name) { Set-Property $Shape $fields[$name] (Get-Value $Object $name) }
    }
    $jsonFields=@{requirementIds='RequirementIds';confidence='Confidence';mainNodeIds='MainNodeIds';mainEdgeIds='MainEdgeIds'}
    foreach ($name in $jsonFields.Keys) {
        if (Test-Field $Object $name) {
            $value = ConvertTo-Json -InputObject (Get-Value $Object $name) -Compress -Depth 8
            Set-Property $Shape $jsonFields[$name] $value
        }
    }
}

function Get-TextHash([string]$Text) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-','').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function Get-IconCatalog {
    if ($null -ne $script:catalog) { return $script:catalog }
    $path = Join-Path $IconDirectory 'catalog.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Icon catalog not found: $path" }
    $entries = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    if ($entries -isnot [array]) { throw 'Icon catalog must be a JSON array.' }
    $index = @{}
    foreach ($entry in $entries) {
        $id = Get-Value $entry 'id'
        if ($id -isnot [string] -or [string]::IsNullOrWhiteSpace($id)) { throw 'Catalog entry requires a nonempty id.' }
        if ($index.ContainsKey($id)) { throw "Ambiguous catalog id '$id'." }
        $index[$id] = $entry
    }
    $script:catalog = $index
    return $index
}

function Resolve-IconRef([string]$Id) {
    $catalog = Get-IconCatalog
    if (-not $catalog.ContainsKey($Id)) { throw "Missing iconRef '$Id' in catalog." }
    $entry = $catalog[$Id]
    if (Test-Field $entry 'usable') {
        if ($entry.usable -isnot [bool]) { throw "Catalog usable must be a boolean for '$Id'." }
        if (-not $entry.usable) { throw "Catalog iconRef '$Id' is not usable: $(Get-Value $entry 'issue' 'SVG audit failed.')" }
    }
    $relative = Get-Value $entry 'path' ''
    if ($relative -isnot [string] -or -not $relative -or $relative -match '[:\x00-\x1f]' -or $relative -match '^[\\/]') {
        throw "Catalog path for '$Id' must be a contained relative SVG path."
    }
    $segments = @($relative.Replace('/','\').Split('\'))
    if (@($segments | Where-Object { $_ -in @('', '.', '..') -or $_ -match '[. ]$' }).Count) {
        throw "Catalog path for '$Id' contains unsafe path segments."
    }
    $root = (Get-AbsolutePath $IconDirectory).TrimEnd('\')
    $path = [IO.Path]::GetFullPath((Join-Path $root ($segments -join '\')))
    if (-not $path.StartsWith($root+'\', [StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetExtension($path) -ine '.svg') {
        throw "Catalog path for '$Id' must be a contained relative SVG path; only .svg is supported."
    }
    $current = $root
    foreach ($segment in @('') + $segments) {
        if ($segment) { $current = Join-Path $current $segment }
        if (-not (Test-Path -LiteralPath $current)) { throw "Catalog SVG missing: $path" }
        if ((Get-Item -LiteralPath $current).Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw "Catalog path cannot traverse a reparse point: $current"
        }
    }
    $hash = Get-Value $entry 'sha256' ''
    if ($hash -notmatch '^[a-fA-F0-9]{64}$' -or (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ine $hash) {
        throw "SHA256 mismatch or invalid catalog hash for '$Id'."
    }
    return $path
}

function Get-AbsolutePath([string]$Path) {
    if ($Path -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+(?:\\|$))') {
        throw "Use a fully qualified Windows path, such as C:\Diagrams\example.vsdx: $Path"
    }
    return [IO.Path]::GetFullPath($Path)
}

function Quote-Formula([string]$Text) {
    return '"' + $Text.Replace('"', '""') + '"'
}

function Set-Property($Shape, [string]$Name, [string]$Value) {
    if (-not $Shape.CellExistsU("Prop.$Name", 0)) {
        [void]$Shape.AddNamedRow(243, $Name, 0)
        $Shape.CellsU("Prop.$Name.Label").FormulaU = Quote-Formula $Name
    }
    $Shape.CellsU("Prop.$Name").FormulaU = Quote-Formula $Value
}

function Read-Property($Shape, [string]$Name) {
    if ($Shape.CellExistsU("Prop.$Name", 0)) {
        return $Shape.CellsU("Prop.$Name").ResultStr(0)
    }
    return ''
}

function Read-Environment([string]$Path) {
    $config = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    if ($null -eq $config -or $null -eq $config.PSObject.Properties['syncRoots'] -or $config.syncRoots -isnot [array]) {
        throw 'Environment configuration requires a syncRoots array.'
    }
    $seen = @{}
    foreach ($root in $config.syncRoots) {
        $local = (Get-AbsolutePath (Get-Value $root 'local' '')).TrimEnd('\') + '\'
        $cloud = Get-Value $root 'cloud' ''
        $uri = $null
        if (-not [uri]::TryCreate($cloud, [UriKind]::Absolute, [ref]$uri) -or
            $uri.Scheme -ne 'https' -or $uri.Query -or $uri.Fragment -or $uri.UserInfo) {
            throw 'Each sync root requires an observed HTTPS cloud folder URL without query, fragment, or credentials.'
        }
        if ($seen.ContainsKey($local)) { throw "Duplicate local sync root: $local" }
        $seen[$local] = $true
        [pscustomobject]@{ local=$local; cloud=([uri]::UnescapeDataString($uri.AbsoluteUri).TrimEnd('/')) }
    }
}

function Get-CloudPath([string]$Path) {
    foreach ($root in @($script:syncRoots | Sort-Object { $_.local.Length } -Descending)) {
        if ($Path.StartsWith($root.local, [StringComparison]::OrdinalIgnoreCase)) {
            return $root.cloud + '/' + $Path.Substring($root.local.Length).Replace('\', '/')
        }
    }
    return ''
}

function Get-OpenDocument([string]$Path) {
    $cloudPath = Get-CloudPath $Path
    $matches = @()
    $unresolvedCloudCopy = $false
    for ($i = 1; $i -le $script:app.Documents.Count; $i++) {
        $candidate = $script:app.Documents.Item($i)
        $name = [uri]::UnescapeDataString($candidate.FullName)
        if ($candidate.FullName -ieq $Path -or ($cloudPath -and $name -ieq $cloudPath)) { $matches += $candidate }
        if ($name.StartsWith('https://', [StringComparison]::OrdinalIgnoreCase) -and
            [IO.Path]::GetFileName($name) -ieq [IO.Path]::GetFileName($Path)) { $unresolvedCloudCopy = $true }
    }
    if ($matches.Count -gt 1) { throw "More than one open document resolves to $Path." }
    if ($matches.Count -eq 1) { return $matches[0] }
    if ($unresolvedCloudCopy) {
        throw "An open cloud drawing has this filename, but its location cannot be matched safely: $Path. Supply a verified sync-root mapping using -EnvironmentPath, or choose a unique local working-copy path."
    }
    return $null
}

function Open-Stencil([string]$Path) {
    if ($script:stencils.ContainsKey($Path)) { return $script:stencils[$Path] }
    $stencil = Get-OpenDocument $Path
    if ($null -eq $stencil) {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Required Visio stencil missing: $Path" }
        $stencil = $script:app.Documents.OpenEx($Path, 66)
        $script:ownedStencils.Add($stencil)
    }
    $script:stencils[$Path] = $stencil
    return $stencil
}

function Resolve-StencilPath([string]$FileName) {
    if ($script:stencilPaths.ContainsKey($FileName)) { return $script:stencilPaths[$FileName] }
    $directories = @()
    if ($StencilDirectory) {
        $directories = @($StencilDirectory)
    } else {
        $directories += [IO.Path]::GetDirectoryName($script:app.GetBuiltInStencilFile(2, 2))
        $content = Join-Path $script:app.Path 'Visio Content'
        if (Test-Path -LiteralPath $content -PathType Container) {
            $directories += $content
            $directories += @(Get-ChildItem -LiteralPath $content -Directory | Sort-Object Name | ForEach-Object { $_.FullName })
        }
        $directories += @($script:app.StencilPaths -split ';' | Where-Object { $_ })
        $directories += $script:app.MyShapesPath
    }
    foreach ($directory in @($directories | Where-Object { $_ } | Select-Object -Unique)) {
        foreach ($name in @($FileName, ($FileName -replace '_U\.', '_M.'))) {
            $path = Join-Path $directory $name
            if (Test-Path -LiteralPath $path -PathType Leaf) {
                $script:stencilPaths[$FileName] = $path
                return $path
            }
        }
    }
    throw "Required Azure stencil '$FileName' is not installed in Visio's content/search folders. Use an authorized desktop Visio installation with Azure stencils, or supply -StencilDirectory. No substitute icon was used."
}

function Get-Master([string]$Icon) {
    if ($Icon -eq 'container') {
        $stencil = Open-Stencil ($script:app.GetBuiltInStencilFile(2, 2))
        return $stencil.Masters.ItemU('Plain')
    }
    if (-not $script:icons.ContainsKey($Icon)) { throw "Unknown icon '$Icon'. Available: $($script:icons.Keys -join ', ')" }
    $entry = $script:icons[$Icon]
    $stencil = Open-Stencil (Resolve-StencilPath $entry[0])
    return $stencil.Masters.ItemU($entry[1])
}

function Add-ToLayer($Page, $Shape, [string]$Name) {
    $layer = $null
    for ($i = 1; $i -le $Page.Layers.Count; $i++) {
        if ($Page.Layers.Item($i).Name -eq $Name) { $layer = $Page.Layers.Item($i); break }
    }
    if ($null -eq $layer) { $layer = $Page.Layers.Add($Name) }
    [void]$layer.Add($Shape, 0)
}

function Set-TextStyle($Shape, [double]$Size = 12, [bool]$Left = $false) {
    $Shape.CellsU('Char.Font').FormulaForceU = 'FONT("Segoe UI")'
    $Shape.CellsU('Char.Size').FormulaForceU = "$Size pt"
    $Shape.CellsU('Char.Color').FormulaForceU = 'RGB(30,41,59)'
    $Shape.CellsU('Para.HorzAlign').FormulaForceU = $(if ($Left) { '0' } else { '1' })
    $Shape.CellsU('VerticalAlign').FormulaForceU = '1'
}

function Set-DetailText($Shape, [double]$Size) {
    Set-TextStyle $Shape $Size $true
    $Shape.CellsU('VerticalAlign').FormulaForceU = '0'
    $Shape.CellsU('TopMargin').FormulaForceU = '0.16 in'
    $Shape.CellsU('BottomMargin').FormulaForceU = '0.12 in'
    $Shape.CellsU('Char.Style').FormulaForceU = '0'
    $text = [string]$Shape.Text
    $firstLine = $text.IndexOf("`n")
    if ($firstLine -lt 0) { $firstLine = $text.Length }
    if ($firstLine -gt 0) {
        $heading = $Shape.Characters
        $heading.Begin = 0
        $heading.End = $firstLine
        $heading.CharProps(2) = 1
        $heading.CharProps(7) = $Size + 1
    }
}

function Set-Bounds($Shape, $Node) {
    $Shape.CellsU('Width').ResultIUForce = [double]$Node.width
    $Shape.CellsU('Height').ResultIUForce = [double]$Node.height
    $Shape.CellsU('PinX').ResultIU = [double]$Node.x
    $Shape.CellsU('PinY').ResultIU = [double]$Node.y
}

function Get-CardStyle($Node, [string]$Profile = 'legacy') {
    if (Test-Field $Node 'cardStyle') { return $Node.cardStyle }
    if ($Profile -eq 'enterprise') {
        if ((Get-Value $Node 'icon' '') -or (Get-Value $Node 'iconRef' '')) { return 'icon' }
        return 'label'
    }
    return 'standard'
}

function Get-DisplayLabel($Node) {
    return Get-Value $Node 'displayLabel' (([string](Get-Value $Node 'label' '') -split '\r\n|\n|\r', 2)[0].Trim())
}

function Get-CaptionLayout($Node, [string]$Style) {
    $caption = Get-DisplayLabel $Node
    $lines = @($caption -split '\r\n|\n|\r')
    if ($lines.Count -gt 3 -or @($lines | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count) {
        throw 'displayLabel must contain 1-3 nonempty caption lines for icon/label cards.'
    }
    $font = [double](Get-Value $Node 'fontSize' 10)
    $captionHeight = $lines.Count*$font/72*1.25+0.08
    $longest = ($lines | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    if ($Node.width -lt [Math]::Max(0.65,$longest*$font/72*0.5+0.08) -or $Node.height -lt $captionHeight) {
        throw 'Caption geometry is too small; widen the node, use 1-3 shorter displayLabel lines, or reduce fontSize.'
    }
    $size = 0.0
    if ($Style -eq 'icon') {
        $available = [Math]::Min($Node.width-0.08,$Node.height-$captionHeight-0.10)
        $size = [double](Get-Value $Node 'iconSize' ([Math]::Min(0.72,$available)))
        if ($size -lt 0.35 -or $size -gt $available+0.001) {
            throw 'Icon geometry is too small for iconSize and caption; enlarge width/height or reduce iconSize (minimum 0.35 inches).'
        }
    }
    return [pscustomobject]@{caption=$caption;fontSize=$font;captionHeight=$captionHeight;iconSize=$size}
}

function Get-RoleShape($Shape, [string]$Role, [switch]$Optional) {
    $matches = @()
    for ($i=1; $i -le $Shape.Shapes.Count; $i++) {
        $child=$Shape.Shapes.Item($i)
        if ((Read-Property $child 'AvRole') -eq $Role) { $matches += $child }
    }
    if ($matches.Count -eq 1) { return $matches[0] }
    if (-not $Optional) { throw "Shape $($Shape.ID) requires exactly one native AvRole='$Role' child; restore the managed group before updating/exporting." }
    return $null
}

function Set-NodeLabel($Shape, $Spec) {
    if (-not (Test-Field $Spec 'label') -and -not (Test-Field $Spec 'displayLabel')) { return }
    $style = Read-Property $Shape 'CardStyle'
    if ($style -in @('icon','label')) {
        $caption = Get-RoleShape $Shape 'caption'
        if (Test-Field $Spec 'label') { $Shape.Characters.Text=$Spec.label; Set-Property $Shape 'FullLabel' $Spec.label }
        if ((Test-Field $Spec 'label') -or (Test-Field $Spec 'displayLabel')) {
            $caption.Text=Get-DisplayLabel $Spec
            $w=$Shape.CellsU('Width').ResultIU; $h=$Shape.CellsU('Height').ResultIU
            $layoutSpec=[pscustomobject]@{label=[string]$Shape.Characters.Text;displayLabel=[string]$caption.Text;width=$w;height=$h;fontSize=[double](Read-Property $Shape 'BaseFontSize')}
            if ($style -eq 'icon') { $layoutSpec | Add-Member -NotePropertyName iconSize -NotePropertyValue ([double](Read-Property $Shape 'IconSize')) }
            $layout=Get-CaptionLayout $layoutSpec $style
            $widthFormula=$w.ToString('R',[Globalization.CultureInfo]::InvariantCulture)+' in'
            $caption.CellsU('Width').FormulaForceU="MIN($widthFormula,TEXTWIDTH(TheText,$widthFormula-0.08 in)+0.08 in)"
            $caption.CellsU('Height').ResultIUForce=$layout.captionHeight
            $caption.CellsU('PinX').ResultIU=$w/2
            $caption.CellsU('PinY').ResultIU=$(if ($style -eq 'icon') { $layout.captionHeight/2 } else { $h/2 })
            if ($style -eq 'icon') {
                $icon=Get-RoleShape $Shape 'icon'
                $icon.CellsU('PinY').ResultIU=$h-$layout.iconSize/2-0.05
            }
        }
        $Shape.CellsU('HideText').FormulaU='1'
    } elseif (Test-Field $Spec 'label') {
        $Shape.Text=$Spec.label
        if ($style -eq 'detail') { Set-DetailText $Shape ([double](Read-Property $Shape 'BaseFontSize')) }
    }
}

function Add-CaptionNode($Page, $Node, [string]$Style) {
    $layout=Get-CaptionLayout $Node $Style
    $x=[double]$Node.x; $y=[double]$Node.y; $w=[double]$Node.width; $h=[double]$Node.height
    $anchor=$Page.DrawRectangle(($x-$w/2),($y-$h/2),($x+$w/2),($y+$h/2))
    $captionTop=$(if ($Style -eq 'icon') { $y-$h/2+$layout.captionHeight } else { $y+$h/2 })
    $caption=$Page.DrawRectangle(($x-$w/2),($y-$h/2),($x+$w/2),$captionTop)
    foreach ($part in @($anchor,$caption)) {
        $part.CellsU('LinePattern').FormulaU='0'; $part.CellsU('FillPattern').FormulaU='0'
        $part.CellsU('TextBkgnd').FormulaU='0'
        foreach ($margin in @('LeftMargin','RightMargin','TopMargin','BottomMargin')) { $part.CellsU($margin).FormulaU='0' }
    }
    Set-Property $anchor 'AvRole' 'anchor'
    Set-Property $caption 'AvRole' 'caption'
    $caption.Text=$layout.caption
    Set-TextStyle $caption $layout.fontSize
    $selection=$Page.CreateSelection(0,0,$null)
    [void]$selection.Select($anchor,2); [void]$selection.Select($caption,2)
    if ($Style -eq 'icon') {
        $iconRef=Get-Value $Node 'iconRef' ''
        if ($iconRef) { $icon=$Page.Import((Resolve-IconRef $iconRef)) }
        else { $icon=$Page.Drop((Get-Master $Node.icon),$x,$y); $icon.Text='' }
        $iw=$icon.CellsU('Width').ResultIU; $ih=$icon.CellsU('Height').ResultIU
        if ($iw -le 0 -or $ih -le 0) { throw "Imported icon has invalid dimensions on $($Node.id)." }
        $scale=$layout.iconSize/[Math]::Max($iw,$ih)
        $icon.CellsU('Width').ResultIUForce=$iw*$scale; $icon.CellsU('Height').ResultIUForce=$ih*$scale
        $icon.CellsU('PinX').ResultIU=$x
        $icon.CellsU('PinY').ResultIU=$y+$h/2-$layout.iconSize/2-0.05
        Set-Property $icon 'AvRole' 'icon'
        [void]$selection.Select($icon,2)
    }
    $shape=$selection.Group()
    $shape.CellsU('HideText').FormulaU='1'
    Set-Property $shape 'CardStyle' $Style
    Set-Property $shape 'BaseFontSize' ([string]$layout.fontSize)
    Set-Property $shape 'Details' (Get-Value $Node 'details' '')
    if ($Style -eq 'icon') {
        Set-Property $shape 'IconSize' ([string]$layout.iconSize)
        Set-Property $shape 'IconAspect' ([string]($iw/$ih))
    }
    Set-NodeLabel $shape $Node
    return $shape
}

function Add-Hyperlink($Shape, [string]$Url) {
    if ($Url) {
        $link = $Shape.AddHyperlink()
        $link.Address = $Url
        $link.Description = 'Microsoft Learn'
    }
}

function Set-ContainerStyle($Shape, [string]$Color, [string]$Weight = '1.2 pt', [string]$Fill = '') {
    $Shape.CellsU('FillPattern').FormulaForceU = $(if ($Fill) { '1' } else { '0' })
    if ($Fill) { $Shape.CellsU('FillForegnd').FormulaForceU = $Fill }
    $Shape.CellsU('LineColor').FormulaForceU = $Color
    $Shape.CellsU('LineWeight').FormulaForceU = $Weight
    for ($i = 1; $i -le $Shape.Shapes.Count; $i++) {
        Set-ContainerStyle $Shape.Shapes.Item($i) $Color $Weight $Fill
    }
}

function Add-Node($Page, $Node, [string]$Profile = 'legacy') {
    $x = [double]$Node.x; $y = [double]$Node.y
    $w = [double]$Node.width; $h = [double]$Node.height
    switch ($Node.kind) {
        'container' {
            $shape = $Page.Drop((Get-Master 'container'), $x, $y)
            $shape.ContainerProperties.ResizeAsNeeded = 0
            Set-Bounds $shape $Node
            $shape.Text = $Node.label
            $boundaryStyle = (Get-Value $Node 'containerStyle') -eq 'boundary'
            Set-ContainerStyle $shape (Get-Value $Node 'color' 'RGB(100,116,139)') $(if ($boundaryStyle) { '0.8 pt' } else { '1.2 pt' }) $(if ($boundaryStyle) { Get-Value $Node 'fill' '' } else { '' })
            $shape.CellsU('LinePattern').FormulaForceU = [string](Get-Value $Node 'linePattern' 1)
            Set-TextStyle $shape ([double](Get-Value $Node 'fontSize' 11)) $true
        }
        'card' {
            $cardStyle = Get-CardStyle $Node $Profile
            if ($cardStyle -in @('icon','label')) {
                $shape=Add-CaptionNode $Page $Node $cardStyle
                break
            }
            $box = $Page.DrawRectangle(($x-$w/2), ($y-$h/2), ($x+$w/2), ($y+$h/2))
            $box.CellsU('FillForegnd').FormulaU = Get-Value $Node 'fill' 'RGB(255,255,255)'
            $box.CellsU('LineColor').FormulaU = Get-Value $Node 'color' 'RGB(148,163,184)'
            $box.CellsU('LineWeight').FormulaU = '0.75 pt'
            $box.CellsU('LinePattern').FormulaU = [string](Get-Value $Node 'linePattern' 1)
            $box.CellsU('Rounding').FormulaU = '0.08 in'
            $iconName = Get-Value $Node 'icon' ''
            $iconRef = Get-Value $Node 'iconRef' ''
            if ($iconName -or $iconRef) {
                if ($iconRef) {
                    $icon = $Page.Import((Resolve-IconRef $iconRef))
                } else {
                    $icon = $Page.Drop((Get-Master $iconName), ($x-$w/2+0.43), $y)
                    $icon.Text = ''
                }
                $iw = $icon.CellsU('Width').ResultIU
                $ih = $icon.CellsU('Height').ResultIU
                if ($iw -le 0 -or $ih -le 0) { throw "Imported icon has invalid dimensions on $($Node.id)." }
                $scale = [Math]::Min(0.55/$iw, [Math]::Min(0.55, $h*0.65)/$ih)
                $icon.CellsU('Width').ResultIUForce = $iw*$scale
                $icon.CellsU('Height').ResultIUForce = $ih*$scale
                $icon.CellsU('PinX').ResultIU = $x-$w/2+0.43
                $icon.CellsU('PinY').ResultIU = $(if ($cardStyle -eq 'detail') { $y+$h/2-0.44 } else { $y })
                $selection = $Page.CreateSelection(0, 0, $null)
                [void]$selection.Select($box, 2)
                [void]$selection.Select($icon, 2)
                $shape = $selection.Group()
                $shape.CellsU('TxtPinX').FormulaU = 'Width/2+0.33 in'
                $shape.CellsU('TxtWidth').FormulaU = 'Width-0.85 in'
            } else {
                $shape = $box
                $shape.CellsU('TxtWidth').FormulaU = 'Width-0.16 in'
            }
            $shape.Text = $Node.label
            $fontSize = [double](Get-Value $Node 'fontSize' 12)
            if ($cardStyle -eq 'detail') { Set-DetailText $shape $fontSize }
            else { Set-TextStyle $shape $fontSize }
            Set-Property $shape 'CardStyle' $cardStyle
            Set-Property $shape 'BaseFontSize' ([string]$fontSize)
        }
        'note' {
            $shape = $Page.DrawRectangle(($x-$w/2), ($y-$h/2), ($x+$w/2), ($y+$h/2))
            $shape.Text = $Node.label
            $shape.CellsU('FillForegnd').FormulaU = Get-Value $Node 'fill' 'RGB(241,245,249)'
            $shape.CellsU('LinePattern').FormulaU = '0'
            $shape.CellsU('LeftMargin').FormulaU = '0.14 in'
            $shape.CellsU('RightMargin').FormulaU = '0.14 in'
            Set-TextStyle $shape ([double](Get-Value $Node 'fontSize' 12)) $true
        }
    }
    $shape.NameU = 'av-' + $Node.id
    Set-Property $shape 'AvId' $Node.id
    Set-Property $shape 'ComponentId' (Get-Value $Node 'component' $Node.id)
    Set-Property $shape 'Kind' $Node.kind
    Set-Property $shape 'State' (Get-Value $Node 'state' 'Illustrative')
    Set-Property $shape 'Purpose' (Get-Value $Node 'purpose' $Node.label)
    Set-Property $shape 'ParentId' (Get-Value $Node 'parent' '')
    Set-Property $shape 'Source' (Get-Value $Node 'url' '')
    Set-Property $shape 'Icon' (Get-Value $Node 'icon' '')
    Set-Property $shape 'IconRef' (Get-Value $Node 'iconRef' '')
    Set-Metadata $shape $Node
    Add-Hyperlink $shape (Get-Value $Node 'url' '')
    Add-ToLayer $Page $shape (Get-Value $Node 'layer' 'Architecture')
    return $shape
}

function Get-ServicePort($Target, [string]$Side, [double]$Position = 0.5) {
    $u=0.5; $v=0.5
    $width=$Target.CellsU('Width').ResultIU; $height=$Target.CellsU('Height').ResultIU
    $left=0.0; $bottom=0.0; $right=$width; $top=$height
    $style=Read-Property $Target 'CardStyle'
    $region='body'
    $leftFormula='0'; $bottomFormula='0'; $widthFormula='Width'; $heightFormula='Height'
    if ($style -in @('icon','label')) {
        $region=$(if ($style -eq 'icon' -and $Side -ne 'bottom') { 'glyph' } else { 'caption' })
        $part=Get-RoleShape $Target $(if ($region -eq 'glyph') { 'icon' } else { 'caption' })
        $gx=$part.CellsU('PinX').ResultIU; $gy=$part.CellsU('PinY').ResultIU
        $gw=$part.CellsU('Width').ResultIU; $gh=$part.CellsU('Height').ResultIU
        $left=$gx-$gw/2; $right=$gx+$gw/2; $bottom=$gy-$gh/2; $top=$gy+$gh/2
        if ($region -eq 'caption') {
            $captionExtent=Get-NativeCaptionBounds $part -Local
            $left=$captionExtent.left; $right=$captionExtent.right
            $bottom=$captionExtent.bottom; $top=$captionExtent.top
        }
        $widthFormula="Sheet.$($part.ID)!Width"; $heightFormula="Sheet.$($part.ID)!Height"
        $leftFormula="(Sheet.$($part.ID)!PinX-$widthFormula/2)"
        $bottomFormula="(Sheet.$($part.ID)!PinY-$heightFormula/2)"
        if ($region -eq 'caption') {
            $span=(($right-$left)/$gw).ToString('R',[Globalization.CultureInfo]::InvariantCulture)
            $offset=(($left-$gx)/$gw).ToString('R',[Globalization.CultureInfo]::InvariantCulture)
            $widthFormula="Sheet.$($part.ID)!Width*$span"
            $leftFormula="(Sheet.$($part.ID)!PinX+Sheet.$($part.ID)!Width*($offset))"
        }
    }
    $dx=0; $dy=0
    $fraction=$Position.ToString('R',[Globalization.CultureInfo]::InvariantCulture)
    $xFormula="$leftFormula+$widthFormula*$fraction"; $yFormula="$bottomFormula+$heightFormula*$fraction"
    switch ($Side) {
        'left' { $u=$left/$width; $v=($bottom+($top-$bottom)*$Position)/$height; $dx=-1; $xFormula=$leftFormula }
        'right' { $u=$right/$width; $v=($bottom+($top-$bottom)*$Position)/$height; $dx=1; $xFormula="$leftFormula+$widthFormula" }
        'top' { $u=($left+($right-$left)*$Position)/$width; $v=$top/$height; $dy=1; $yFormula="$bottomFormula+$heightFormula" }
        'bottom' { $u=($left+($right-$left)*$Position)/$width; $v=$bottom/$height; $dy=-1; $yFormula=$bottomFormula }
        default { throw "Invalid connector side '$Side'." }
    }
    return [pscustomobject]@{u=$u;v=$v;dirX=$dx;dirY=$dy;region=$region;xFormula=$xFormula;yFormula=$yFormula}
}

function Glue-End($Cell, $Target, [string]$Side, [double]$Position = 0.5) {
    $port=Get-ServicePort $Target $Side $Position
    if (-not (Read-Property $Target 'AvId')) {
        $Cell.GlueToPos($Target,$port.u,$port.v)
        return
    }
    $row='AvPort'+$Side+($Position.ToString('R',[Globalization.CultureInfo]::InvariantCulture) -replace '\.','_' -replace '-','m' -replace '\+','p')
    if (-not $Target.CellExistsU("Connections.$row.X",0)) { [void]$Target.AddNamedRow(7,$row,185) }
    $Target.CellsU("Connections.$row.X").FormulaU=$port.xFormula
    $Target.CellsU("Connections.$row.Y").FormulaU=$port.yFormula
    # Type=0 connection vectors point inward, opposite the outward route escape.
    $Target.CellsU("Connections.$row.DirX").FormulaU=[string](-$port.dirX)
    $Target.CellsU("Connections.$row.DirY").FormulaU=[string](-$port.dirY)
    $Target.CellsU("Connections.$row.Type").FormulaU='0'
    $Cell.GlueTo($Target.CellsU("Connections.$row.X"))
}

function Get-NativeBounds($Shape) {
    $points=@()
    foreach ($corner in @(@(0.0,0.0),@($Shape.CellsU('Width').ResultIU,0.0),
        @($Shape.CellsU('Width').ResultIU,$Shape.CellsU('Height').ResultIU),@(0.0,$Shape.CellsU('Height').ResultIU))) {
        [double]$x=0; [double]$y=0
        $Shape.XYToPage([double]$corner[0],[double]$corner[1],[ref]$x,[ref]$y)
        $points+=[pscustomobject]@{x=$x;y=$y}
    }
    return [pscustomobject]@{
        left=($points.x | Measure-Object -Minimum).Minimum; right=($points.x | Measure-Object -Maximum).Maximum
        bottom=($points.y | Measure-Object -Minimum).Minimum; top=($points.y | Measure-Object -Maximum).Maximum
    }
}

function Get-NativeCaptionBounds($Caption, [switch]$Local) {
    if ($Local) {
        $w=$Caption.CellsU('Width').ResultIU; $h=$Caption.CellsU('Height').ResultIU
        $x=$Caption.CellsU('PinX').ResultIU; $y=$Caption.CellsU('PinY').ResultIU
        $frame=[pscustomobject]@{left=$x-$w/2;right=$x+$w/2;bottom=$y-$h/2;top=$y+$h/2}
    } else { $frame=Get-NativeBounds $Caption }
    [double]$left=0; [double]$bottom=0; [double]$right=0; [double]$top=0
    $Caption.BoundingBox($(if ($Local) { 2 } else { 8194 }),[ref]$left,[ref]$bottom,[ref]$right,[ref]$top)
    foreach ($value in @($left,$bottom,$right,$top)) { Confirm-Number $value 'Native caption text extent' }
    if ($right -le $left -or $top -le $bottom) { throw 'Native caption text extent is unavailable.' }
    # Caption line boxes include 0.04-inch breathing room; the native text extent
    # trims older portable children whose invisible width spans the whole node.
    $left=[Math]::Max([double]$frame.left,$left-0.04); $right=[Math]::Min([double]$frame.right,$right+0.04)
    if ($right -le $left) { throw 'Native caption text lies outside its caption frame.' }
    return [pscustomobject]@{left=$left;right=$right;bottom=$frame.bottom;top=$frame.top;basis='Native BoundingBox(visBBoxUprightText), caption line frame and 0.04-inch horizontal clearance.'}
}

function Get-NativeRouteGeometry($Shape) {
    $segments=@(); $paths=@(); $diagonals=0; $curves=0; $pathIndex=0
    try {
        $count=[int]$Shape.GeometryCount
        if ($count -eq 0 -or $count -gt 100) { throw 'Connector has no supported observable Geometry sections.' }
        for ($g=0; $g -lt $count; $g++) {
            $section=10+$g; $geometry='Geometry'+($g+1)
            if (($Shape.CellExistsU("$geometry.NoShow",0) -and $Shape.CellsU("$geometry.NoShow").ResultIU -ne 0) -or
                ($Shape.CellExistsU("$geometry.NoLine",0) -and $Shape.CellsU("$geometry.NoLine").ResultIU -ne 0)) { continue }
            $rows=[int]$Shape.RowCount($section)
            if ($rows -gt 20000) { throw 'Connector Geometry section exceeds the inspection limit.' }
            $points=@()
            for ($row=1; $row -lt $rows; $row++) {
                $tag=[int]$Shape.RowType($section,$row)
                if ($tag -notin @(138,139,140,144,195,236,237,238,239,240)) { throw "Unsupported Geometry row type $tag; actual stroke is not verified." }
                $localX=[double]$Shape.CellsSRC($section,$row,0).ResultIU
                $localY=[double]$Shape.CellsSRC($section,$row,1).ResultIU
                if ($tag -in @(236,237,238,239,240)) {
                    $localX*=$Shape.CellsU('Width').ResultIU; $localY*=$Shape.CellsU('Height').ResultIU
                }
                [double]$x=0; [double]$y=0
                $Shape.XYToPage($localX,$localY,[ref]$x,[ref]$y)
                Confirm-Number $x 'Native path X'; Confirm-Number $y 'Native path Y'
                if ($tag -in @(138,238)) {
                    if ($points.Count) { $paths+=[pscustomobject]@{index=$pathIndex;section=$section;points=$points} }
                    $points=@(); $pathIndex++
                } elseif (-not $points.Count) { throw 'Geometry line/curve has no preceding MoveTo; actual stroke is not verified.' }
                $points+=[pscustomobject]@{x=$x;y=$y}
                if ($tag -in @(138,238)) { continue }
                $a=$points[$points.Count-2]; $dx=$x-$a.x; $dy=$y-$a.y
                $linear=$tag -in @(139,239)
                if ($tag -eq 140 -and [Math]::Abs($Shape.CellsSRC($section,$row,2).ResultIU) -le 0.000000001) { $linear=$true }
                if ($linear -and [Math]::Abs($dx) -le 0.001 -and [Math]::Abs($dy) -le 0.001) { continue }
                $axis=$(if (-not $linear) { 'curve' } elseif ([Math]::Abs($dy) -le 0.001) { 'horizontal' } elseif ([Math]::Abs($dx) -le 0.001) { 'vertical' } else { 'diagonal' })
                if ($axis -eq 'diagonal') { $diagonals++ }
                if ($axis -eq 'curve') { $curves++ }
                $segments+=[pscustomobject]@{path=$pathIndex;section=$section;row=$row;rowType=$tag;fromX=$a.x;fromY=$a.y;toX=$x;toY=$y;axis=$axis}
            }
            if ($points.Count) { $paths+=[pscustomobject]@{index=$pathIndex;section=$section;points=$points} }
        }
        if (-not $segments.Count) { throw 'Connector has no nonzero stroke segments.' }
        return [pscustomobject]@{available=$true;coordinateSystem='page';basis='Geometry rows and XYToPage; curves are flagged, never approximated as straight lines.';tolerance=0.001;paths=$paths;segments=$segments;diagonalCount=$diagonals;curvedCount=$curves;orthogonal=($diagonals -eq 0 -and $curves -eq 0);error=''}
    } catch {
        return [pscustomobject]@{available=$false;coordinateSystem='page';basis='Geometry rows and XYToPage';tolerance=0.001;paths=@();segments=@();diagonalCount=$null;curvedCount=$null;orthogonal=$null;error=$_.Exception.Message}
    }
}

function Test-RouteCrossing($Segment, $Bounds) {
    $epsilon=0.001
    if ($Segment.axis -eq 'horizontal') {
        return $Segment.fromY -gt $Bounds.bottom+$epsilon -and $Segment.fromY -lt $Bounds.top-$epsilon -and
            [Math]::Min($Segment.fromX,$Segment.toX) -lt $Bounds.right-$epsilon -and
            [Math]::Max($Segment.fromX,$Segment.toX) -gt $Bounds.left+$epsilon
    }
    if ($Segment.axis -eq 'vertical') {
        return $Segment.fromX -gt $Bounds.left+$epsilon -and $Segment.fromX -lt $Bounds.right-$epsilon -and
            [Math]::Min($Segment.fromY,$Segment.toY) -lt $Bounds.top-$epsilon -and
            [Math]::Max($Segment.fromY,$Segment.toY) -gt $Bounds.bottom+$epsilon
    }
    return $false
}

function Find-BoundaryRoute($Start, $End, [string]$StartSide, [string]$EndSide, $Obstacles, [double]$Width, [double]$Height) {
    $clearance=0.15
    $escapes=@()
    foreach ($endpoint in @(@($Start,$StartSide),@($End,$EndSide))) {
        $x=[double]$endpoint[0].x; $y=[double]$endpoint[0].y
        switch ($endpoint[1]) { 'left' {$x-=$clearance} 'right' {$x+=$clearance} 'bottom' {$y-=$clearance} 'top' {$y+=$clearance} }
        $escapes+=[pscustomobject]@{x=$x;y=$y}
    }
    $a=$escapes[0]; $b=$escapes[1]
    $xs=@($a.x,$b.x); $ys=@($a.y,$b.y)
    foreach ($bounds in $Obstacles) {
        $xs+=@(($bounds.left-$clearance),($bounds.right+$clearance))
        $ys+=@(($bounds.bottom-$clearance),($bounds.top+$clearance))
    }
    $candidates=@()
    foreach ($x in @($xs | Sort-Object -Unique)) {
        $candidates+=,@($Start,$a,[pscustomobject]@{x=$x;y=$a.y},[pscustomobject]@{x=$x;y=$b.y},$b,$End)
    }
    foreach ($y in @($ys | Sort-Object -Unique)) {
        $candidates+=,@($Start,$a,[pscustomobject]@{x=$a.x;y=$y},[pscustomobject]@{x=$b.x;y=$y},$b,$End)
    }
    $best=$null; $bestLength=[double]::PositiveInfinity
    foreach ($candidate in $candidates) {
        $points=@(); $valid=$true; $length=0.0
        foreach ($point in $candidate) {
            if ($point.x -lt 0 -or $point.x -gt $Width -or $point.y -lt 0 -or $point.y -gt $Height) { $valid=$false; break }
            if ($points.Count -and [Math]::Abs($point.x-$points[-1].x) -lt 0.000001 -and [Math]::Abs($point.y-$points[-1].y) -lt 0.000001) { continue }
            if ($points.Count) {
                $previous=$points[-1]
                $segment=[pscustomobject]@{fromX=$previous.x;fromY=$previous.y;toX=$point.x;toY=$point.y;axis=$(if ([Math]::Abs($previous.y-$point.y) -lt 0.000001) {'horizontal'} else {'vertical'})}
                foreach ($bounds in $Obstacles) { if (Test-RouteCrossing $segment $bounds) { $valid=$false; break } }
                if (-not $valid) { break }
                if ($points.Count -ge 2) {
                    $earlier=$points[-2]
                    $ax=$previous.x-$earlier.x; $ay=$previous.y-$earlier.y
                    $bx=$point.x-$previous.x; $by=$point.y-$previous.y
                    if ($ax*$bx+$ay*$by -lt -0.000001) { $valid=$false; break }
                    if (([Math]::Abs($ax) -lt 0.000001 -and [Math]::Abs($bx) -lt 0.000001) -or
                        ([Math]::Abs($ay) -lt 0.000001 -and [Math]::Abs($by) -lt 0.000001)) { $points=@($points[0..($points.Count-2)]) }
                }
                $length+=[Math]::Abs($point.x-$previous.x)+[Math]::Abs($point.y-$previous.y)
            }
            $points+=$point
        }
        if ($valid -and $length -lt $bestLength) { $best=$points; $bestLength=$length }
    }
    if ($null -eq $best) { throw 'No clear orthogonal boundary approach exists in the available gutters; adjust the layout or the selected boundary port.' }
    return ,$best
}

function Repair-NativeBoundaryRoutes($Page, $Index) {
    $records=$null
    foreach ($edge in $Index.Values) {
        if (-not $edge.OneD -or (Read-Property $edge 'Kind') -ne 'edge') { continue }
        $source=Read-Property $edge 'SourceId'; $target=Read-Property $edge 'TargetId'
        if (-not $source -or -not $target -or -not $Index.ContainsKey($source) -or -not $Index.ContainsKey($target)) { continue }
        $boundaryIds=@($source,$target | Where-Object { (Read-Property $Index[$_] 'Kind') -eq 'container' })
        if (-not $boundaryIds.Count) { continue }
        $geometry=Get-NativeRouteGeometry $edge
        if (-not $geometry.available -or -not $geometry.orthogonal) { continue }
        $needsRepair=$false
        foreach ($pair in @(@('Source',$source,'Begin'),@('Target',$target,'End'))) {
            if ($pair[1] -notin $boundaryIds) { continue }
            $side=Read-Property $edge ($pair[0]+'Side')
            $x=$edge.CellsU($pair[2]+'X').ResultIU; $y=$edge.CellsU($pair[2]+'Y').ResultIU
            $outward=$false
            foreach ($segment in $geometry.segments) {
                $dx=0.0; $dy=0.0
                if ([Math]::Abs($segment.fromX-$x) -lt 0.001 -and [Math]::Abs($segment.fromY-$y) -lt 0.001) { $dx=$segment.toX-$x; $dy=$segment.toY-$y }
                elseif ([Math]::Abs($segment.toX-$x) -lt 0.001 -and [Math]::Abs($segment.toY-$y) -lt 0.001) { $dx=$segment.fromX-$x; $dy=$segment.fromY-$y }
                if (($side -eq 'left' -and $dx -lt -0.001) -or ($side -eq 'right' -and $dx -gt 0.001) -or
                    ($side -eq 'bottom' -and $dy -lt -0.001) -or ($side -eq 'top' -and $dy -gt 0.001)) { $outward=$true; break }
            }
            if (-not $outward) { $needsRepair=$true }
        }
        if (-not $needsRepair) { continue }
        if ($null -eq $records) { $records=@($Index.Values | Where-Object {-not $_.OneD} | ForEach-Object { Inspect-Shape $_ $false }) }
        $obstacles=@()
        foreach ($record in $records) {
            if ($record.kind -notin @('card','note') -and $record.id -notin $boundaryIds) { continue }
            foreach ($field in @('serviceBounds','captionBounds')) { $bounds=Get-Value $record $field; if ($null -ne $bounds) { $obstacles+=$bounds } }
        }
        $start=[pscustomobject]@{x=$edge.CellsU('BeginX').ResultIU;y=$edge.CellsU('BeginY').ResultIU}
        $end=[pscustomobject]@{x=$edge.CellsU('EndX').ResultIU;y=$edge.CellsU('EndY').ResultIU}
        $points=Find-BoundaryRoute $start $end (Read-Property $edge 'SourceSide') (Read-Property $edge 'TargetSide') $obstacles $Page.PageSheet.CellsU('PageWidth').ResultIU $Page.PageSheet.CellsU('PageHeight').ResultIU
        [double]$originX=0; [double]$originY=0
        $edge.XYToPage(0,0,[ref]$originX,[ref]$originY)
        if ([Math]::Abs($edge.CellsU('Angle').ResultIU) -gt 0.000001 -or
            [Math]::Abs($originX-$start.x) -gt 0.000001 -or [Math]::Abs($originY-$start.y) -gt 0.000001) { throw 'Unsupported boundary connector transform; cannot safely author a page-aligned route.' }
        # Visio ignores outward port constraints on native containers. Preserve
        # full-shape glue, with explicit orthogonal gutter geometry whose endpoint
        # coordinates remain formula-linked when a service moves.
        $edge.CellsU('ConFixedCode').FormulaU='2'
        for ($g=[int]$edge.GeometryCount-1; $g -ge 0; $g--) { $edge.DeleteSection(10+$g) }
        [void]$edge.AddSection(10)
        [void]$edge.AddRow(10,0,137)
        $edge.CellsU('Geometry1.NoFill').FormulaU='1'
        for ($i=0; $i -lt $points.Count; $i++) {
            $row=$edge.AddRow(10,-2,$(if ($i -eq 0) {138} else {139}))
            foreach ($axis in @(@('x','X',0),@('y','Y',1))) {
                $value=[double]$points[$i].($axis[0]); $coordinate=$axis[1]
                $formula=$value.ToString('R',[Globalization.CultureInfo]::InvariantCulture)+" in-Begin$coordinate"
                if ($i -eq 0 -or [Math]::Abs($value-$start.($axis[0])) -lt 0.000001) { $formula='0' }
                if ($i -eq $points.Count-1 -or ([Math]::Abs($value-$end.($axis[0])) -lt 0.000001 -and [Math]::Abs($value-$start.($axis[0])) -ge 0.000001)) { $formula="End$coordinate-Begin$coordinate" }
                $edge.CellsSRC(10,$row,$axis[2]).FormulaU=$formula
            }
        }
    }
}

function Get-NativeRoutingReport($Page) {
    $issues=@(); $attachments=@(); $edges=@($Page.shapes | Where-Object { $_.kind -eq 'edge' }); $segmentCount=0; $diagonalCount=0; $curvedCount=0
    foreach ($edge in $edges) {
        $route=Get-Value $edge 'routeGeometry'
        if ($null -eq $route -or -not $route.available) {
            $issues+=[pscustomobject]@{code='RouteGeometryUnavailable';edgeId=$edge.id;message='Actual stroke geometry could not be verified.'}
            continue
        }
        $segmentCount+=$route.segments.Count; $diagonalCount+=$route.diagonalCount
        $curvedCount+=Get-Value $route 'curvedCount' 0
        if (-not $route.orthogonal) { $issues+=[pscustomobject]@{code='DiagonalRoute';edgeId=$edge.id;message='Actual page-coordinate stroke contains a diagonal or curved segment.'} }
        $strokeEnds=@()
        foreach ($path in $route.paths) { $strokeEnds+=@($path.points[0],$path.points[-1]) }
        $endpointIds=@()
        foreach ($endpoint in @('source','target')) {
            $nativeIds=@($edge.($endpoint+'VisioIds'))
            if ($nativeIds.Count -ne 1) {
                $issues+=[pscustomobject]@{code='EndpointGlue';edgeId=$edge.id;message="$endpoint must be glued to exactly one full service shape."}
                continue
            }
            $endpointIds+=$nativeIds[0]
            $targets=@($Page.shapes | Where-Object { $_.visioId -eq $nativeIds[0] -and $_.kind -ne 'edge' })
            if ($targets.Count -ne 1 -or ($edge.$endpoint -and $targets[0].id -cne $edge.$endpoint)) {
                $issues+=[pscustomobject]@{code='EndpointIdentity';edgeId=$edge.id;message="$endpoint glue does not match the declared full service shape."}
                continue
            }
            $bounds=Get-Value $targets[0] 'serviceBounds'
            if ($null -eq $bounds) {
                $issues+=[pscustomobject]@{code='EndpointGeometryUnavailable';edgeId=$edge.id;message="$endpoint visible service bounds are unavailable."}
                continue
            }
            $x=$(if ($endpoint -eq 'source') { $edge.beginX } else { $edge.endX })
            $y=$(if ($endpoint -eq 'source') { $edge.beginY } else { $edge.endY })
            $side=Get-Value $edge ($endpoint+'Side') ''
            $style=Get-Value $targets[0] 'cardStyle' ''
            $region=$(if ($style -eq 'icon') { 'glyph' } elseif ($style -eq 'label') { 'caption' } else { 'body' })
            $captionBounds=Get-Value $targets[0] 'captionBounds'
            if ($style -eq 'icon' -and $side -in @('','bottom') -and $null -ne $captionBounds -and
                [Math]::Abs($y-$captionBounds.bottom) -le 0.01 -and
                $x -ge $captionBounds.left-0.01 -and $x -le $captionBounds.right+0.01) {
                $bounds=$captionBounds; $region='caption'
            }
            $portBounds=$bounds; $expectedRegion=$region
            if ($style -eq 'icon' -and $side -eq 'bottom') {
                if ($null -eq $captionBounds) {
                    $issues+=[pscustomobject]@{code='EndpointGeometryUnavailable';edgeId=$edge.id;message="$endpoint bottom port requires actual visible caption bounds."}
                    continue
                }
                $portBounds=$captionBounds; $expectedRegion='caption'
            }
            $attachments+=[pscustomobject]@{edgeId=$edge.id;endpoint=$endpoint;nodeId=$targets[0].id;region=$region;expectedRegion=$expectedRegion;side=$side;x=$x;y=$y;bounds=$bounds}
            if (@($strokeEnds | Where-Object { [Math]::Abs($_.x-$x) -le 0.01 -and [Math]::Abs($_.y-$y) -le 0.01 }).Count -eq 0) {
                $issues+=[pscustomobject]@{code='StrokeEndpointMismatch';edgeId=$edge.id;message="Actual stroke does not reach the glued $endpoint endpoint."}
            }
            $onX=[Math]::Abs($x-$bounds.left) -le 0.01 -or [Math]::Abs($x-$bounds.right) -le 0.01
            $onY=[Math]::Abs($y-$bounds.bottom) -le 0.01 -or [Math]::Abs($y-$bounds.top) -le 0.01
            if (-not (($onX -and $y -ge $bounds.bottom-0.01 -and $y -le $bounds.top+0.01) -or
                ($onY -and $x -ge $bounds.left-0.01 -and $x -le $bounds.right+0.01))) {
                $issues+=[pscustomobject]@{code='EndpointWhitespace';edgeId=$edge.id;message="$endpoint does not visibly reach the service perimeter."}
            }
            if ($side) {
                $position=[double](Get-Value $edge ($endpoint+'Position') 0.5)
                $px=$portBounds.left+($portBounds.right-$portBounds.left)*$position
                $py=$portBounds.bottom+($portBounds.top-$portBounds.bottom)*$position
                switch ($side) {
                    'left' { $px=$portBounds.left }
                    'right' { $px=$portBounds.right }
                    'top' { $py=$portBounds.top }
                    'bottom' { $py=$portBounds.bottom }
                }
                if ([Math]::Abs($px-$x) -gt 0.01 -or [Math]::Abs($py-$y) -gt 0.01) {
                    $issues+=[pscustomobject]@{code='SelectedPortMismatch';edgeId=$edge.id;message="$endpoint does not reach its selected service port."}
                }
                $outward=$false
                foreach ($segment in $route.segments) {
                    $dx=0.0; $dy=0.0
                    if ([Math]::Abs($segment.fromX-$x) -le 0.001 -and [Math]::Abs($segment.fromY-$y) -le 0.001) {
                        $dx=$segment.toX-$x; $dy=$segment.toY-$y
                    } elseif ([Math]::Abs($segment.toX-$x) -le 0.001 -and [Math]::Abs($segment.toY-$y) -le 0.001) {
                        $dx=$segment.fromX-$x; $dy=$segment.fromY-$y
                    }
                    if (($side -eq 'left' -and $dx -lt -0.001 -and [Math]::Abs($dy) -le 0.001) -or
                        ($side -eq 'right' -and $dx -gt 0.001 -and [Math]::Abs($dy) -le 0.001) -or
                        ($side -eq 'bottom' -and $dy -lt -0.001 -and [Math]::Abs($dx) -le 0.001) -or
                        ($side -eq 'top' -and $dy -gt 0.001 -and [Math]::Abs($dx) -le 0.001)) { $outward=$true; break }
                }
                if (-not $outward) { $issues+=[pscustomobject]@{code='EndpointDirection';edgeId=$edge.id;message="$endpoint stroke does not escape outward from its selected $side port."} }
            }
        }
        foreach ($node in $Page.shapes) {
            if ($node.kind -notin @('card','note')) { continue }
            $boundsList=@()
            foreach ($field in @('serviceBounds','captionBounds')) { $bounds=Get-Value $node $field; if ($null -ne $bounds) { $boundsList+=$bounds } }
            if (-not $boundsList.Count) {
                $issues+=[pscustomobject]@{code='ObstacleBoundsUnavailable';edgeId=$edge.id;message="Visible bounds unavailable for unrelated node '$($node.id)'."}
            }
            foreach ($bounds in $boundsList) {
                if (@($route.segments | Where-Object { Test-RouteCrossing $_ $bounds }).Count) {
                    $issues+=[pscustomobject]@{code='RouteObstruction';edgeId=$edge.id;message="Route crosses a node/caption '$($node.id)' rather than routing around it."}
                    break
                }
            }
        }
    }
    return [pscustomobject]@{valid=($issues.Count -eq 0);edgeCount=$edges.Count;segmentCount=$segmentCount;diagonalCount=$diagonalCount;curvedCount=$curvedCount;issues=$issues;attachments=$attachments;basis='Actual stroke segments, native full-shape glue, glyph/caption service perimeters and node/caption bounds; icon bottom ports may attach below the actual caption.'}
}

function Complete-NativeRouting($Document, [switch]$Force) {
    if (-not $Force -and (Read-Property $Document.DocumentSheet 'RoutingPolicy') -cne 'orthogonal-v1.6') { return }
    for ($p=1; $p -le $Document.Pages.Count; $p++) {
        $page=$Document.Pages.Item($p); $selection=$page.CreateSelection(0,0,$null); $count=0; $fixed=@()
        $index=Get-ShapeIndex $page
        for ($i=1; $i -le $page.Shapes.Count; $i++) {
            $shape=$page.Shapes.Item($i)
            if ($shape.OneD -and (Read-Property $shape 'Kind') -eq 'edge') {
                foreach ($pair in @(@('Source','BeginX'),@('Target','EndX'))) {
                    $id=Read-Property $shape ($pair[0]+'Id'); $side=Read-Property $shape ($pair[0]+'Side')
                    if (-not $id -or -not $side) { continue }
                    $targets=@(for ($c=1; $c -le $shape.Connects.Count; $c++) {
                        $connection=$shape.Connects.Item($c)
                        if ($connection.FromCell.Name -eq $pair[1]) { [int]$connection.ToSheet.ID }
                    })
                    if (-not $index.ContainsKey($id) -or $targets.Count -ne 1 -or $targets[0] -ne $index[$id].ID) {
                        throw "Cannot refresh $($pair[0]) port on '$(Read-Property $shape 'AvId')': actual full-service glue differs from its declared endpoint."
                    }
                    $position=Read-Property $shape ($pair[0]+'Position')
                    Glue-End $shape.CellsU($pair[1]) $index[$id] $side $(if ($position) { [double]$position } else { 0.5 })
                }
                $shape.CellsU('ShapeRouteStyle').FormulaU='1'
                $shape.CellsU('ConFixedCode').FormulaU='0'
                $shape.CellsU('ConLineRouteExt').FormulaU='1'
                $shape.CellsU('ConLineJumpCode').FormulaU='1'
                $shape.CellsU('Rounding').FormulaU='0'
                [void]$selection.Select($shape,2); $count++
            } elseif (-not $shape.OneD) {
                $fixed+=[pscustomobject]@{shape=$shape;x=$shape.CellsU('PinX').ResultIU;y=$shape.CellsU('PinY').ResultIU;width=$shape.CellsU('Width').ResultIU;height=$shape.CellsU('Height').ResultIU;angle=$shape.CellsU('Angle').ResultIU}
                if ($shape.CellExistsU('ShapeFixedCode',0)) {
                    $shape.CellsU('ShapeFixedCode').FormulaU=[string](([int]$shape.CellsU('ShapeFixedCode').ResultIU -band (-bnot 32)) -bor 65)
                }
                $permeable=$(if ((Read-Property $shape 'Kind') -in @('card','note')) { '0' } else { '1' })
                foreach ($cell in @('ShapePermeableX','ShapePermeableY')) { if ($shape.CellExistsU($cell,0)) { $shape.CellsU($cell).FormulaU=$permeable } }
            }
        }
        if ($count) { [void]$selection.Layout() }
        Repair-NativeBoundaryRoutes $page $index
        foreach ($record in $fixed) {
            foreach ($pair in @(@('x','PinX'),@('y','PinY'),@('width','Width'),@('height','Height'),@('angle','Angle'))) {
                if ([Math]::Abs($record.($pair[0])-$record.shape.CellsU($pair[1]).ResultIU) -gt 0.001) { throw 'Connector-only routing moved or resized a node; rolling back instead of altering architecture layout.' }
            }
        }
    }
}

function Add-Edge($Page, $Edge, $Index) {
    $style = $script:edgeStyles[$Edge.kind]
    $shape = $Page.Drop($script:app.ConnectorToolDataObject, 0, 0)
    $sourceId = Get-Value $Edge 'source'; $targetId = Get-Value $Edge 'target'
    $start = 'right'; $end = 'left'
    $dx = 0; $dy = 0
    if ($sourceId -and $targetId) {
        $dx = $Index[$targetId].CellsU('PinX').ResultIU - $Index[$sourceId].CellsU('PinX').ResultIU
        $dy = $Index[$targetId].CellsU('PinY').ResultIU - $Index[$sourceId].CellsU('PinY').ResultIU
    }
    $iconEndpoint=($sourceId -and (Read-Property $Index[$sourceId] 'CardStyle') -eq 'icon') -or ($targetId -and (Read-Property $Index[$targetId] 'CardStyle') -eq 'icon')
    if ($iconEndpoint -or [Math]::Abs($dx) -ge [Math]::Abs($dy)) {
        $start = $(if ($dx -ge 0) { 'right' } else { 'left' })
        $end = $(if ($dx -ge 0) { 'left' } else { 'right' })
    } else {
        $start = $(if ($dy -ge 0) { 'top' } else { 'bottom' })
        $end = $(if ($dy -ge 0) { 'bottom' } else { 'top' })
    }
    if ($sourceId) { Glue-End $shape.CellsU('BeginX') $Index[$sourceId] (Get-Value $Edge 'sourceSide' $start) ([double](Get-Value $Edge 'sourcePosition' 0.5)) }
    else {
        $shape.CellsU('BeginX').ResultIU = $Edge.beginX
        $shape.CellsU('BeginY').ResultIU = $Edge.beginY
    }
    if ($targetId) { Glue-End $shape.CellsU('EndX') $Index[$targetId] (Get-Value $Edge 'targetSide' $end) ([double](Get-Value $Edge 'targetPosition' 0.5)) }
    else {
        $shape.CellsU('EndX').ResultIU = $Edge.endX
        $shape.CellsU('EndY').ResultIU = $Edge.endY
    }
    $shape.Text = Get-Value $Edge 'label' ''
    Set-TextStyle $shape 10
    $shape.CellsU('LineColor').FormulaU = Get-Value $Edge 'color' $style[0]
    $shape.CellsU('LineWeight').FormulaU = '1.5 pt'
    $shape.CellsU('LinePattern').FormulaU = [string]$style[1]
    $shape.CellsU('BeginArrow').FormulaU = '0'
    $shape.CellsU('EndArrow').FormulaU = [string]$style[2]
    Set-EdgeSemantics $shape $Edge
    $shape.CellsU('TextBkgnd').FormulaU = 'RGB(255,255,255)+1'
    $shape.CellsU('ShapeRouteStyle').FormulaU = '1'
    $shape.CellsU('ConFixedCode').FormulaU = '0'
    $shape.CellsU('ConLineRouteExt').FormulaU = '1'
    $shape.NameU = 'av-' + $Edge.id
    Set-Property $shape 'AvId' $Edge.id
    Set-Property $shape 'Kind' 'edge'
    Set-Property $shape 'Relationship' $Edge.kind
    Set-Property $shape 'SourceId' $sourceId
    Set-Property $shape 'TargetId' $targetId
    foreach ($endpoint in @('source','target')) {
        $prefix = $(if ($endpoint -eq 'source') { 'Source' } else { 'Target' })
        Set-Property $shape ($prefix+'Side') (Get-Value $Edge ($endpoint+'Side') $(if ($endpoint -eq 'source') { $start } else { $end }))
        Set-Property $shape ($prefix+'Position') ([string](Get-Value $Edge ($endpoint+'Position') 0.5))
    }
    Set-Property $shape 'State' (Get-Value $Edge 'state' 'Illustrative')
    Set-Property $shape 'Purpose' (Get-Value $Edge 'purpose' (Get-Value $Edge 'label' ''))
    Set-Metadata $shape $Edge
    Add-ToLayer $Page $shape $style[3]
    return $shape
}

function Set-EdgeSemantics($Shape, $Spec) {
    if (Test-Field $Spec 'routeStyle') { $Shape.CellsU('ShapeRouteStyle').FormulaU=$(if ($Spec.routeStyle -eq 'straight') { '2' } else { '1' }) }
    if (Test-Field $Spec 'direction') {
        $Shape.CellsU('BeginArrow').FormulaU = $(if ($Spec.direction -in @('backward','both')) { '13' } else { '0' })
        $Shape.CellsU('EndArrow').FormulaU = $(if ($Spec.direction -in @('forward','both')) { '13' } else { '0' })
    }
    if (Test-Field $Spec 'dashed') { $Shape.CellsU('LinePattern').FormulaU = $(if ($Spec.dashed) { '2' } else { '1' }) }
}

function Get-ShapeIndex($Page) {
    $index = @{}
    for ($i = 1; $i -le $Page.Shapes.Count; $i++) {
        $shape = $Page.Shapes.Item($i)
        $id = Read-Property $shape 'AvId'
        if ($id) {
            if ($index.ContainsKey($id)) { throw "Duplicate AvId '$id' on page '$($Page.Name)'." }
            $index[$id] = $shape
        }
    }
    return $index
}

function Confirm-ArchitecturePack($Model, [switch]$Required) {
    if (-not (Test-Field $Model 'outputContract')) {
        if ($Required) { throw 'New requires outputContract architecture-pack-v1.6. Use -LegacyModel only for explicit legacy creation.' }
        return
    }
    if ((Get-Value $Model 'outputContract') -isnot [string] -or $Model.outputContract -cne 'architecture-pack-v1.6') { throw 'Unknown outputContract; expected architecture-pack-v1.6.' }
    Confirm-Number (Get-Value $Model 'schemaVersion') 'schemaVersion' 1 1
    if ((Get-Value $Model 'presentationProfile') -isnot [string] -or $Model.presentationProfile -cnotin @('enterprise','reference')) { throw 'Architecture pack presentationProfile must be enterprise or reference.' }
    $hardening=Get-Value $Model 'hardening'
    Confirm-Fields $hardening @('status','reason') 'hardening'
    if ((Get-Value $hardening 'status') -isnot [string] -or $hardening.status -cnotin @('proposed','already-enterprise','not-applicable')) { throw 'Invalid hardening status.' }
    $reason=Get-Value $hardening 'reason'
    if ($reason -isnot [string] -or [string]::IsNullOrWhiteSpace($reason)) { throw 'Hardening reason must be a nonempty string.' }
    $pages=Get-Value $Model 'pages'
    if ($pages -isnot [array] -or $pages.Count -ne 3) { throw 'Architecture pack requires exactly three ordered pages.' }
    $views=@('main','hardening','flowchart')
    for ($i=0; $i -lt 3; $i++) {
        $page=$pages[$i]
        if ((Get-Value $page 'view') -isnot [string] -or $page.view -cne $views[$i]) { throw 'Architecture pack page views must be main, hardening, flowchart in that order.' }
        $role=$(if ($i -eq 0 -or ($i -eq 1 -and $hardening.status -eq 'proposed')) { 'diagram' } else { 'notes' })
        if ((Get-Value $page 'role' 'diagram') -isnot [string] -or (Get-Value $page 'role' 'diagram') -cne $role) { throw "Architecture pack $($views[$i]) page role must be $role." }
        if ((Get-Value $page 'nodes') -isnot [array] -or (Get-Value $page 'edges') -isnot [array]) { throw 'Architecture pack pages require nodes and edges arrays.' }
    }
    function Pack-Words([string]$Text) { return [regex]::Matches($Text,'[\p{L}\p{N}]+(?:[-/][\p{L}\p{N}]+)*').Count }
    function Pack-Signature($Page) {
        # Ignore layout, identifiers, provenance and presentation-only changes.
        $nodes=@{}; $entities=@()
        foreach ($node in $Page.nodes) {
            if ($node.kind -eq 'note') { continue }
            $semantic=@($node.kind, (Get-Value $node 'label' ''), (Get-Value $node 'details' ''),
                (Get-Value $node 'purpose' (Get-Value $node 'label' '')), (Get-Value $node 'boundaryType' ''))
            $key=ConvertTo-Json -InputObject $semantic -Compress
            $nodes[$node.id]=$key; $entities+=$key
        }
        $relations=@()
        foreach ($node in $Page.nodes) {
            $parent=Get-Value $node 'parent' ''
            if ($parent -and $nodes.ContainsKey($parent) -and $nodes.ContainsKey($node.id)) {
                $relations+=ConvertTo-Json -InputObject @('containment',$nodes[$parent],$nodes[$node.id]) -Compress
            }
        }
        foreach ($edge in $Page.edges) {
            $source=Get-Value $edge 'source'; $target=Get-Value $edge 'target'
            if ($null -eq $source -or $null -eq $target -or -not $nodes.ContainsKey($source) -or -not $nodes.ContainsKey($target)) { continue }
            $defaultDirection=$(if ($edge.kind -in @('peering','dns','governance','association')) { 'none' } else { 'forward' })
            $relations+=ConvertTo-Json -InputObject @($nodes[$source],$nodes[$target],$edge.kind,
                (Get-Value $edge 'label' ''),(Get-Value $edge 'details' ''),(Get-Value $edge 'direction' $defaultDirection)) -Compress
        }
        return ConvertTo-Json -InputObject @(@($entities | Sort-Object),@($relations | Sort-Object)) -Depth 8 -Compress
    }
    foreach ($page in @($pages[0]) + @($pages[1] | Where-Object { $hardening.status -eq 'proposed' })) {
        $entities=@{}; $realEdges=0
        foreach ($node in $page.nodes) { if ((Get-Value $node 'kind') -eq 'card') { $entities[$node.id]=$true } }
        foreach ($edge in $page.edges) {
            $source=Get-Value $edge 'source'; $target=Get-Value $edge 'target'
            if ($null -ne $source -and $null -ne $target -and $source -cne $target -and $entities.ContainsKey($source) -and $entities.ContainsKey($target)) { $realEdges++ }
        }
        if ($entities.Count -lt 2 -or $realEdges -lt 1) { throw 'Architecture pack diagram pages require at least two real card entities and an attached non-self relationship.' }
    }
    if ($hardening.status -eq 'proposed') {
        if ((Pack-Signature $pages[0]) -ceq (Pack-Signature $pages[1])) {
            throw 'Proposed hardening must contain a semantic architecture change with its reason, not a duplicate main design.'
        }
    } else {
        if (@($pages[1].nodes | Where-Object { $_.kind -ne 'note' }).Count -or $pages[1].edges.Count) {
            throw 'Non-proposed hardening must be an applicability review using notes, not an invented redundant design.'
        }
        if (@($pages[1].nodes | Where-Object { (Pack-Words (Get-Value $_ 'label' '')) -ge 40 }).Count -eq 0) {
            throw 'Hardening applicability review requires a visible note with at least 40 words.'
        }
    }
    $mainNodes=@{}; $mainEdges=@{}; $coveredNodes=@{}; $coveredEdges=@{}; $flowCards=@{}
    foreach ($node in $pages[0].nodes) { $mainNodes[$node.id]=$node.kind }
    foreach ($edge in $pages[0].edges) { $mainEdges[$edge.id]=$true }
    foreach ($node in $pages[2].nodes) {
        if ($node.kind -ne 'card') { continue }
        if ((Get-Value $node 'cardStyle') -cnotin @('detail','standard') -or (Pack-Words (Get-Value $node 'label' '')) -lt 10) {
            throw 'Flowchart cards require detail or standard cardStyle and meaningful labels of at least 10 words.'
        }
        Confirm-Metadata $node
        foreach ($field in @('mainNodeIds','mainEdgeIds')) {
            if ((Get-Value $node $field) -isnot [array]) { throw "Flowchart cards require $field string arrays referencing MAIN only." }
        }
        if ($node.mainNodeIds.Count + $node.mainEdgeIds.Count -eq 0) { throw 'Each flowchart card must reference MAIN nodes or edges.' }
        $flowCards[$node.id]=$true
    }
    foreach ($page in $pages) {
        foreach ($node in @($page.nodes) + @($page.edges)) {
            foreach ($field in @('mainNodeIds','mainEdgeIds')) {
                if (-not (Test-Field $node $field)) { continue }
                if ($page.view -cne 'flowchart' -or $node.kind -cne 'card') { throw 'MAIN coverage metadata is only valid on flowchart cards.' }
                $known=$(if ($field -eq 'mainNodeIds') { $mainNodes } else { $mainEdges })
                $covered=$(if ($field -eq 'mainNodeIds') { $coveredNodes } else { $coveredEdges })
                foreach ($id in (Get-Value $node $field)) {
                    if (-not $known.ContainsKey($id) -or $id -cnotin @($known.Keys)) { throw "Unknown MAIN reference '$id' in $field; hardening-only IDs are not allowed." }
                    $covered[$id]=$true
                }
            }
        }
    }
    foreach ($id in $mainNodes.Keys) {
        if ($mainNodes[$id] -eq 'card' -and -not $coveredNodes.ContainsKey($id)) { throw "Flowchart is missing MAIN card coverage for '$id'." }
    }
    foreach ($id in $mainEdges.Keys) {
        if (-not $coveredEdges.ContainsKey($id)) { throw "Flowchart is missing MAIN edge coverage for '$id'." }
    }
    if ($flowCards.Count -lt 2 -or $pages[2].edges.Count -lt 1) { throw 'Flowchart requires at least two cards and connected directed edges.' }
    $adjacent=@{}
    foreach ($id in $flowCards.Keys) { $adjacent[$id]=@() }
    foreach ($edge in $pages[2].edges) {
        $source=Get-Value $edge 'source'; $target=Get-Value $edge 'target'
        $defaultDirection=$(if ((Get-Value $edge 'kind') -in @('peering','dns','governance','association')) { 'none' } else { 'forward' })
        if ($null -eq $source -or $null -eq $target -or $source -eq $target -or
            -not $flowCards.ContainsKey($source) -or -not $flowCards.ContainsKey($target) -or
            (Get-Value $edge 'direction' $defaultDirection) -cnotin @('forward','backward','both')) {
            throw 'Flowchart edges must be directed, attached between distinct flowchart cards.'
        }
        $adjacent[$source]+=$target; $adjacent[$target]+=$source
    }
    $visited=@{}; $queue=[Collections.Generic.Queue[string]]::new()
    $queue.Enqueue([string]@($flowCards.Keys)[0])
    while ($queue.Count) {
        $id=$queue.Dequeue()
        if ($visited.ContainsKey($id)) { continue }
        $visited[$id]=$true
        foreach ($next in $adjacent[$id]) { if (-not $visited.ContainsKey($next)) { $queue.Enqueue($next) } }
    }
    if ($visited.Count -ne $flowCards.Count) { throw 'Flowchart cards must form one connected directed flow.' }
    if (@($pages[2].nodes | Where-Object { $_.kind -eq 'note' -and (Pack-Words (Get-Value $_ 'label' '')) -ge 40 }).Count -eq 0) {
        throw 'Flowchart requires a visible MAIN architecture narrative note with at least 40 words.'
    }
}

function Confirm-Model($Model) {
    Confirm-Number (Get-Value $Model 'schemaVersion') 'schemaVersion' 1 1
    $pages = Get-Value $Model 'pages'
    if ($pages -isnot [array] -or $pages.Count -eq 0) { throw 'Model requires a nonempty pages array.' }
    if ((Test-Field $Model 'title') -and $Model.title -isnot [string]) { throw 'Model title must be a string.' }
    if ((Get-Value $Model 'presentationProfile' 'legacy') -cnotin @('enterprise','reference','legacy')) { throw 'presentationProfile must be enterprise, reference, or legacy.' }
    if ((Test-Field $Model 'conversionMode') -and
        $Model.conversionMode -cnotin @('new-design','faithful','reference-plus-proposal')) {
        throw 'conversionMode must be new-design, faithful, or reference-plus-proposal.'
    }
    if ((Test-Field $Model 'referenceContract') -and
        ($Model.referenceContract -isnot [string] -or [string]::IsNullOrWhiteSpace($Model.referenceContract))) {
        throw 'referenceContract must name a source contract JSON file.'
    }
    $names = @{}
    foreach ($page in $pages) {
        $name = Get-Value $page 'name'
        if ($name -isnot [string] -or [string]::IsNullOrWhiteSpace($name) -or $names.ContainsKey($name)) { throw "Invalid or repeated page name: $name" }
        $names[$page.name] = $true
        if ((Get-Value $page 'role' 'diagram') -cnotin @('diagram','notes')) { throw 'Page role must be diagram or notes.' }
        foreach ($dimension in @('width','height')) { Confirm-Number (Get-Value $page $dimension) "Page $dimension" 0.001 1000000 }
        if ((Get-Value $page 'nodes') -isnot [array] -or (Get-Value $page 'edges') -isnot [array]) { throw 'Each page requires nodes and edges arrays.' }
        foreach ($field in @('title','subtitle')) {
            if ((Test-Field $page $field) -and $page.$field -isnot [string]) { throw "Page $field must be a string." }
        }
        if (Test-Field $page 'titleFontSize') { Confirm-Number $page.titleFontSize 'titleFontSize' 8 72 }
        if ((Test-Field $page 'furniture') -and $page.furniture -isnot [bool]) { throw 'furniture must be a boolean.' }
        $footer = Get-Value $page 'footer'
        if ($null -ne $footer -and $footer -isnot [string] -and -not ($footer -is [bool] -and -not $footer)) { throw 'footer must be a string, null, or false.' }
        if ((Test-Field $page 'allowOffPage') -and $page.allowOffPage -isnot [bool]) { throw 'allowOffPage must be a boolean.' }
        if ((Get-Value $page 'furniture' $true) -and ($page.width -le 1.2 -or $page.height -le 1.3)) { throw 'Page furniture requires width > 1.2 and height > 1.3 inches.' }
        $nodes = @{}; $all = @{}
        foreach ($node in $page.nodes) {
            Confirm-Identifier (Get-Value $node 'id')
            if ($all.ContainsKey($node.id)) { throw "Duplicate id: $($node.id)" }
            if ((Get-Value $node 'kind') -cnotin @('container','card','note')) { throw "Unsupported node kind: $(Get-Value $node 'kind')" }
            if ($null -eq (Get-Value $node 'label')) { throw "Missing label on $($node.id)." }
            Confirm-Metadata $node
            if ((Test-Field $node 'sourceId') -and $node.sourceId) { Confirm-Identifier $node.sourceId }
            if ((Test-Field $node 'cardStyle') -and
                ($node.kind -ne 'card' -or $node.cardStyle -cnotin @('standard','detail','icon','label'))) {
                throw 'cardStyle must be standard, detail, icon, or label and is only supported on cards.'
            }
            $profile=$(if (Test-Field $Model 'outputContract') { 'enterprise' } else { Get-Value $Model 'presentationProfile' 'legacy' })
            $cardStyle=Get-CardStyle $node $profile
            if ((Test-Field $node 'containerStyle') -and ($node.kind -ne 'container' -or $node.containerStyle -cne 'boundary')) { throw 'containerStyle must be boundary and is only supported on containers.' }
            if ((Test-Field $node 'boundaryType') -and ($node.kind -ne 'container' -or [string]::IsNullOrWhiteSpace($node.boundaryType))) { throw 'boundaryType requires a nonempty semantic boundary name on a container.' }
            if (((Test-Field $node 'displayLabel') -or (Test-Field $node 'iconSize')) -and
                ($node.kind -ne 'card' -or $cardStyle -notin @('icon','label'))) { throw 'displayLabel/iconSize require an icon or label card.' }
            if (Test-Field $node 'iconSize') {
                if ($cardStyle -ne 'icon') { throw 'iconSize is only supported on icon cards.' }
                Confirm-Number $node.iconSize 'iconSize' 0.35 3
            }
            if ((Get-Value $node 'cardStyle' 'standard') -eq 'detail' -and $node.height -lt 0.9) {
                throw 'Detail cards must be at least 0.9 inches high.'
            }
            foreach ($colorField in @('color','fill')) {
                $color = Get-Value $node $colorField ''
                if ((Test-Field $node $colorField) -and ($color -isnot [string] -or $color -notmatch '^RGB\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$')) {
                    throw "$colorField must be an RGB(r,g,b) color, not a general formula."
                }
                if ($color -and (@($color -split '[^\d]+' | Where-Object { $_ -ne '' -and [int]$_ -gt 255 }).Count -gt 0)) {
                    throw "RGB channels must be between 0 and 255 on $($node.id)."
                }
            }
            foreach ($field in @('x','y','width','height')) {
                Confirm-Number (Get-Value $node $field) "$field on $($node.id)"
            }
            if ($node.width -le 0 -or $node.height -le 0) { throw "Invalid bounds on $($node.id)." }
            if (-not (Get-Value $page 'allowOffPage' $false) -and
                ($node.x-$node.width/2 -lt -0.001 -or $node.x+$node.width/2 -gt $page.width+0.001 -or
                $node.y-$node.height/2 -lt -0.001 -or $node.y+$node.height/2 -gt $page.height+0.001)) { throw "Node $($node.id) is outside its page." }
            $icon = Get-Value $node 'icon' ''
            if ($icon -isnot [string] -or ($icon -and -not $script:icons.ContainsKey($icon))) { throw "Unsupported icon '$icon'." }
            $iconRef = Get-Value $node 'iconRef' ''
            if ($iconRef -isnot [string]) { throw 'iconRef must be a catalog id string.' }
            if ($icon -and $iconRef) { throw 'Use icon or iconRef, not both.' }
            if ($node.kind -eq 'card' -and $cardStyle -eq 'icon' -and -not ($icon -or $iconRef)) { throw 'cardStyle icon requires a valid icon or iconRef; use cardStyle label for a generic function.' }
            if ($node.kind -eq 'card' -and $cardStyle -eq 'label' -and ($icon -or $iconRef)) { throw 'cardStyle label cannot contain an icon; use cardStyle icon.' }
            if (($icon -or $iconRef) -and ($node.kind -ne 'card' -or ($cardStyle -ne 'icon' -and $node.width -lt 1.1))) { throw 'Icons require a card at least 1.1 inches wide.' }
            if ($iconRef) { [void](Resolve-IconRef $iconRef) }
            if (Test-Field $node 'fontSize') { Confirm-Number $node.fontSize 'fontSize' 1 200 }
            if ($node.kind -eq 'card' -and $cardStyle -in @('icon','label')) { [void](Get-CaptionLayout $node $cardStyle) }
            if (Test-Field $node 'linePattern') {
                Confirm-Number $node.linePattern 'linePattern' 0 23
                if ([int]$node.linePattern -ne $node.linePattern) { throw 'linePattern must be an integer.' }
            }
            $nodes[$node.id] = $node; $all[$node.id] = $true
        }
        foreach ($node in $page.nodes) {
            $parent = Get-Value $node 'parent' ''
            if ($parent -isnot [string]) { throw 'parent must be a shape id string.' }
            if ($parent) {
                if (-not $nodes.ContainsKey($parent) -or $nodes[$parent].kind -ne 'container') { throw "Invalid parent '$parent' for $($node.id)." }
                $boundary = $nodes[$parent]
                if ([Math]::Abs($node.x-$boundary.x)+$node.width/2 -gt $boundary.width/2+0.001 -or
                    [Math]::Abs($node.y-$boundary.y)+$node.height/2 -gt $boundary.height/2+0.001) {
                    throw "Node $($node.id) is outside parent $parent."
                }
                $ancestors = @{$node.id = $true}
                $current = $parent
                while ($current) {
                    if ($ancestors.ContainsKey($current)) { throw "Container cycle at $current." }
                    $ancestors[$current] = $true
                    if (-not $nodes.ContainsKey($current)) { throw "Missing ancestor $current." }
                    $current = Get-Value $nodes[$current] 'parent' ''
                }
            }
        }
        foreach ($edge in $page.edges) {
            Confirm-Identifier (Get-Value $edge 'id')
            if ($all.ContainsKey($edge.id)) { throw "Duplicate edge id: $($edge.id)" }
            foreach ($endpoint in @('source','target')) {
                if (-not (Test-Field $edge $endpoint)) { throw "Missing endpoint on $($edge.id)." }
                $id = Get-Value $edge $endpoint
                if ($null -ne $id) {
                    Confirm-Identifier $id
                    if (-not $nodes.ContainsKey($id)) { throw "Missing endpoint on $($edge.id)." }
                } else {
                    $prefix = $(if ($endpoint -eq 'source') { 'begin' } else { 'end' })
                    foreach ($axis in @('X','Y')) { Confirm-Number (Get-Value $edge ($prefix+$axis)) "$prefix$axis for unattached edge $($edge.id)" }
                }
            }
            $kind = Get-Value $edge 'kind' ''
            if ($kind -isnot [string] -or -not $script:edgeStyles.ContainsKey($kind)) { throw "Unsupported relationship $kind." }
            Confirm-Metadata $edge
            if ((Test-Field $edge 'sourceId') -and $edge.sourceId) { Confirm-Identifier $edge.sourceId }
            if (Test-Field $edge 'color') {
                $color = $edge.color
                if ($color -isnot [string] -or $color -notmatch '^RGB\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$' -or
                    @($color -split '[^\d]+' | Where-Object { $_ -ne '' -and [int]$_ -gt 255 }).Count) {
                    throw 'Edge color must be RGB channels from 0 to 255.'
                }
            }
            Confirm-EdgeSemantics $edge
            foreach ($field in @('sourceSide','targetSide')) {
                if ((Test-Field $edge $field) -and $edge.$field -cnotin @('left','right','top','bottom')) { throw "Invalid $field on $($edge.id)." }
            }
            $all[$edge.id] = $true
        }
    }
    Confirm-ArchitecturePack $Model
}

function Confirm-EdgeSemantics($Spec) {
    if ((Test-Field $Spec 'routeStyle') -and $Spec.routeStyle -cnotin @('straight','orthogonal')) { throw 'routeStyle must be straight or orthogonal.' }
    if ((Test-Field $Spec 'direction') -and $Spec.direction -cnotin @('forward','backward','both','none')) { throw 'Invalid direction; use forward, backward, both, or none.' }
    if ((Test-Field $Spec 'dashed') -and $Spec.dashed -isnot [bool]) { throw 'dashed must be a boolean.' }
    foreach ($field in @('sourcePosition','targetPosition')) {
        if (Test-Field $Spec $field) { Confirm-Number $Spec.$field $field 0 1 }
    }
}

function Add-PageFurniture($Page, $Spec) {
    $w = [double]$Spec.width; $h = [double]$Spec.height
    $Page.PageSheet.CellsU('PageWidth').ResultIU = $w
    $Page.PageSheet.CellsU('PageHeight').ResultIU = $h
    $Page.PageSheet.CellsU('PrintPageOrientation').FormulaU = '2'
    Set-Property $Page.PageSheet 'AvFurniture' ([string](Get-Value $Spec 'furniture' $true))
    if (Test-Field $Spec 'role') { Set-Property $Page.PageSheet 'AvPageRole' $Spec.role }
    if (Test-Field $Spec 'view') { Set-Property $Page.PageSheet 'AvPageView' $Spec.view }
    if (-not (Get-Value $Spec 'furniture' $true)) { return }
    $background = $Page.DrawRectangle(0, 0, $w, $h)
    $background.NameU = 'av-background'
    Set-Property $background 'AvFurniture' 'background'
    $background.CellsU('FillForegnd').FormulaU = 'RGB(255,255,255)'
    $background.CellsU('LinePattern').FormulaU = '0'
    $background.CellsU('LockSelect').FormulaU = '1'
    $background.CellsU('LockMoveX').FormulaU = '1'
    $background.CellsU('LockMoveY').FormulaU = '1'
    Add-ToLayer $Page $background 'Background'
    $title = $Page.DrawRectangle(0.6, ($h-0.85), ($w-0.6), ($h-0.22))
    $title.Text = Get-Value $Spec 'title' $Spec.name
    Set-Property $title 'AvFurniture' 'title'
    $title.CellsU('LinePattern').FormulaU = '0'
    $title.CellsU('FillPattern').FormulaU = '0'
    Set-TextStyle $title ([double](Get-Value $Spec 'titleFontSize' 23)) $true
    $title.CellsU('Char.Style').FormulaU = '1'
    $subtitle = $Page.DrawRectangle(0.6, ($h-1.22), ($w-0.6), ($h-0.84))
    $subtitle.Text = Get-Value $Spec 'subtitle' ''
    Set-Property $subtitle 'AvFurniture' 'subtitle'
    $subtitle.CellsU('LinePattern').FormulaU = '0'
    $subtitle.CellsU('FillPattern').FormulaU = '0'
    Set-TextStyle $subtitle 11 $true
    $footerText = Get-Value $Spec 'footer' 'ILLUSTRATIVE | Blue dashed: peering | Orange arrows: traffic | Purple dashed: DNS | Gray: governance | Green dashed: telemetry'
    if ($footerText -is [string] -and $footerText) {
        $footer = $Page.DrawRectangle(0.6, 0.14, ($w-0.6), 0.62)
        $footer.Text = $footerText
        Set-Property $footer 'AvFurniture' 'footer'
        $footer.CellsU('LinePattern').FormulaU = '0'
        $footer.CellsU('FillPattern').FormulaU = '0'
        Set-TextStyle $footer 10 $true
        Add-ToLayer $Page $footer 'Annotations'
    }
    foreach ($shape in @($title,$subtitle)) { Add-ToLayer $Page $shape 'Annotations' }
}

function Confirm-PackMerge($Document, $Model) {
    $existing=Read-Property $Document.DocumentSheet 'OutputContract'
    if ($Document.DocumentSheet.CellExistsU('Prop.OutputContract',0)) {
        if ($existing -cne 'architecture-pack-v1.6' -or (Get-Value $Model 'outputContract') -cne $existing) {
            throw 'Merge cannot remove or change the document outputContract.'
        }
        Confirm-ArchitecturePack $Model -Required
        if ($Document.Pages.Count -ne 3) { throw 'Merge cannot repair or change architecture pack page count.' }
        for ($i=0; $i -lt 3; $i++) {
            $page=$Document.Pages.Item($i+1)
            if ($page.Name -cne $Model.pages[$i].name -or
                (Read-Property $page.PageSheet 'AvPageView') -cne $Model.pages[$i].view -or
                (Read-Property $page.PageSheet 'AvPageRole') -cne (Get-Value $Model.pages[$i] 'role' 'diagram')) {
                throw 'Merge cannot add, rename, reorder or change architecture pack pages.'
            }
        }
        $current=Read-Property $Document.DocumentSheet 'Hardening' | ConvertFrom-Json
        if ($current.status -cne $Model.hardening.status -or $current.reason -cne $Model.hardening.reason) {
            throw 'Merge is additive and cannot change hardening status or reason. Use an approved new model.'
        }
    } elseif (Test-Field $Model 'outputContract') {
        throw 'Merge cannot convert a legacy document into an architecture pack; use New.'
    }
}

function Confirm-NativeArchitecturePack($Document) {
    if ($Document.DocumentSheet.CellExistsU('Prop.OutputContract',0)) { [void](Export-Model $Document) }
}

function Apply-Model($Document, $Model, [bool]$Creating) {
    if (-not $Creating) { Confirm-PackMerge $Document $Model }
    if ($Creating -and (Test-Field $Model 'outputContract')) {
        Set-Property $Document.DocumentSheet 'OutputContract' $Model.outputContract
        Set-Property $Document.DocumentSheet 'Hardening' (ConvertTo-Json -InputObject $Model.hardening -Depth 8 -Compress)
        Set-Property $Document.DocumentSheet 'RoutingPolicy' 'orthogonal-v1.6'
    }
    if (Test-Field $Model 'presentationProfile') {
        $existingProfile=Read-Property $Document.DocumentSheet 'PresentationProfile'
        if (-not $existingProfile) { $existingProfile='legacy' }
        if (-not $Creating -and $existingProfile -ne $Model.presentationProfile) {
            throw 'Merge cannot switch presentation profile or restyle existing shapes. Use an approved new working copy/model for that change.'
        }
        Set-Property $Document.DocumentSheet 'PresentationProfile' $Model.presentationProfile
    }
    foreach ($spec in $Model.pages) {
        $page = $null
        for ($i = 1; $i -le $Document.Pages.Count; $i++) {
            if ($Document.Pages.Item($i).Name -eq $spec.name) { $page = $Document.Pages.Item($i); break }
        }
        if ($null -eq $page) {
            if ($Creating -and $Document.Pages.Count -eq 1 -and $Document.Pages.Item(1).Shapes.Count -eq 0) {
                $page = $Document.Pages.Item(1)
            } else { $page = $Document.Pages.Add() }
            $page.Name = $spec.name
        }
        if ($Creating -or $page.Shapes.Count -eq 0) { Add-PageFurniture $page $spec }
        $index = Get-ShapeIndex $page
        $added = @{}
        foreach ($node in $spec.nodes) {
            if (-not $index.ContainsKey($node.id)) {
                $profile=$(if (Test-Field $Model 'outputContract') { 'enterprise' } else { Get-Value $Model 'presentationProfile' 'legacy' })
                $index[$node.id] = Add-Node $page $node $profile
                $added[$node.id] = $true
            } elseif ((Read-Property $index[$node.id] 'Kind') -ne $node.kind) {
                throw "Existing id $($node.id) has a different kind. Merge does not replace shapes."
            }
        }
        foreach ($node in $spec.nodes) {
            $parent = Get-Value $node 'parent' ''
            if ($parent -and $added.ContainsKey($node.id)) {
                [void]$index[$parent].ContainerProperties.AddMember($index[$node.id], 2)
            }
        }
        foreach ($edge in $spec.edges) {
            if (-not $index.ContainsKey($edge.id)) { $index[$edge.id] = Add-Edge $page $edge $index }
            elseif ((Read-Property $index[$edge.id] 'SourceId') -ne $edge.source -or
                    (Read-Property $index[$edge.id] 'TargetId') -ne $edge.target) {
                throw "Edge $($edge.id) changed endpoints. Merge only adds; it never silently rewires."
            }
        }
        # Inner containers must remain above their ancestors, without relaying out user shapes.
        $containers = @($spec.nodes | Where-Object { $_.kind -eq 'container' })
        for ($i = $containers.Count-1; $i -ge 0; $i--) {
            if ($added.ContainsKey($containers[$i].id)) { [void]$index[$containers[$i].id].SendToBack() }
        }
        for ($i = 1; $i -le $page.Shapes.Count; $i++) {
            if ($page.Shapes.Item($i).NameU -eq 'av-background') { [void]$page.Shapes.Item($i).SendToBack(); break }
        }
    }
    if ($Creating -and (Test-Field $Model 'outputContract')) {
        for ($i=0; $i -lt $Model.pages.Count; $i++) { $Document.Pages.Item($Model.pages[$i].name).Index=$i+1 }
    }
    Complete-NativeRouting $Document -Force:$Creating
}

function Test-Container($Shape) {
    return $Shape.CellExistsU('User.msvStructureType', 0) -and $Shape.CellsU('User.msvStructureType').ResultStr(0) -eq 'Container'
}

function Test-PortableBoundary($Record, [string]$Renderer) {
    return $Renderer -ceq 'azure-visio-1.6' -and -not (Get-Value $Record 'isContainer' $false) -and
        (Get-Value $Record 'id' '') -and (Get-Value $Record 'kind' '') -ceq 'container' -and
        (Get-Value $Record 'containerStyle' '') -ceq 'boundary' -and
        -not [string]::IsNullOrWhiteSpace((Get-Value $Record 'boundaryType' ''))
}

function Get-NativeShapes($Shapes, [int]$RootId = -1) {
    for ($i = 1; $i -le $Shapes.Count; $i++) {
        $shape = $Shapes.Item($i)
        $root = $(if ($RootId -lt 0) { [int]$shape.ID } else { $RootId })
        [pscustomobject]@{ shape=$shape; rootId=$root; topLevel=($RootId -lt 0) }
        Get-NativeShapes $shape.Shapes $root
    }
}

function Inspect-Shape($Shape, [bool]$Recurse = $true) {
    $kind = Read-Property $Shape 'Kind'
    $container = Test-Container $Shape
    if ($Shape.OneD) { $kind = 'edge' }
    elseif ($container) { $kind = 'container' }
    $record = [ordered]@{
        visioId=[int]$Shape.ID; id=Read-Property $Shape 'AvId'; nameU=[string]$Shape.NameU
        component=Read-Property $Shape 'ComponentId'; kind=$kind; label=[string]$Shape.Text
        x=$Shape.CellsU('PinX').ResultIU; y=$Shape.CellsU('PinY').ResultIU
        width=$Shape.CellsU('Width').ResultIU; height=$Shape.CellsU('Height').ResultIU
        angle=$Shape.CellsU('Angle').ResultIU; parent=Read-Property $Shape 'ParentId'
        state=Read-Property $Shape 'State'; purpose=Read-Property $Shape 'Purpose'
        sourceRef=Read-Property $Shape 'SourceRef'; sourceId=Read-Property $Shape 'SourceIdRef'; url=Read-Property $Shape 'Source'
        cardStyle=Read-Property $Shape 'CardStyle'; baseFontSize=Read-Property $Shape 'BaseFontSize'
        icon=Read-Property $Shape 'Icon'; iconRef=Read-Property $Shape 'IconRef'
        details=Read-Property $Shape 'Details'; containerStyle=Read-Property $Shape 'ContainerStyle'; boundaryType=Read-Property $Shape 'BoundaryType'
        boundaryFill=''; boundaryColor=''
        furniture=Read-Property $Shape 'AvFurniture'; isContainer=[bool]$container
        textHidden=($Shape.CellExistsU('HideText',0) -and $Shape.CellsU('HideText').ResultIU -ne 0)
        internalShapeCount=$Shape.Shapes.Count; requirementIds=@(); confidence=$null
    }
    if ($record.containerStyle -eq 'boundary') {
        $record.boundaryColor=$Shape.CellsU('LineColor').FormulaU
        if ($Shape.CellsU('FillPattern').ResultIU -eq 1) { $record.boundaryFill=$Shape.CellsU('FillForegnd').FormulaU }
    }
    if (-not $Shape.OneD) {
        try { $record.serviceBounds=Get-NativeBounds $Shape } catch { $record.serviceBounds=$null }
    }
    if ($record.cardStyle -in @('icon','label')) {
        # Shape.Text can proxy a group's caption; Characters addresses its own text.
        $record.label=[string]$Shape.Characters.Text
        $warnings=[System.Collections.Generic.List[string]]::new()
        $caption=Get-RoleShape $Shape 'caption' -Optional
        $record.displayLabel=$(if ($null -ne $caption) { [string]$caption.Text } else { $null })
        $record.fullLabel=Read-Property $Shape 'FullLabel'
        $record.canonicalLabelSource='root.Characters.Text'; $record.displayLabelSource="child Prop.AvRole=caption"
        $record.canonicalTextHidden=($Shape.CellsU('HideText').ResultIU -ne 0)
        if ($null -eq $caption) { $warnings.Add('Native caption child missing or ambiguous; export/update is blocked until repaired.') }
        if (-not $record.canonicalTextHidden) { $warnings.Add('Canonical group text is visible; restore HideText=1 for icon-first presentation.') }
        if ($record.fullLabel -cne $record.label) { $warnings.Add('FullLabel metadata differs from actual canonical root.Characters.Text; actual group text takes precedence.') }
        if ($record.cardStyle -eq 'icon') {
            $record.iconSize=[double](Read-Property $Shape 'IconSize')
            $icon=Get-RoleShape $Shape 'icon' -Optional
            if ($null -eq $icon) { $warnings.Add('Native icon child missing or ambiguous; original catalog/master artwork, not manual edits, will be used by a clone.') }
            else {
                $record.actualIconWidth=$icon.CellsU('Width').ResultIU; $record.actualIconHeight=$icon.CellsU('Height').ResultIU
                $aspect=[double](Read-Property $Shape 'IconAspect')
                if ($record.actualIconHeight -le 0 -or [Math]::Abs($record.actualIconWidth/$record.actualIconHeight-$aspect) -gt 0.001 -or
                    [Math]::Abs([Math]::Max($record.actualIconWidth,$record.actualIconHeight)-$record.iconSize) -gt 0.001) {
                    $warnings.Add('Unsupported manual icon resize/aspect edit detected; clone restores original artwork at stored iconSize.')
                }
            }
            $warnings.Add('Icon internal artwork edits cannot be fully detected or reproduced; clone uses the original icon/iconRef, never internal primitives as nodes.')
        }
        $record.presentationWarnings=[string[]]$warnings.ToArray()
        if ($null -ne $caption) {
            try { $record.captionBounds=Get-NativeCaptionBounds $caption } catch { $record.captionBounds=$null }
            if ($record.cardStyle -eq 'label') { $record.serviceBounds=$record.captionBounds }
        }
        if ($record.cardStyle -eq 'icon') {
            $glyph=Get-RoleShape $Shape 'icon' -Optional
            try { $record.serviceBounds=Get-NativeBounds $glyph } catch { $record.serviceBounds=$null }
        }
    }
    $requirements = Read-Property $Shape 'RequirementIds'
    if ($requirements) {
        $parsedRequirements = ConvertFrom-Json -InputObject $requirements
        if ($parsedRequirements -isnot [array] -or
            @($parsedRequirements | Where-Object { $_ -isnot [string] }).Count) {
            throw "Invalid RequirementIds Shape Data on shape $($Shape.ID): expected a JSON string array."
        }
        $record.requirementIds = [string[]]$parsedRequirements
    }
    $confidence = Read-Property $Shape 'Confidence'
    if ($confidence) { $record.confidence = ConvertFrom-Json -InputObject $confidence }
    foreach ($pair in @(@('mainNodeIds','MainNodeIds'),@('mainEdgeIds','MainEdgeIds'))) {
        if (-not $Shape.CellExistsU('Prop.'+$pair[1],0)) { continue }
        $raw=Read-Property $Shape $pair[1]
        $parsed=ConvertFrom-Json -InputObject $raw
        if ($parsed -isnot [array] -or @($parsed | Where-Object { $_ -isnot [string] -or [string]::IsNullOrWhiteSpace($_) }).Count) {
            throw "Invalid $($pair[1]) Shape Data on shape $($Shape.ID): expected a JSON string array."
        }
        $record[$pair[0]]=$parsed
    }
    $layers = @()
    for ($i = 1; $i -le $Shape.LayerCount; $i++) { $layers += $Shape.Layer($i).Name }
    $record.layers = $layers
    if (-not $record.id -and -not $container -and $Shape.Shapes.Count) {
        $record.descendantLabels = @(Get-NativeShapes $Shape.Shapes |
            ForEach-Object { ([string]$_.shape.Text).Trim() } | Where-Object { $_ } | Select-Object -Unique)
    }
    if ($Shape.OneD) {
        $record.routeGeometry=Get-NativeRouteGeometry $Shape
        $record.relationship = Read-Property $Shape 'Relationship'
        $record.source = Read-Property $Shape 'SourceId'
        $record.target = Read-Property $Shape 'TargetId'
        $record.connections = $Shape.Connects.Count
        $glue = @()
        for ($i = 1; $i -le $Shape.Connects.Count; $i++) {
            $connection = $Shape.Connects.Item($i)
            $glue += [pscustomobject]@{
                fromCell=$connection.FromCell.Name; toCell=$connection.ToCell.Name
                toVisioId=[int]$connection.ToSheet.ID; toId=Read-Property $connection.ToSheet 'AvId'
            }
        }
        $record.glue = $glue
        $record.sourceVisioIds = @($glue | Where-Object { $_.fromCell -in @('BeginX','BeginY') } | Select-Object -ExpandProperty toVisioId -Unique)
        $record.targetVisioIds = @($glue | Where-Object { $_.fromCell -in @('EndX','EndY') } | Select-Object -ExpandProperty toVisioId -Unique)
        $record.beginX=$Shape.CellsU('BeginX').ResultIU; $record.beginY=$Shape.CellsU('BeginY').ResultIU
        $record.endX=$Shape.CellsU('EndX').ResultIU; $record.endY=$Shape.CellsU('EndY').ResultIU
        $begin = $Shape.CellsU('BeginArrow').ResultIU -ne 0
        $end = $Shape.CellsU('EndArrow').ResultIU -ne 0
        $record.direction = $(if ($begin -and $end) { 'both' } elseif ($begin) { 'backward' } elseif ($end) { 'forward' } else { 'none' })
        $record.linePattern = [int]$Shape.CellsU('LinePattern').ResultIU
        $record.color = $Shape.CellsU('LineColor').FormulaU
        $record.dashed = $record.linePattern -gt 1
        $record.routeStyle=''
        if ($Shape.CellExistsU('ShapeRouteStyle',0)) {
            $routing=[int]$Shape.CellsU('ShapeRouteStyle').ResultIU
            if ($routing -in @(1,2)) { $record.routeStyle=$(if ($routing -eq 2) { 'straight' } else { 'orthogonal' }) }
        }
        foreach ($endpoint in @('source','target')) {
            $prefix = $(if ($endpoint -eq 'source') { 'Source' } else { 'Target' })
            $side=Read-Property $Shape ($prefix+'Side')
            $position=Read-Property $Shape ($prefix+'Position')
            if ($side -and $position) {
                $record[$endpoint+'Side']=$side
                $record[$endpoint+'Position']=[double]$position
            }
        }
    }
    if ($container) { $record.memberIds = @($Shape.ContainerProperties.GetMemberShapes(0)) }
    if ($Recurse -and -not $record.id -and -not $container -and $Shape.Shapes.Count) {
        $children = @()
        for ($i = 1; $i -le $Shape.Shapes.Count; $i++) { $children += Inspect-Shape $Shape.Shapes.Item($i) $true }
        $record.children = $children
    }
    return [pscustomobject]$record
}

function Inspect-Document($Document) {
    $pages = @()
    for ($p = 1; $p -le $Document.Pages.Count; $p++) {
        $page = $Document.Pages.Item($p)
        $shapes = @()
        for ($i = 1; $i -le $page.Shapes.Count; $i++) {
            $shapes += Inspect-Shape $page.Shapes.Item($i)
        }
        $layers = @()
        for ($i = 1; $i -le $page.Layers.Count; $i++) { $layers += $page.Layers.Item($i).Name }
        $pageRecord=[pscustomobject]@{
            pageId=[int]$page.ID; name=$page.Name; layers=$layers; shapes=$shapes
            width=$page.PageSheet.CellsU('PageWidth').ResultIU; height=$page.PageSheet.CellsU('PageHeight').ResultIU
            furniture=Read-Property $page.PageSheet 'AvFurniture'
            role=Read-Property $page.PageSheet 'AvPageRole'
            view=Read-Property $page.PageSheet 'AvPageView'
        }
        $pageRecord | Add-Member -NotePropertyName routing -NotePropertyValue (Get-NativeRoutingReport $pageRecord)
        $pages+=$pageRecord
    }
    $snapshot=[pscustomobject]@{
        document=$Document.FullName; saved=$Document.Saved; pages=$pages
        referenceContract=Read-Property $Document.DocumentSheet 'ReferenceContract'
        conversionMode=Read-Property $Document.DocumentSheet 'ConversionMode'
        presentationProfile=Read-Property $Document.DocumentSheet 'PresentationProfile'
        portableRenderer=Read-Property $Document.DocumentSheet 'PortableRenderer'
        routingPolicy=Read-Property $Document.DocumentSheet 'RoutingPolicy'
    }
    if ($Document.DocumentSheet.CellExistsU('Prop.OutputContract',0)) {
        $snapshot | Add-Member -NotePropertyName outputContract -NotePropertyValue (Read-Property $Document.DocumentSheet 'OutputContract')
    }
    if ($Document.DocumentSheet.CellExistsU('Prop.Hardening',0)) {
        $snapshot | Add-Member -NotePropertyName hardening -NotePropertyValue (ConvertFrom-Json -InputObject (Read-Property $Document.DocumentSheet 'Hardening'))
    }
    return $snapshot
}

function Export-Model($Document) {
    $snapshot = Inspect-Document $Document
    $renderer=Get-Value $snapshot 'portableRenderer' ''
    $portable=$renderer -ceq 'azure-visio-1.6'
    $warnings = [System.Collections.Generic.List[string]]::new()
    $pages = @()
    foreach ($page in $snapshot.pages) {
        if ((Get-Value $snapshot 'routingPolicy' '') -ceq 'orthogonal-v1.6') {
            $routing=Get-Value $page 'routing'
            if ($null -eq $routing -or -not $routing.valid) {
                $problems=foreach ($issue in (Get-Value $routing 'issues' @())) { "$($issue.edgeId) [$($issue.code)]: $($issue.message)" }
                $failure=[InvalidOperationException]::new("Native architecture routing validation failed on '$($page.name)': $($problems -join '; ')")
                $failure.Data['NativeRoutingSnapshot']=ConvertTo-Json -InputObject $page -Depth 40 -Compress
                throw $failure
            }
        }
        $ids = @{}; $used = @{}; $records=@{}; $semanticBoundaries=@{}
        foreach ($shape in $page.shapes) {
            $records[[int]$shape.visioId]=$shape
            if ($shape.id) {
                Confirm-Identifier $shape.id
                if ($used.ContainsKey($shape.id)) { throw "Duplicate AvId '$($shape.id)' on page '$($page.name)'." }
                $used[$shape.id] = $true
            }
            if (Test-PortableBoundary $shape $renderer) { $semanticBoundaries[$shape.id]=$shape }
        }
        if ($semanticBoundaries.Count) {
            $warnings.Add("Page '$($page.name)': portable semantic-boundary mode preserves boundary rectangles and validated ParentId metadata; no native ContainerProperties membership is claimed.")
        }
        $nodes = @(); $edges = @()
        $spec = [ordered]@{ name=$page.name; width=$page.width; height=$page.height; title=''; subtitle=''; furniture=($page.furniture -eq 'True'); footer=$null }
        if ($page.role) { $spec.role=$page.role }
        if (Get-Value $page 'view' '') { $spec.view=$page.view }
        foreach ($shape in $page.shapes) {
            if ($shape.furniture) {
                if ($shape.furniture -in @('title','subtitle','footer')) { $spec[$shape.furniture] = $shape.label }
                continue
            }
            if ($shape.nameU -eq 'av-background' -and -not $shape.id) {
                $warnings.Add("Page '$($page.name)': legacy background omitted; legacy untagged annotations become editable notes.")
                continue
            }
            $id = $shape.id
            if (-not $id) {
                $id = "native-p$($page.pageId)-s$($shape.visioId)"
                while ($used.ContainsKey($id)) { $id += '-n' }
                $used[$id] = $true
            }
            $ids[[int]$shape.visioId] = $id
        }
        foreach ($shape in $page.shapes) {
            if (-not $ids.ContainsKey([int]$shape.visioId)) { continue }
            $item = [ordered]@{
                id=$ids[[int]$shape.visioId]; label=$shape.label; state=$shape.state
                purpose=$shape.purpose; sourceRef=$shape.sourceRef; requirementIds=@($shape.requirementIds)
            }
            if ($shape.sourceId) { $item.sourceId=$shape.sourceId }
            if ($shape.details) { $item.details=$shape.details }
            foreach ($field in @('mainNodeIds','mainEdgeIds')) {
                if (Test-Field $shape $field) { $item[$field]=Get-Value $shape $field }
            }
            if (-not $item.label -and (Test-Field $shape 'descendantLabels') -and $shape.descendantLabels.Count) {
                $item.label = $shape.descendantLabels -join "`n"
                $warnings.Add("Shape '$($item.id)': visible child labels combined into a generic group label.")
            }
            if ($null -ne $shape.confidence) { $item.confidence = $shape.confidence }
            if ($shape.url) { $item.url=$shape.url }
            if ($shape.kind -eq 'edge') {
                $kind = $shape.relationship
                if (-not $script:edgeStyles.ContainsKey($kind)) {
                    $kind = 'association'
                    $warnings.Add("Page '$($page.name)', shape $($shape.visioId): relationship unknown; exported as association, not packet traffic.")
                }
                $item.kind=$kind; $item.source=$null; $item.target=$null
                $item.direction=$shape.direction; $item.dashed=$shape.dashed
                if ($shape.routeStyle) { $item.routeStyle=$shape.routeStyle }
                if ($shape.color -match '^RGB\(\s*\d{1,3}\s*,\s*\d{1,3}\s*,\s*\d{1,3}\s*\)$') { $item.color=$shape.color }
                $item.beginX=$shape.beginX; $item.beginY=$shape.beginY; $item.endX=$shape.endX; $item.endY=$shape.endY
                $item.nativeEndpoints = $shape.glue
                foreach ($field in @('sourceSide','targetSide','sourcePosition','targetPosition')) {
                    if (Test-Field $shape $field) { $item[$field]=$shape.$field }
                }
                foreach ($endpoint in @('source','target')) {
                    $targets = @($shape.($endpoint+'VisioIds'))
                    if ($targets.Count -eq 1 -and $ids.ContainsKey([int]$targets[0]) -and
                        @($page.shapes | Where-Object { $_.visioId -eq $targets[0] -and $_.kind -ne 'edge' }).Count -eq 1) {
                        $item[$endpoint] = $ids[[int]$targets[0]]
                    } else {
                        $warnings.Add("Page '$($page.name)', edge '$($item.id)': $endpoint is unattached, ambiguous, or glued to a subshape/edge; preserved as an unglued endpoint with nativeEndpoints evidence.")
                    }
                    $stored = $shape.$endpoint
                    if ($stored -and $stored -ne $item[$endpoint]) { $warnings.Add("Edge '$($item.id)': stored $endpoint '$stored' differs from actual native glue; live glue takes precedence.") }
                }
                $edges += [pscustomobject]$item
            } else {
                $item.kind = $(if ($shape.isContainer -or (Test-PortableBoundary $shape $renderer)) { 'container' } elseif ($shape.kind -eq 'note') { 'note' } else { 'card' })
                if ((Test-Field $snapshot 'outputContract') -and (Get-Value $shape 'textHidden' $false) -and
                    ($item.kind -eq 'note' -or ($item.kind -eq 'card' -and $shape.cardStyle -notin @('icon','label')))) {
                    throw "Cannot export architecture pack: '$($item.id)' text is hidden, not a visible explanation."
                }
                $item.x=$shape.x; $item.y=$shape.y; $item.width=$shape.width; $item.height=$shape.height
                if ($shape.component) { $item.component=$shape.component }
                if ($shape.layers.Count) { $item.layer=$shape.layers[0] }
                if ($shape.icon) { $item.icon=$shape.icon }
                if ($shape.iconRef) { $item.iconRef=$shape.iconRef }
                if ($shape.cardStyle) { $item.cardStyle=$shape.cardStyle }
                foreach ($field in @('containerStyle','boundaryType')) { if ($shape.$field) { $item[$field]=$shape.$field } }
                if ($shape.containerStyle -eq 'boundary') {
                    foreach ($pair in @(@('boundaryFill','fill'),@('boundaryColor','color'))) {
                        if ($shape.($pair[0]) -match '^RGB\(\s*\d{1,3}\s*,\s*\d{1,3}\s*,\s*\d{1,3}\s*\)$') { $item[$pair[1]]=$shape.($pair[0]) }
                    }
                }
                if ($shape.cardStyle -in @('icon','label')) {
                    if ($null -eq $shape.displayLabel) { throw "Cannot export '$($item.id)': native caption child missing or ambiguous." }
                    $item.displayLabel=$shape.displayLabel; $item.details=$shape.details
                    if ($shape.cardStyle -eq 'icon') { $item.iconSize=$shape.iconSize }
                    foreach ($warning in $shape.presentationWarnings) { $warnings.Add("Shape '$($item.id)': $warning") }
                }
                if ($shape.baseFontSize) { $item.fontSize=[double]$shape.baseFontSize }
                if (-not $shape.id -or ($shape.internalShapeCount -gt 0 -and -not $shape.icon -and -not $shape.iconRef -and -not $shape.isContainer -and $shape.cardStyle -ne 'label')) {
                    $warnings.Add("Page '$($page.name)', shape $($shape.visioId): native shape/group becomes a generic editable semantic $($item.kind); internal artwork and formatting are not reproduced.")
                }
                if ([Math]::Abs($shape.angle) -gt 0.0001) { $warnings.Add("Shape '$($item.id)': rotation is not reproduced.") }
                if ($item.width -le 0 -or $item.height -le 0) { throw "Cannot export zero-sized native shape $($shape.visioId) as an editable card." }
                $nodes += [pscustomobject]$item
            }
        }
        foreach ($node in $nodes) {
            $nativeId = @($ids.Keys | Where-Object { $ids[$_] -eq $node.id })[0]
            $parents = @($page.shapes | Where-Object { $_.isContainer -and $_.memberIds -contains $nativeId })
            # Keep only immediate containers when Visio returns recursive membership.
            $parents = @($parents | Where-Object {
                $candidate = $_
                @($parents | Where-Object { $_.visioId -ne $candidate.visioId -and $candidate.memberIds -contains $_.visioId }).Count -eq 0
            })
            if ($parents.Count -eq 1 -and $ids.ContainsKey([int]$parents[0].visioId)) {
                $parent = @($nodes | Where-Object { $_.id -eq $ids[[int]$parents[0].visioId] })[0]
                if ([Math]::Abs($node.x-$parent.x)+$node.width/2 -le $parent.width/2+0.001 -and
                    [Math]::Abs($node.y-$parent.y)+$node.height/2 -le $parent.height/2+0.001) {
                    $node | Add-Member -NotePropertyName parent -NotePropertyValue $parent.id
                    if ($portable -and $records[$nativeId].parent -and $records[$nativeId].parent -cne $parent.id) {
                        $warnings.Add("Node '$($node.id)': native container membership overrides stored ParentId '$($records[$nativeId].parent)'.")
                    }
                } else { $warnings.Add("Node '$($node.id)': native container membership exceeds container bounds; parent omitted rather than moving shapes.") }
            } elseif ($parents.Count -gt 1) { $warnings.Add("Node '$($node.id)': multiple native containers cannot be represented by one parent; membership omitted.") }
            elseif ($portable -and $records[$nativeId].parent) {
                $parentId=$records[$nativeId].parent
                Confirm-Identifier $parentId
                if (-not $records[$nativeId].id) { throw "Portable semantic parent on '$($node.id)' requires a managed AvId." }
                if ($semanticBoundaries.ContainsKey($parentId) -and $semanticBoundaries[$parentId].id -ceq $parentId) {
                    $parent=$semanticBoundaries[$parentId]
                    if ([Math]::Abs($node.x-$parent.x)+$node.width/2 -gt $parent.width/2+0.001 -or
                        [Math]::Abs($node.y-$parent.y)+$node.height/2 -gt $parent.height/2+0.001) {
                        throw "Portable semantic node '$($node.id)' is outside parent '$parentId'; refusing to drop or change ParentId."
                    }
                    $node | Add-Member -NotePropertyName parent -NotePropertyValue $parentId
                } elseif (@($page.shapes | Where-Object { $_.isContainer -and $_.id -ceq $parentId }).Count -eq 1) {
                    $warnings.Add("Node '$($node.id)': stored ParentId '$parentId' is not supported by native membership and was omitted; native containers remain authoritative.")
                } else { throw "Unknown or invalid portable semantic parent '$parentId' for '$($node.id)'." }
            }
        }
        # Export live coordinates, including deliberate off-page native content.
        if ($nodes.Count) {
            $minX = [Math]::Min(0, [double](($nodes | ForEach-Object { $_.x-$_.width/2 } | Measure-Object -Minimum).Minimum))
            $minY = [Math]::Min(0, [double](($nodes | ForEach-Object { $_.y-$_.height/2 } | Measure-Object -Minimum).Minimum))
            $maxX = [Math]::Max($page.width, [double](($nodes | ForEach-Object { $_.x+$_.width/2 } | Measure-Object -Maximum).Maximum))
            $maxY = [Math]::Max($page.height, [double](($nodes | ForEach-Object { $_.y+$_.height/2 } | Measure-Object -Maximum).Maximum))
            if ($minX -lt -0.001 -or $minY -lt -0.001 -or $maxX -gt $page.width+0.001 -or $maxY -gt $page.height+0.001) {
                $warnings.Add("Page '$($page.name)': off-page coordinates preserved using allowOffPage; no layout or page-size changes were made.")
                $spec.allowOffPage=$true
            }
        }
        $spec.nodes=$nodes; $spec.edges=$edges; $pages += [pscustomobject]$spec
    }
    $warnings.Add('Semantic roundtrip only: custom formatting, routing, connection-point side, layers beyond the first, native artwork, and unsupported Shape Data are not fully reproduced.')
    $exported=[ordered]@{schemaVersion=1;title=[string]$Document.Title;pages=$pages;warnings=@($warnings);sourceDocument=$Document.FullName}
    if ($snapshot.referenceContract) {
        $exported.referenceContract=$snapshot.referenceContract
        $exported.conversionMode=$snapshot.conversionMode
    }
    if ($snapshot.presentationProfile) { $exported.presentationProfile=$snapshot.presentationProfile }
    foreach ($field in @('outputContract','hardening')) {
        if (Test-Field $snapshot $field) { $exported[$field]=Get-Value $snapshot $field }
    }
    if ($portable -or (Test-Field $snapshot 'outputContract')) {
        Confirm-Model ([pscustomobject]$exported)
        if ((Get-Value $snapshot 'presentationProfile') -ceq 'enterprise') {
            [void](& $script:enterpriseStylePath -InputModel ([pscustomobject]$exported))
        }
    }
    return [pscustomobject]$exported
}

function Confirm-TargetRef($Reference, [bool]$RequirePage = $true) {
    $allowed = @('shapeId','visioId')
    if ($RequirePage) { $allowed += 'pageId' }
    Confirm-Fields $Reference $allowed 'target'
    if ($RequirePage) {
        Confirm-Number (Get-Value $Reference 'pageId') 'pageId' 0 2147483647
        if ([Math]::Truncate($Reference.pageId) -ne $Reference.pageId) { throw 'pageId must be an integer.' }
    }
    if ((Test-Field $Reference 'shapeId') -eq (Test-Field $Reference 'visioId')) { throw 'Each target requires exactly one shapeId (AvId) or visioId.' }
    if (Test-Field $Reference 'shapeId') { Confirm-Identifier $Reference.shapeId }
    else {
        Confirm-Number $Reference.visioId 'visioId' 1 2147483647
        if ([Math]::Truncate($Reference.visioId) -ne $Reference.visioId) { throw 'visioId must be an integer.' }
    }
}

function Confirm-Changes($Changes, [string]$Operation) {
    if ($Operation -eq 'Update') { Confirm-Fields $Changes @('schemaVersion','updates') 'changes' }
    else { Confirm-Fields $Changes @('schemaVersion','targets','approvedTargets','previewToken') 'changes' }
    Confirm-Number (Get-Value $Changes 'schemaVersion') 'schemaVersion' 1 1
    $listName = $(if ($Operation -eq 'Update') { 'updates' } else { 'targets' })
    $list = Get-Value $Changes $listName
    if ($list -isnot [array] -or $list.Count -eq 0) { throw "$listName must be a nonempty array." }
    foreach ($entry in $list) {
        if ($Operation -eq 'Delete') { Confirm-TargetRef $entry; continue }
        Confirm-Fields $entry @('pageId','shapeId','visioId','set') 'update'
        $reference = [ordered]@{}
        foreach ($field in @('pageId','shapeId','visioId')) {
            if (Test-Field $entry $field) { $reference[$field] = $entry.$field }
        }
        Confirm-TargetRef ([pscustomobject]$reference)
        $set = Get-Value $entry 'set'
        Confirm-Fields $set @('label','displayLabel','details','x','y','width','height','state','purpose','sourceRef','requirementIds','mainNodeIds','mainEdgeIds','confidence','source','target','sourceSide','targetSide','sourcePosition','targetPosition','relationship','direction','dashed','routeStyle') 'set'
        if (@($set.PSObject.Properties).Count -eq 0) { throw 'set must contain at least one supported change.' }
        Confirm-Metadata $set
        foreach ($field in @('x','y','width','height')) {
            if (Test-Field $set $field) {
                Confirm-Number $set.$field $field
                if ($field -in @('width','height') -and $set.$field -le 0) { throw "$field must be positive." }
            }
        }
        foreach ($endpoint in @('source','target')) {
            if (Test-Field $set $endpoint) { Confirm-TargetRef $set.$endpoint $false }
            $side = $endpoint+'Side'
            if ((Test-Field $set $side) -and $set.$side -cnotin @('left','right','top','bottom')) { throw "Invalid $side." }
        }
        if ((Test-Field $set 'relationship') -and
            ($set.relationship -isnot [string] -or -not $script:edgeStyles.ContainsKey($set.relationship))) { throw 'Unsupported relationship.' }
        Confirm-EdgeSemantics $set
    }
    if ($Operation -eq 'Delete') {
        if (Test-Field $Changes 'approvedTargets') {
            if ($Changes.approvedTargets -isnot [array] -or $Changes.approvedTargets.Count -eq 0) { throw 'approvedTargets must be a nonempty array.' }
            foreach ($reference in $Changes.approvedTargets) { Confirm-TargetRef $reference }
        }
        if ((Test-Field $Changes 'previewToken') -and $Changes.previewToken -notmatch '^[a-f0-9]{64}$') { throw 'Invalid previewToken.' }
    }
}

function Resolve-Target($Document, $Reference, $Page = $null) {
    if ($null -eq $Page) {
        for ($i = 1; $i -le $Document.Pages.Count; $i++) {
            if ($Document.Pages.Item($i).ID -eq $Reference.pageId) { $Page = $Document.Pages.Item($i); break }
        }
        if ($null -eq $Page) { throw "Page id $($Reference.pageId) not found." }
    }
    $matches = @()
    for ($i = 1; $i -le $Page.Shapes.Count; $i++) {
        $shape = $Page.Shapes.Item($i)
        if (((Test-Field $Reference 'visioId') -and $shape.ID -eq $Reference.visioId) -or
            ((Test-Field $Reference 'shapeId') -and (Read-Property $shape 'AvId') -eq $Reference.shapeId)) { $matches += $shape }
    }
    if ($matches.Count -ne 1) {
        throw "Target $(ConvertTo-Json -InputObject $Reference -Compress) resolved to $($matches.Count) full shapes. Use an exact pageId and top-level AvId/Visio ID; subshape edits are not supported."
    }
    return [pscustomobject]@{ page=$Page; shape=$matches[0]; key="$($Page.ID):$($matches[0].ID)" }
}

function Confirm-Unlocked($Shape, [string[]]$Cells) {
    foreach ($cell in $Cells) {
        if ($Shape.CellExistsU($cell,0) -and $Shape.CellsU($cell).ResultIU -ne 0) { throw "Shape $($Shape.ID) is protected by $cell." }
    }
}

function Get-UpdatePlan($Document, $Changes) {
    $plans = @(); $seen = @{}
    $portableModel=$null
    $sheet=Get-Value $Document 'DocumentSheet'
    if ($null -ne $sheet -and (Read-Property $sheet 'PortableRenderer') -ceq 'azure-visio-1.6') {
        $portableModel=Export-Model $Document
    }
    foreach ($entry in $Changes.updates) {
        $plan = Resolve-Target $Document $entry
        if ($seen.ContainsKey($plan.key)) { throw "Duplicate update target $($plan.key)." }
        $seen[$plan.key]=$plan
        $shape=$plan.shape; $set=$entry.set
        $style=Read-Property $shape 'CardStyle'
        if ((Test-Field $set 'mainNodeIds') -or (Test-Field $set 'mainEdgeIds')) {
            if ((Read-Property $plan.page.PageSheet 'AvPageView') -cne 'flowchart' -or (Read-Property $shape 'Kind') -cne 'card') {
                throw 'MAIN coverage updates require a native flowchart card.'
            }
        }
        if (((Test-Field $set 'displayLabel') -or (Test-Field $set 'details')) -and $style -notin @('icon','label')) { throw 'displayLabel/details updates require a managed icon or label card.' }
        if ($style -eq 'icon' -and ((Test-Field $set 'width') -or (Test-Field $set 'height'))) {
            throw 'Width/height updates on icon groups are blocked to preserve brand aspect ratio. Use x/y to move; ExportModel, edit node geometry/iconSize, and New to create a resized copy.'
        }
        if ($style -in @('icon','label') -and ((Test-Field $set 'label') -or (Test-Field $set 'displayLabel') -or (Test-Field $set 'width') -or (Test-Field $set 'height'))) {
            $caption=Get-RoleShape $shape 'caption'
            Confirm-Unlocked $caption @('LockTextEdit')
            $layoutSpec=[pscustomobject]@{
                label=Get-Value $set 'label' ([string]$shape.Characters.Text)
                displayLabel=Get-Value $set 'displayLabel' $(if (Test-Field $set 'label') { Get-DisplayLabel $set } else { [string]$caption.Text })
                width=Get-Value $set 'width' $shape.CellsU('Width').ResultIU
                height=Get-Value $set 'height' $shape.CellsU('Height').ResultIU
                fontSize=[double](Read-Property $shape 'BaseFontSize')
            }
            if ($style -eq 'icon') { $layoutSpec | Add-Member -NotePropertyName iconSize -NotePropertyValue ([double](Read-Property $shape 'IconSize')) }
            [void](Get-CaptionLayout $layoutSpec $style)
        }
        $plan | Add-Member -NotePropertyName set -NotePropertyValue $set
        $bounds = [ordered]@{}
        foreach ($field in @('x','y','width','height')) {
            $cell = @{x='PinX';y='PinY';width='Width';height='Height'}[$field]
            $bounds[$field] = Get-Value $set $field $shape.CellsU($cell).ResultIU
        }
        $plan | Add-Member -NotePropertyName bounds -NotePropertyValue ([pscustomobject]$bounds)
        $geometry = @($set.PSObject.Properties.Name | Where-Object { $_ -in @('x','y','width','height') }).Count -gt 0
        $plan | Add-Member -NotePropertyName geometry -NotePropertyValue $geometry
        if (Test-Field $set 'label') { Confirm-Unlocked $shape @('LockTextEdit') }
        if ($geometry) {
            if ($shape.OneD) { throw 'Connector geometry is controlled by explicit source/target edits, not node bounds.' }
            if ($null -ne $portableModel) {
                $semanticPage=@($portableModel.pages | Where-Object { $_.name -ceq $plan.page.Name })[0]
                $semanticId=Read-Property $shape 'AvId'
                if ($semanticId) {
                    if ((Test-PortableBoundary (Inspect-Shape $shape $false) 'azure-visio-1.6') -and
                        @($semanticPage.nodes | Where-Object { (Get-Value $_ 'parent' '') -ceq $semanticId }).Count) {
                        throw 'Geometry edits of nonempty portable semantic boundaries are blocked to preserve child ownership and bounds.'
                    }
                    $semanticNode=@($semanticPage.nodes | Where-Object { $_.id -ceq $semanticId })[0]
                    $parentId=Get-Value $semanticNode 'parent' ''
                    if ($parentId) {
                        $parent=@($semanticPage.nodes | Where-Object { $_.id -ceq $parentId })[0]
                        if ([Math]::Abs($bounds.x-$parent.x)+$bounds.width/2 -gt $parent.width/2+0.001 -or
                            [Math]::Abs($bounds.y-$parent.y)+$bounds.height/2 -gt $parent.height/2+0.001) {
                            throw "Updated shape $($plan.key) would leave its declared parent; no implicit semantic reparenting is allowed."
                        }
                    }
                }
            }
            if ([Math]::Abs($shape.CellsU('Angle').ResultIU) -gt 0.0001 -or
                [Math]::Abs($shape.CellsU('LocPinX').ResultIU-$shape.CellsU('Width').ResultIU/2) -gt 0.001 -or
                [Math]::Abs($shape.CellsU('LocPinY').ResultIU-$shape.CellsU('Height').ResultIU/2) -gt 0.001) {
                throw 'Geometry edits require an unrotated, center-pinned full shape.'
            }
            if ((Test-Container $shape) -and @($shape.ContainerProperties.GetMemberShapes(0)).Count) {
                throw 'Geometry edits of nonempty containers are blocked to avoid implicit member layout.'
            }
            $locks = @()
            foreach ($field in @('x','y','width','height')) {
                if (Test-Field $set $field) { $locks += @{x='LockMoveX';y='LockMoveY';width='LockWidth';height='LockHeight'}[$field] }
            }
            Confirm-Unlocked $shape $locks
            if ($bounds.x-$bounds.width/2 -lt -0.001 -or $bounds.y-$bounds.height/2 -lt -0.001 -or
                $bounds.x+$bounds.width/2 -gt $plan.page.PageSheet.CellsU('PageWidth').ResultIU+0.001 -or
                $bounds.y+$bounds.height/2 -gt $plan.page.PageSheet.CellsU('PageHeight').ResultIU+0.001) { throw "Updated shape $($plan.key) would be outside its page." }
        }
        foreach ($field in @('source','target','sourceSide','targetSide','sourcePosition','targetPosition','relationship','direction','dashed','routeStyle')) {
            if ((Test-Field $set $field) -and -not $shape.OneD) { throw "$field is only valid for native 1-D connectors." }
        }
        foreach ($endpoint in @('source','target')) {
            $side = $endpoint+'Side'
            if (-not (Test-Field $set $endpoint) -and -not (Test-Field $set $side) -and -not (Test-Field $set ($endpoint+'Position'))) { continue }
            Confirm-Unlocked $shape @('LockBegin','LockEnd')
            if (Test-Field $set $endpoint) { $target = Resolve-Target $Document $set.$endpoint $plan.page }
            else {
                $native = Inspect-Shape $shape $false
                $glued = @($native.($endpoint+'VisioIds'))
                if ($glued.Count -ne 1) { throw "Changing $side requires one known native endpoint or an explicit $endpoint." }
                $target = Resolve-Target $Document ([pscustomobject]@{visioId=$glued[0]}) $plan.page
            }
            if ($target.shape.OneD) { throw 'Connector endpoints must target 2-D full shapes.' }
            $plan | Add-Member -NotePropertyName $endpoint -NotePropertyValue $target.shape
        }
        $plans += $plan
    }
    foreach ($plan in $plans) {
        if (-not $plan.geometry) { continue }
        for ($i = 1; $i -le $plan.page.Shapes.Count; $i++) {
            $parent = $plan.page.Shapes.Item($i)
            if (-not (Test-Container $parent) -or @($parent.ContainerProperties.GetMemberShapes(0)) -notcontains $plan.shape.ID) { continue }
            $b=$plan.bounds
            if ([Math]::Abs($b.x-$parent.CellsU('PinX').ResultIU)+$b.width/2 -gt $parent.CellsU('Width').ResultIU/2+0.001 -or
                [Math]::Abs($b.y-$parent.CellsU('PinY').ResultIU)+$b.height/2 -gt $parent.CellsU('Height').ResultIU/2+0.001) {
                throw "Updated shape $($plan.key) would leave its native container."
            }
        }
    }
    return $plans
}

function Apply-Updates($Plans) {
    foreach ($plan in $Plans) {
        $shape=$plan.shape; $set=$plan.set
        Set-NodeLabel $shape $set
        foreach ($field in @('width','height','x','y')) {
            if (Test-Field $set $field) {
                $cell = @{x='PinX';y='PinY';width='Width';height='Height'}[$field]
                $shape.CellsU($cell).ResultIU = [double]$set.$field
            }
        }
        Set-Metadata $shape $set
        if (Test-Field $set 'relationship') { Set-Property $shape 'Relationship' $set.relationship }
        Set-EdgeSemantics $shape $set
        foreach ($endpoint in @('source','target')) {
            if (-not (Test-Field $plan $endpoint)) { continue }
            $target=$plan.$endpoint
            $side = Get-Value $set ($endpoint+'Side') $(if ($endpoint -eq 'source') { 'right' } else { 'left' })
            $position = [double](Get-Value $set ($endpoint+'Position') 0.5)
            $cell = $(if ($endpoint -eq 'source') { 'BeginX' } else { 'EndX' })
            Glue-End $shape.CellsU($cell) $target $side $position
            Set-Property $shape $(if ($endpoint -eq 'source') { 'SourceId' } else { 'TargetId' }) (Read-Property $target 'AvId')
            Set-Property $shape $(if ($endpoint -eq 'source') { 'SourceVisioId' } else { 'TargetVisioId' }) ([string]$target.ID)
            $prefix = $(if ($endpoint -eq 'source') { 'Source' } else { 'Target' })
            Set-Property $shape ($prefix+'Side') $side
            Set-Property $shape ($prefix+'Position') ([string]$position)
        }
    }
    foreach ($plan in $Plans) {
        foreach ($field in @('x','y','width','height')) {
            if (Test-Field $plan.set $field) {
                $cell = @{x='PinX';y='PinY';width='Width';height='Height'}[$field]
                if ([Math]::Abs($plan.shape.CellsU($cell).ResultIU-[double]$plan.set.$field) -gt 0.001) { throw "Visio did not apply $field on $($plan.key); rolling back." }
            }
        }
        foreach ($endpoint in @('source','target')) {
            if (Test-Field $plan $endpoint) {
                $actual=Inspect-Shape $plan.shape $false
                if (@($actual.($endpoint+'VisioIds')) -notcontains $plan.$endpoint.ID) { throw "Visio did not glue $endpoint on $($plan.key); rolling back." }
            }
        }
    }
}

function Get-CanonicalTargets($Resolved) {
    return @($Resolved | Sort-Object { [int]$_.page.ID }, { [int]$_.shape.ID } |
        ForEach-Object { [pscustomobject]@{pageId=[int]$_.page.ID;visioId=[int]$_.shape.ID} })
}

function Get-DeletePlan($Document, $Changes) {
    $requested = @{}; $required = @{}; $reasons = @{}
    foreach ($reference in $Changes.targets) {
        $target=Resolve-Target $Document $reference
        if ($requested.ContainsKey($target.key)) { throw "Duplicate delete target $($target.key)." }
        $requested[$target.key]=$target; $required[$target.key]=$target; $reasons[$target.key]=@('explicit target')
    }
    $expanded=$true
    while ($expanded) {
        $expanded=$false
        foreach ($target in @($required.Values)) {
            $native=@(Get-NativeShapes $target.page.Shapes)
            $owned=@($native | Where-Object { $_.rootId -eq $target.shape.ID } | ForEach-Object { [int]$_.shape.ID })
            $dependencies=@()
            if (Test-Container $target.shape) {
                foreach ($memberId in @($target.shape.ContainerProperties.GetMemberShapes(0))) {
                    $dependencies += [pscustomobject]@{id=[int]$memberId;reason="member of container $($target.shape.ID)"}
                }
            }
            for ($i=1; $i -le $target.page.Connects.Count; $i++) {
                $connection=$target.page.Connects.Item($i)
                if ($owned -contains $connection.ToSheet.ID) {
                    $from=@($native | Where-Object { $_.shape.ID -eq $connection.FromSheet.ID })[0]
                    if ($from.rootId -ne $target.shape.ID) {
                        $dependencies += [pscustomobject]@{id=[int]$from.rootId;reason="native connection to $($target.shape.ID)"}
                    }
                }
            }
            $avId=Read-Property $target.shape 'AvId'
            if ($avId) {
                foreach ($entry in @($native | Where-Object { $_.topLevel -and $_.shape.OneD })) {
                    if ((Read-Property $entry.shape 'SourceId') -eq $avId -or (Read-Property $entry.shape 'TargetId') -eq $avId) {
                        $dependencies += [pscustomobject]@{id=[int]$entry.shape.ID;reason="declared endpoint $avId"}
                    }
                }
            }
            foreach ($dependency in $dependencies) {
                $dependent=Resolve-Target $Document ([pscustomobject]@{visioId=$dependency.id}) $target.page
                if (-not $required.ContainsKey($dependent.key)) {
                    $required[$dependent.key]=$dependent; $reasons[$dependent.key]=@(); $expanded=$true
                }
                if ($reasons[$dependent.key] -notcontains $dependency.reason) { $reasons[$dependent.key]+=$dependency.reason }
            }
        }
    }
    $canonical = @(Get-CanonicalTargets @($required.Values))
    $snapshot = Inspect-Document $Document
    $token = Get-TextHash ((ConvertTo-Json -InputObject $snapshot -Depth 80 -Compress) + (ConvertTo-Json -InputObject $canonical -Compress))
    $details=@()
    foreach ($reference in $canonical) {
        $key="$($reference.pageId):$($reference.visioId)"
        $target=$required[$key]
        $internal=@(Get-NativeShapes $target.shape.Shapes | ForEach-Object { [int]$_.shape.ID })
        $details += [pscustomobject]@{
            pageId=$reference.pageId; pageName=$target.page.Name; visioId=$reference.visioId
            shapeId=Read-Property $target.shape 'AvId'; label=[string]$target.shape.Text
            requested=$requested.ContainsKey($key); reasons=$reasons[$key]; groupDescendantVisioIds=$internal
        }
    }
    return [pscustomobject]@{
        requested=$requested; required=$required; token=$token
        report=[pscustomobject]@{
            action='Delete'; preview=$true; applied=$false; document=$Document.FullName
            targets=$details; requiredTargets=$canonical
            missingTargets=@($canonical | Where-Object { -not $requested.ContainsKey("$($_.pageId):$($_.visioId)") })
            previewToken=$token
            instruction='No changes made. Obtain approval for every requiredTarget. Apply using targets AND approvedTargets equal to requiredTargets, this previewToken, and -ApplyDelete. No automatic cascade.'
        }
    }
}

function Confirm-DeleteApproval($Document, $Changes, $Plan) {
    if ($Plan.report.missingTargets.Count) { throw 'Delete blocked: connected shapes or container descendants are undeclared. Preview and explicitly include every required target.' }
    if (-not (Test-Field $Changes 'approvedTargets') -or -not (Test-Field $Changes 'previewToken')) {
        throw 'ApplyDelete requires approvedTargets and previewToken from a preview approved by the user.'
    }
    $approved=@{}
    foreach ($reference in $Changes.approvedTargets) {
        $target=Resolve-Target $Document $reference
        if ($approved.ContainsKey($target.key)) { throw 'Duplicate approved target.' }
        $approved[$target.key]=$target
    }
    if ($approved.Count -ne $Plan.required.Count -or @($approved.Keys | Where-Object { -not $Plan.required.ContainsKey($_) }).Count) {
        throw 'approvedTargets must match the complete deletion list exactly.'
    }
    if ($Changes.previewToken -cne $Plan.token) { throw 'Delete preview is stale or belongs to a different target list/drawing. Preview again and obtain renewed approval.' }
    foreach ($target in $Plan.required.Values) { Confirm-Unlocked $target.shape @('LockDelete') }
}

function Apply-Deletions($Plan) {
    # Delete connectors first. Empty containers last, using member-first recursion.
    $remaining=@{}
    foreach ($key in $Plan.required.Keys) { $remaining[$key]=$Plan.required[$key] }
    while ($remaining.Count) {
        $progress=$false
        foreach ($target in @($remaining.Values | Sort-Object { -[int][bool]$_.shape.OneD })) {
            if (Test-Container $target.shape) {
                if (@($target.shape.ContainerProperties.GetMemberShapes(0)).Count) { continue }
            }
            [void]$target.shape.Delete()
            $remaining.Remove($target.key); $progress=$true
        }
        if (-not $progress) { throw 'Delete blocked by remaining native container members; rolling back.' }
    }
}

function Confirm-UpdateIsolation($Before, $After, $Plans) {
    $changes=@{}
    foreach ($plan in $Plans) { $changes[$plan.key]=$plan }
    foreach ($page in $Before.pages) {
        $current=@($After.pages | Where-Object { $_.pageId -eq $page.pageId })[0]
        if ($current.shapes.Count -ne $page.shapes.Count) { throw 'Update changed the shape count; rolling back.' }
        foreach ($shape in $page.shapes) {
            $live=@($current.shapes | Where-Object { $_.visioId -eq $shape.visioId })
            if ($live.Count -ne 1) { throw 'Update replaced an existing shape; rolling back.' }
            $live=$live[0]; $key="$($page.pageId):$($shape.visioId)"
            foreach ($field in @('x','y','width','height','angle')) {
                if ($shape.kind -eq 'edge' -or ($changes.ContainsKey($key) -and (Test-Field $changes[$key].set $field))) { continue }
                if ([Math]::Abs($live.$field-$shape.$field) -gt 0.001) { throw "Update implicitly changed $field on $key; rolling back." }
            }
            if ($shape.isContainer -and (@($shape.memberIds | Sort-Object) -join ',') -ne (@($live.memberIds | Sort-Object) -join ',')) {
                throw "Update implicitly changed native membership on $key; rolling back."
            }
        }
    }
}

function Confirm-RequiredMasters($Model) {
    $required=@{}
    foreach ($page in $Model.pages) {
        foreach ($node in $page.nodes) {
            if ($node.kind -eq 'container') { $required['container']=$true }
            $icon=Get-Value $node 'icon' ''
            if ($icon) { $required[$icon]=$true }
        }
    }
    foreach ($key in $required.Keys) { [void](Get-Master $key) }
    return @($required.Keys)
}

function Export-Document($Document, [string]$Directory) {
    Confirm-NativeArchitecturePack $Document
    if (-not $Directory) { throw 'Export requires -OutputDirectory.' }
    $Directory = Get-AbsolutePath $Directory
    if (-not (Test-Path -LiteralPath $Directory)) { [void][IO.Directory]::CreateDirectory($Directory) }
    $paths = @((Join-Path $Directory ([IO.Path]::GetFileNameWithoutExtension($Document.Name)+'.pdf')))
    for ($i = 1; $i -le $Document.Pages.Count; $i++) {
        $safeName = $Document.Pages.Item($i).Name -replace '[<>:"/\\|?*]', '-'
        $paths += Join-Path $Directory ($safeName+'.png')
    }
    foreach ($path in $paths) {
        if ((Test-Path -LiteralPath $path) -and -not $OverwriteExports) { throw "Export exists: $path. Use a new folder or explicit -OverwriteExports." }
    }
    $Document.ExportAsFixedFormat(1, $paths[0], 0, 0)
    for ($i = 1; $i -le $Document.Pages.Count; $i++) { $Document.Pages.Item($i).Export($paths[$i]) }
    foreach ($path in $paths) {
        if (-not (Test-Path -LiteralPath $path) -or (Get-Item -LiteralPath $path).Length -eq 0) { throw "Export missing or empty: $path" }
    }
    return $paths
}

if ($IconDirectory) { $IconDirectory = Get-AbsolutePath $IconDirectory }
else { $IconDirectory = Join-Path $PSScriptRoot 'IconLibrary' }
if ($LegacyModel -and $Action -ne 'New') { throw '-LegacyModel is only valid with New.' }
if ($Action -eq 'Catalog') {
    $entries=@()
    if (Test-Path -LiteralPath (Join-Path $IconDirectory 'catalog.json') -PathType Leaf) {
        $entries=@((Get-IconCatalog).Values | Sort-Object id)
    }
    [pscustomobject]@{ builtinIcons=@($script:icons.Keys | Sort-Object); iconDirectory=$IconDirectory; catalog=$entries } | ConvertTo-Json -Depth 12
    return
}
$model = $null
$styleReport = $null
if ($Action -in @('New','Merge','Validate') -or ($Action -eq 'Check' -and $ModelPath)) {
    if (-not $ModelPath) { throw "$Action requires -ModelPath." }
    $model = Get-Content -LiteralPath (Get-AbsolutePath $ModelPath) -Raw | ConvertFrom-Json
    Confirm-Model $model
    if ($Action -eq 'New') { Confirm-ArchitecturePack $model -Required:(-not $LegacyModel) }
    $contract = $ReferencePath
    if (-not $contract -and (Test-Field $model 'referenceContract')) {
        $contract = $model.referenceContract
        if (-not [IO.Path]::IsPathRooted($contract)) {
            if ($contract -match '[:]|(^|[\\/])\.\.([\\/]|$)') { throw 'Relative referenceContract cannot traverse outside the model folder.' }
            $contract = Join-Path ([IO.Path]::GetDirectoryName((Get-AbsolutePath $ModelPath))) $contract
        }
    }
    if ((Get-Value $model 'conversionMode' '') -in @('faithful','reference-plus-proposal') -and -not $contract) {
        throw 'Reference conversion requires a persisted source contract. Supply -ReferencePath or referenceContract; do not substitute a requirements-only redesign.'
    }
    if ($contract) {
        $contract = Get-AbsolutePath $contract
        [void](& (Join-Path $PSScriptRoot 'Reference-Fidelity.ps1') -ModelPath (Get-AbsolutePath $ModelPath) -ReferencePath $contract)
    }
    if ((Get-Value $model 'presentationProfile' '') -eq 'enterprise' -or (Test-Field $model 'outputContract')) {
        $styleReport = & (Join-Path $PSScriptRoot 'Enterprise-Style.ps1') -ModelPath (Get-AbsolutePath $ModelPath) | ConvertFrom-Json
        if (-not (Get-Value $styleReport 'valid' $false)) { throw 'Enterprise style validation did not return valid:true.' }
    }
}
if ($ReferencePath -and $null -eq $model) { throw '-ReferencePath is only supported with model-based New, Merge, Validate, or Check.' }
if ($Action -eq 'Validate') {
    $warnings=@()
    foreach ($page in $model.pages) {
        foreach ($edge in $page.edges) {
            if ($null -eq $edge.source -or $null -eq $edge.target) { $warnings += "Page '$($page.name)', edge '$($edge.id)' has an explicitly unglued endpoint." }
        }
    }
    [pscustomobject]@{ valid=$true; schemaVersion=1; pages=$model.pages.Count; warnings=$warnings; comUsed=$false; styleReport=$styleReport } | ConvertTo-Json -Depth 12
    return
}
$changes=$null
if ($Action -in @('Update','Delete')) {
    if (-not $ChangesPath) { throw "$Action requires -ChangesPath." }
    $changes=Get-Content -LiteralPath (Get-AbsolutePath $ChangesPath) -Raw | ConvertFrom-Json
    Confirm-Changes $changes $Action
}
if ($ApplyDelete -and $Action -ne 'Delete') { throw '-ApplyDelete is only valid with Delete.' }
if ($Action -eq 'ExportModel' -and $ModelPath) {
    $ModelPath=Get-AbsolutePath $ModelPath
    if ([IO.Path]::GetExtension($ModelPath) -ine '.json') { throw 'ExportModel ModelPath must end in .json.' }
    if (Test-Path -LiteralPath $ModelPath) { throw "Refusing to overwrite $ModelPath." }
    if (-not (Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($ModelPath)) -PathType Container)) { throw 'ExportModel destination directory must already exist.' }
}
if ($StencilDirectory) {
    $StencilDirectory = Get-AbsolutePath $StencilDirectory
    if (-not (Test-Path -LiteralPath $StencilDirectory -PathType Container)) { throw "StencilDirectory not found: $StencilDirectory" }
}
if ($EnvironmentPath) {
    $EnvironmentPath = Get-AbsolutePath $EnvironmentPath
    $script:syncRoots = @(Read-Environment $EnvironmentPath)
} else {
    $defaultEnvironment = Join-Path $PSScriptRoot 'environment.json'
    if (Test-Path -LiteralPath $defaultEnvironment -PathType Leaf) { $script:syncRoots = @(Read-Environment $defaultEnvironment) }
}
if ($Action -ne 'Check') {
    $DocumentPath = Get-AbsolutePath $DocumentPath
    if ([IO.Path]::GetExtension($DocumentPath) -ine '.vsdx') { throw 'DocumentPath must end in .vsdx.' }
}
if ($Action -eq 'New' -and (Test-Path -LiteralPath $DocumentPath)) { throw "Refusing to overwrite $DocumentPath." }
if ($Action -eq 'Rename' -and (-not $ComponentId -or -not $Label)) { throw 'Rename requires -ComponentId and a nonempty -Label.' }
if ($Action -eq 'Highlight' -and -not $EdgeIds) { throw 'Highlight requires explicit -EdgeIds.' }
if ($Action -eq 'Template') {
    if (-not $TemplatePath) { throw 'Template requires -TemplatePath.' }
    $TemplatePath = Get-AbsolutePath $TemplatePath
    if ([IO.Path]::GetExtension($TemplatePath) -ine '.vstx' -or (Test-Path -LiteralPath $TemplatePath)) {
        throw 'Use a new .vstx destination. Existing templates are never overwritten.'
    }
}
$newDocument = $null
$newSaved = $false
try {
    $script:app = Connect-Visio ([bool]$LaunchVisio -and -not [bool]$NoLaunchVisio)
    if ($Action -eq 'Check') {
        if ($null -eq $model) {
            [void](Get-Master 'container')
            foreach ($icon in $script:icons.Keys) { [void](Get-Master $icon) }
            $masters=@($script:icons.Keys)
        } else { $masters=@(Confirm-RequiredMasters $model | Where-Object { $_ -ne 'container' }) }
        [pscustomobject]@{
            ready=$true
            version='1.6.0'
            visioVersion=$script:app.Version
            powerShellVersion=$PSVersionTable.PSVersion.ToString()
            azureMastersVerified=$masters.Count
            requiredOnly=($null -ne $model)
            iconRefsVerified=@(if ($null -ne $model) { $model.pages.nodes | ForEach-Object { Get-Value $_ 'iconRef' '' } | Where-Object { $_ } | Sort-Object -Unique })
            syncRootCount=$script:syncRoots.Count
            stencilFiles=@($script:stencilPaths.Values | Sort-Object -Unique)
        } | ConvertTo-Json -Depth 4
        return
    }
    $document = Get-OpenDocument $DocumentPath
    if ($Action -eq 'New') {
        if ($null -ne $document) { throw "Destination already open: $DocumentPath" }
        [void](Confirm-RequiredMasters $model)
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($DocumentPath))
        $document = $script:app.Documents.Add('')
        $newDocument = $document
        $scope = $script:app.BeginUndoScope('Azure Visio: create')
        $commit = $false
        try {
            Apply-Model $document $model $true
            $document.Title = Get-Value $model 'title' ''
            $document.Subject = 'Illustrative Azure architecture. Not a deployment or compliance assessment.'
            if ($contract) {
                Set-Property $document.DocumentSheet 'ReferenceContract' $contract
                Set-Property $document.DocumentSheet 'ConversionMode' (Get-Value $model 'conversionMode' 'faithful')
            }
            Confirm-NativeArchitecturePack $document
            $commit = $true
        } finally { $script:app.EndUndoScope($scope, $commit) }
        [void]$document.SaveAsEx($DocumentPath, 4)
        $newSaved = $true
    } else {
        if ($null -eq $document) {
            if (-not (Test-Path -LiteralPath $DocumentPath -PathType Leaf)) { throw "Drawing not found: $DocumentPath" }
            $document = $script:app.Documents.Open($DocumentPath)
        }
        if ($Action -in @('Update','Delete')) {
            if (-not $document.Saved) { throw 'Drawing has unsaved edits. Save it in Visio first or save a working copy before running this command.' }
            $before=Inspect-Document $document
            $beforeJson=ConvertTo-Json -InputObject $before.pages -Depth 80 -Compress
            $plans=@(); $deletePlan=$null
            if ($Action -eq 'Update') { $plans=@(Get-UpdatePlan $document $changes) }
            else {
                $deletePlan=Get-DeletePlan $document $changes
                if (-not $ApplyDelete) { $deletePlan.report | ConvertTo-Json -Depth 20; return }
                Confirm-DeleteApproval $document $changes $deletePlan
            }
            $scope=$script:app.BeginUndoScope("Azure Visio: $Action")
            $commit=$false
            try {
                if ((Read-Property $document.DocumentSheet 'PortableRenderer') -ceq 'azure-visio-1.6' -and
                    (Read-Property $document.DocumentSheet 'OutputContract') -ceq 'architecture-pack-v1.6') {
                    Set-Property $document.DocumentSheet 'RoutingPolicy' 'orthogonal-v1.6'
                }
                if ($Action -eq 'Update') {
                    Apply-Updates $plans
                    Complete-NativeRouting $document
                    Confirm-UpdateIsolation $before (Inspect-Document $document) $plans
                } else { Apply-Deletions $deletePlan }
                Confirm-NativeArchitecturePack $document
                [void]$document.Save()
                $commit=$true
            } finally {
                $script:app.EndUndoScope($scope,$commit)
                if (-not $commit) {
                    $restored=Inspect-Document $document
                    if ((ConvertTo-Json -InputObject $restored.pages -Depth 80 -Compress) -cne $beforeJson) {
                        throw 'Visio rollback did not restore the inspected state. The drawing was not intentionally saved; inspect it before proceeding.'
                    }
                }
            }
            if ($Action -eq 'Update') {
                [pscustomobject]@{action='Update';applied=$true;document=$document.FullName;updatedTargets=@(Get-CanonicalTargets $plans)} | ConvertTo-Json -Depth 8
            } else {
                [pscustomobject]@{action='Delete';preview=$false;applied=$true;document=$document.FullName;deletedTargets=$deletePlan.report.requiredTargets} | ConvertTo-Json -Depth 8
            }
            return
        }
        if ($Action -in @('Merge','Rename','Highlight')) {
            if (-not $document.Saved) { throw 'Drawing has unsaved edits. Save it in Visio first or save a working copy before running this command.' }
            $scope = $script:app.BeginUndoScope("Azure Visio: $Action")
            $commit = $false
            try {
                if ((Read-Property $document.DocumentSheet 'PortableRenderer') -ceq 'azure-visio-1.6' -and
                    (Read-Property $document.DocumentSheet 'OutputContract') -ceq 'architecture-pack-v1.6') {
                    Set-Property $document.DocumentSheet 'RoutingPolicy' 'orthogonal-v1.6'
                }
                if ($Action -eq 'Merge') { Apply-Model $document $model $false }
                if ($Action -eq 'Rename') {
                    $renameUpdates = @()
                    for ($p = 1; $p -le $document.Pages.Count; $p++) {
                        foreach ($shape in (Get-ShapeIndex $document.Pages.Item($p)).Values) {
                            if ((Read-Property $shape 'ComponentId') -eq $ComponentId) {
                                $renameUpdates += [pscustomobject]@{pageId=[int]$document.Pages.Item($p).ID;visioId=[int]$shape.ID;set=[pscustomobject]@{label=$Label}}
                            }
                        }
                    }
                    if ($renameUpdates.Count -eq 0) { throw "Component '$ComponentId' not found." }
                    Apply-Updates @(Get-UpdatePlan $document ([pscustomobject]@{updates=$renameUpdates}))
                }
                if ($Action -eq 'Highlight') {
                    $found = @{}
                    for ($p = 1; $p -le $document.Pages.Count; $p++) {
                        $index = Get-ShapeIndex $document.Pages.Item($p)
                        foreach ($edgeId in $EdgeIds) {
                            if ($index.ContainsKey($edgeId)) {
                                $shape = $index[$edgeId]
                                if ((Read-Property $shape 'Kind') -ne 'edge') { throw "'$edgeId' is not a connector." }
                                if ($ClearHighlight) {
                                    $previous = Read-Property $shape 'PreviousLineWeight'
                                    if (-not $previous) { throw "'$edgeId' has no saved highlight to clear." }
                                    $shape.CellsU('LineWeight').FormulaU = $previous
                                    Set-Property $shape 'PreviousLineWeight' ''
                                } else {
                                    if (-not (Read-Property $shape 'PreviousLineWeight')) {
                                        Set-Property $shape 'PreviousLineWeight' $shape.CellsU('LineWeight').FormulaU
                                    }
                                    $shape.CellsU('LineWeight').FormulaU = '4 pt'
                                }
                                $found[$edgeId] = $true
                            }
                        }
                    }
                    foreach ($edgeId in $EdgeIds) { if (-not $found.ContainsKey($edgeId)) { throw "Connector '$edgeId' not found." } }
                }
                if ($Action -eq 'Rename') { Complete-NativeRouting $document }
                Confirm-NativeArchitecturePack $document
                [void]$document.Save()
                $commit = $true
            } finally { $script:app.EndUndoScope($scope, $commit) }
        }
    }
    switch ($Action) {
        'Inspect' { Inspect-Document $document | ConvertTo-Json -Depth 80 }
        'ExportModel' {
            $result=Export-Model $document
            $json=ConvertTo-Json -InputObject $result -Depth 80
            if ($ModelPath) {
                $stream=[IO.File]::Open($ModelPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
                $writer=[IO.StreamWriter]::new($stream,[Text.UTF8Encoding]::new($false))
                try { $writer.Write($json) } finally { $writer.Dispose() }
                [pscustomobject]@{action='ExportModel';modelPath=$ModelPath;pages=$result.pages.Count;warnings=$result.warnings} | ConvertTo-Json -Depth 8
            } else { $json }
        }
        'Export' { Export-Document $document $OutputDirectory | ConvertTo-Json }
        'Template' {
            if (-not $document.Saved) { throw 'Save the drawing before creating a template.' }
            Confirm-NativeArchitecturePack $document
            $copy = $script:app.Documents.Add($DocumentPath)
            try { [void]$copy.SaveAs($TemplatePath) } finally { [void]$copy.Close() }
            [pscustomobject]@{ template = $TemplatePath } | ConvertTo-Json
        }
        default {
            $exports = @()
            if ($OutputDirectory) { $exports = @(Export-Document $document $OutputDirectory) }
            [pscustomobject]@{ action = $Action; document = $DocumentPath; visioLocation = $document.FullName; pages = $document.Pages.Count; exports = $exports } | ConvertTo-Json
        }
    }
} finally {
    if ($null -ne $newDocument -and -not $newSaved) {
        $newDocument.Saved = $true
        [void]$newDocument.Close()
    }
    foreach ($stencil in $script:ownedStencils) { [void]$stencil.Close() }
}

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SkillDirectory,
    [Parameter(Mandatory)][string]$OutputPath
)
Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
foreach ($path in @($SkillDirectory, $OutputPath)) {
    if ($path -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+(?:\\|$))') { throw "Use an absolute Windows path: $path" }
}
if ([IO.Path]::GetExtension($OutputPath) -ine '.zip' -or (Test-Path -LiteralPath $OutputPath)) {
    throw 'Provide a new .zip destination; packages are never overwritten.'
}
$files = [ordered]@{ 'SKILL.md' = (Join-Path $SkillDirectory 'SKILL.md') }
foreach ($name in @('AzureVisio.ps1', 'New-ReferenceModel.ps1', 'hub-spoke-reference.json',
                    'environment.example.json', 'Test-AzureVisio.ps1', 'Build-AzureVisioPackage.ps1', 'README.txt',
                    'Import-Draft.ps1', 'Draft-Import.Tests.ps1', 'Install-IconLibrary.ps1', 'icon-sources.json',
                    'Architecture-Guide.txt', 'architecture-references.json', 'CRUD-Guide.txt',
                    'Reference-Workflow.txt', 'Reference-Fidelity.ps1', 'Reference-Fidelity.Tests.ps1',
                    'Quiet-Workflow.txt', 'Enterprise-Style.ps1', 'Install-ReferenceAtlas.ps1')) {
    $files[$name] = Join-Path $PSScriptRoot $name
}
$total = 0
if ($files.Count -gt 21) { throw 'Skill exceeds SKILL.md plus 20 companion files.' }
foreach ($path in $files.Values) {
    $file = Get-Item -LiteralPath $path
    if ($file.PSIsContainer -or $file.Length -gt 5MB) { throw "Missing, invalid, or oversized companion file: $path" }
    $total += $file.Length
}
if ($total -gt 10MB) { throw 'Skill contents exceed the 10 MB companion-file limit.' }
$skill = Get-Content -LiteralPath $files['SKILL.md'] -Raw
if ($skill -notmatch '(?s)^---\r?\n.*?\bname:\s*"?azure-visio"?\r?\n.*?\bdescription:.*?\r?\n---') {
    throw 'SKILL.md must include name: azure-visio and description in YAML frontmatter.'
}
if ((Get-Item -LiteralPath $files['SKILL.md']).Length -gt 1MB) { throw 'SKILL.md exceeds the 1 MB upload limit.' }
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($OutputPath))
$stream = [IO.File]::Open($OutputPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
$complete = $false
try {
    $zip = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create, $true)
    try {
        foreach ($name in $files.Keys) {
            [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $files[$name], $name, [IO.Compression.CompressionLevel]::Optimal)
        }
    } finally { $zip.Dispose() }
    $complete = $true
} finally {
    $stream.Dispose()
    if (-not $complete) { Remove-Item -LiteralPath $OutputPath }
}
[pscustomobject]@{ package=$OutputPath; files=@($files.Keys); bytes=(Get-Item -LiteralPath $OutputPath).Length } | ConvertTo-Json

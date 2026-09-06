[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$OutputDirectory,
    [string[]]$Collections = @('azure','microsoft365','fabric','entra','power-platform','dynamics365'),
    [switch]$AcceptIconTerms
)
Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
if (-not $AcceptIconTerms) {
    throw 'Read the Microsoft icon-use terms on the source pages in icon-sources.json, then specify -AcceptIconTerms for architectural diagram, training, or documentation use only.'
}
if ($OutputDirectory -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+(?:\\|$))' -or
    (Test-Path -LiteralPath $OutputDirectory)) { throw 'Use a NEW absolute OutputDirectory. Existing libraries are never overwritten.' }
$sources = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'icon-sources.json') -Raw | ConvertFrom-Json
foreach ($collection in $Collections) {
    if ($collection -notin $sources.id) { throw "Unknown collection: $collection" }
}
if (@($Collections | Select-Object -Unique).Count -ne $Collections.Count -or $Collections.Count -eq 0) {
    throw 'Specify one or more distinct collections.'
}
Add-Type -AssemblyName System.Net.Http
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$allowedHosts = @('arch-center.azureedge.net','download.microsoft.com','go.microsoft.com','raw.githubusercontent.com',
    'github.com','objects.githubusercontent.com','release-assets.githubusercontent.com')
$handler = [Net.Http.HttpClientHandler]::new()
$handler.AllowAutoRedirect = $false
$client = [Net.Http.HttpClient]::new($handler)
$client.Timeout = [TimeSpan]::FromMinutes(3)
$client.DefaultRequestHeaders.UserAgent.ParseAdd('AzureVisio-OfficialIconInstaller/1.2')
$root = [IO.Path]::GetFullPath($OutputDirectory).TrimEnd('\')
[void][IO.Directory]::CreateDirectory($root)
$records = [Collections.Generic.List[object]]::new()
$downloads = [Collections.Generic.List[object]]::new()

function Download-OfficialZip([string]$Url, [string]$Destination) {
    $uri = [uri]$Url
    for ($redirect = 0; $redirect -le 5; $redirect++) {
        if ($uri.Scheme -ne 'https' -or $uri.Host -notin $allowedHosts -or $uri.UserInfo) {
            throw "Refusing non-allow-listed icon download URL: $uri"
        }
        $response = $client.GetAsync($uri, [Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
        try {
            if ([int]$response.StatusCode -in @(301,302,303,307,308)) {
                if (-not $response.Headers.Location) { throw 'Download redirect has no destination.' }
                $uri = [uri]::new($uri, $response.Headers.Location)
                continue
            }
            [void]$response.EnsureSuccessStatusCode()
            if ($response.Content.Headers.ContentLength -gt 100MB) { throw 'Icon archive exceeds 100 MB.' }
            $inputStream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
            $outputStream = [IO.File]::Open($Destination, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write)
            try {
                $buffer = New-Object byte[] 65536
                $total = 0
                while (($count = $inputStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    $total += $count
                    if ($total -gt 100MB) { throw 'Icon archive exceeds 100 MB.' }
                    $outputStream.Write($buffer, 0, $count)
                }
            } finally { $outputStream.Dispose(); $inputStream.Dispose() }
            return $uri.AbsoluteUri
        } finally { $response.Dispose() }
    }
    throw 'Icon download exceeded five redirects.'
}

function Get-SvgIssues([string]$Path) {
    $settings = [Xml.XmlReaderSettings]::new()
    $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $settings.MaxCharactersInDocument = 10MB
    $reader = [Xml.XmlReader]::Create($Path, $settings)
    try {
        $xml = [Xml.XmlDocument]::new()
        $xml.XmlResolver = $null
        $xml.Load($reader)
    } catch [Xml.XmlException] {
        return "SVG XML is not safely importable: $($_.Exception.Message)"
    } finally { $reader.Dispose() }
    if ($xml.DocumentElement.LocalName -ne 'svg') { return 'Root element is not SVG.' }
    foreach ($node in $xml.SelectNodes('//*')) {
        if ($node.LocalName -in @('script','foreignObject','iframe','image','animate','animateTransform','set')) {
            return "Unsupported active or embedded element: $($node.LocalName)"
        }
        foreach ($attribute in $node.Attributes) {
            if ($attribute.LocalName -match '^on' -or
                ($attribute.LocalName -in @('href','src') -and $attribute.Value -and -not $attribute.Value.StartsWith('#'))) {
                return 'SVG contains an event handler or external/embedded reference.'
            }
        }
    }
    $text = $xml.OuterXml
    if ($text -match '(?i)@import|url\s*\(\s*[''"]?(?!#)[^''")\s]') {
        return 'SVG contains an external CSS resource reference.'
    }
    return ''
}

try {
    foreach ($source in $sources | Where-Object { $_.id -in $Collections }) {
        $archivePath = Join-Path $root ($source.id + '.zip')
        $resolvedUrl = Download-OfficialZip $source.downloadUrl $archivePath
        $destination = Join-Path $root $source.id
        [void][IO.Directory]::CreateDirectory($destination)
        $zip = [IO.Compression.ZipFile]::OpenRead($archivePath)
        try {
            $expandedBytes = 0L
            if ($zip.Entries.Count -gt 10000) { throw "Too many entries in $($source.id)." }
            $paths = @{}
            foreach ($entry in $zip.Entries) {
                $expandedBytes += $entry.Length
                if ($expandedBytes -gt 500MB -or $entry.Length -gt 100MB) { throw 'Expanded icon archive exceeds size limits.' }
                $relative = $entry.FullName.Replace('/', '\')
                if ($relative -match '(^\\|:|(^|\\)\.\.(\\|$))') { throw "Unsafe archive entry: $relative" }
                $path = [IO.Path]::GetFullPath((Join-Path $destination $relative))
                if (-not $path.StartsWith($destination + '\', [StringComparison]::OrdinalIgnoreCase)) {
                    throw "Archive entry escapes its collection: $relative"
                }
                if ($paths.ContainsKey($path)) { throw "Duplicate archive destination: $relative" }
                $paths[$path] = $true
            }
            foreach ($entry in $zip.Entries) {
                $path = Join-Path $destination $entry.FullName.Replace('/', '\')
                if (-not $entry.Name) { [void][IO.Directory]::CreateDirectory($path); continue }
                [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
                [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $path, $false)
            }
        } finally { $zip.Dispose() }
        $svgCount = 0; $safeCount = 0
        foreach ($svg in Get-ChildItem -LiteralPath $destination -Recurse -File -Filter '*.svg' | Sort-Object FullName) {
            $svgCount++
            $relative = $svg.FullName.Substring($root.Length + 1)
            $hash = (Get-FileHash -LiteralPath $svg.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            $pathHash = [Security.Cryptography.SHA256]::Create()
            try {
                $digest = [BitConverter]::ToString($pathHash.ComputeHash([Text.Encoding]::UTF8.GetBytes($relative.ToLowerInvariant()))).Replace('-', '').Substring(0,12).ToLowerInvariant()
            } finally { $pathHash.Dispose() }
            $issue = Get-SvgIssues $svg.FullName
            if (-not $issue) { $safeCount++ }
            $records.Add([ordered]@{
                id=($source.id + '-' + $digest)
                path=$relative
                collection=$source.id
                name=$svg.BaseName
                sourcePage=$source.sourcePage
                downloadUrl=$source.downloadUrl
                sha256=$hash
                usable=(-not [bool]$issue)
                issue=$issue
            })
        }
        $downloads.Add([ordered]@{
            collection=$source.id; sourcePage=$source.sourcePage; downloadUrl=$source.downloadUrl
            resolvedUrl=$resolvedUrl; archive=([IO.Path]::GetFileName($archivePath))
            sha256=(Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash.ToLowerInvariant()
            svgCount=$svgCount; usableSvgCount=$safeCount; status=$source.status
        })
        Write-Host ("{0}: downloaded, {1} SVGs ({2} importable)" -f $source.id,$svgCount,$safeCount)
    }
    ConvertTo-Json -InputObject @($records.ToArray()) -Depth 5 |
        Set-Content -LiteralPath (Join-Path $root 'catalog.json') -Encoding UTF8
    [ordered]@{
        schemaVersion=1; downloadedUtc=[DateTime]::UtcNow.ToString('o'); sources=@($downloads.ToArray())
        terms='Microsoft icons: architectural diagrams, training materials, or documentation only. Preserve original artwork and product identity. See source pages and any included terms.'
        completeness='All files in the selected published packages, not a claim to contain every Microsoft product icon. Microsoft 365 source documentation is archived.'
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $root 'sources.json') -Encoding UTF8
    [pscustomobject]@{library=$root; collections=$downloads.Count; svgCount=$records.Count; importable=@($records | Where-Object { $_.usable }).Count} | ConvertTo-Json
} finally { $client.Dispose(); $handler.Dispose() }

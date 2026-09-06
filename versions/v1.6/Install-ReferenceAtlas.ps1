[CmdletBinding()]
param([Parameter(Mandatory)][string]$OutputDirectory)
Set-StrictMode -Version 2
$ErrorActionPreference='Stop'
if ($OutputDirectory -notmatch '^[A-Za-z]:\\' -or (Test-Path -LiteralPath $OutputDirectory)) {
    throw 'Use a NEW absolute local OutputDirectory for the official reference atlas.'
}
$catalog=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'architecture-references.json') -Raw | ConvertFrom-Json
if (-not $catalog.visualReferences.Count) { throw 'No visual references are configured.' }
$root=[IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($root)
Add-Type -AssemblyName System.Net.Http
$handler=[Net.Http.HttpClientHandler]::new()
$handler.AllowAutoRedirect=$false
$client=[Net.Http.HttpClient]::new($handler)
$client.Timeout=[TimeSpan]::FromMinutes(3)
$records=[Collections.Generic.List[object]]::new()

function Download-Reference([string]$Url,[string]$Path) {
    $uri=[uri]$Url
    for($redirect=0;$redirect -le 5;$redirect++){
        if($uri.Scheme -ne 'https' -or $uri.Host -notin @('learn.microsoft.com','arch-center.azureedge.net') -or $uri.UserInfo){
            throw "Reference URL is not an approved official source: $uri"
        }
        $response=$client.GetAsync($uri,[Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
        try{
            if([int]$response.StatusCode -in @(301,302,303,307,308)){
                if(-not $response.Headers.Location){throw 'Reference redirect has no target.'}
                $uri=[uri]::new($uri,$response.Headers.Location)
                continue
            }
            [void]$response.EnsureSuccessStatusCode()
            if($response.Content.Headers.ContentLength -gt 100MB){throw 'Reference download exceeds 100 MB.'}
            $inputStream=$response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
            $outputStream=[IO.File]::Open($Path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
            try{
                $buffer=New-Object byte[] 65536
                $size=0L
                while(($count=$inputStream.Read($buffer,0,$buffer.Length)) -gt 0){
                    $size+=$count
                    if($size -gt 100MB){throw 'Reference download exceeds 100 MB.'}
                    $outputStream.Write($buffer,0,$count)
                }
            }finally{$outputStream.Dispose();$inputStream.Dispose()}
            return $uri.AbsoluteUri
        }finally{$response.Dispose()}
    }
    throw 'Too many reference redirects.'
}

try{
    foreach($reference in $catalog.visualReferences){
        if($reference.id -notmatch '^[a-z0-9-]+$'){throw 'Unsafe reference ID.'}
        foreach($format in @('svg','vsdx')){
            $property=$reference.PSObject.Properties[$format]
            if($null -eq $property){continue}
            $name=$reference.id+'.'+$format
            $path=Join-Path $root $name
            $resolved=Download-Reference $property.Value $path
            if($format -eq 'svg'){
                $settings=[Xml.XmlReaderSettings]::new()
                $settings.DtdProcessing=[Xml.DtdProcessing]::Ignore
                $settings.XmlResolver=$null
                $settings.MaxCharactersInDocument=100MB
                $reader=[Xml.XmlReader]::Create($path,$settings)
                try{
                    $xml=[Xml.XmlDocument]::new();$xml.XmlResolver=$null;$xml.Load($reader)
                    if($xml.DocumentElement.LocalName -ne 'svg'){throw 'Reference response is not an SVG.'}
                }finally{$reader.Dispose()}
            }else{
                $stream=[IO.File]::OpenRead($path)
                try{$header=New-Object byte[] 4;$read=$stream.Read($header,0,4)}finally{$stream.Dispose()}
                if($read -ne 4 -or ($header -join ',') -ne '80,75,3,4'){throw 'Reference response is not an unencrypted VSDX package; no parsing or decryption attempted.'}
            }
            $records.Add([ordered]@{
                id=$reference.id;title=$reference.title;archetype=$reference.archetype
                article=$reference.article;format=$format;file=$name
                sourceUrl=$property.Value;resolvedUrl=$resolved
                sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
                bytes=(Get-Item -LiteralPath $path).Length
            })
        }
    }
    [ordered]@{
        schemaVersion=1;downloadedUtc=[DateTime]::UtcNow.ToString('o')
        use='Official reference originals for architecture study. Retain attribution; not model training and not redistributed in the skill package.'
        references=@($records.ToArray())
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $root 'atlas.json') -Encoding UTF8
    [pscustomobject]@{complete=$true;directory=$root;architectures=$catalog.visualReferences.Count;assets=$records.Count} | ConvertTo-Json
}finally{$client.Dispose();$handler.Dispose()}

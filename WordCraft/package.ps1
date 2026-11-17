$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$exe = Join-Path $PSScriptRoot 'openmd.exe'
$version = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe).FileVersion
if ($version -notmatch '^\d+\.\d+\.\d+\.\d+$') { throw 'Build a versioned openmd.exe first.' }
$version = $version -replace '\.0$', ''
$release = Join-Path $root '.tmp\releases'
$stage = Join-Path $root ('.tmp\release-stage-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $release, $stage -Force | Out-Null
Copy-Item -LiteralPath $exe -Destination (Join-Path $stage 'openmd.exe')
foreach ($name in @('README.md', 'LICENSE')) {
    Copy-Item -LiteralPath (Join-Path $root $name) -Destination (Join-Path $stage $name)
}
$archive = Join-Path $release "openmd-wordcraft-$version-windows.zip"
Compress-Archive -LiteralPath (Join-Path $stage 'openmd.exe'), (Join-Path $stage 'README.md'), (Join-Path $stage 'LICENSE') -DestinationPath $archive -Force
$hash = (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant()
[IO.File]::WriteAllText($archive + '.sha256', "$hash  $([IO.Path]::GetFileName($archive))`n", [Text.UTF8Encoding]::new($false))
Write-Host $archive
Write-Host "$hash (SHA-256)"

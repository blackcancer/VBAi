param(
    [switch]$Offline,
    [string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repository 'assets/editor/dist' }
if (-not [IO.Path]::IsPathRooted($OutputDirectory)) { $OutputDirectory = Join-Path $repository $OutputDirectory }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
Push-Location $repository
try {
    foreach ($name in @('monaco', 'esbuild')) {
        $lock = Get-Content "assets/editor/$name.lock.json" -Raw | ConvertFrom-Json
        $directory = Join-Path $repository "artifacts/monaco-dependencies/$name"
        $archive = Join-Path $directory 'package.tgz'
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
        if (-not (Test-Path -LiteralPath $archive)) {
            if ($Offline) { throw "Missing cached archive: $archive" }
            Invoke-WebRequest -Uri $lock.url -OutFile $archive -UseBasicParsing
        }
        $sha = [Security.Cryptography.SHA512]::Create()
        try { $actual = 'sha512-' + [Convert]::ToBase64String($sha.ComputeHash([IO.File]::ReadAllBytes($archive))) }
        finally { $sha.Dispose() }
        if ($actual -cne $lock.integrity) { throw "Package integrity mismatch: $name" }
        & tar -xzf $archive -C $directory
        if ($LASTEXITCODE -ne 0) { throw "Extraction failed: $name" }
    }
    $builder = 'artifacts/monaco-dependencies/esbuild/package/esbuild.exe'
    & $builder assets/editor/src/editor.js --bundle --minify --format=iife --loader:.ttf=file '--asset-names=[name]-[hash]' "--outdir=$OutputDirectory" --outbase=assets/editor/src
    if ($LASTEXITCODE -ne 0) { throw 'Editor bundle failed.' }
    & $builder artifacts/monaco-dependencies/monaco/package/esm/vs/editor/editor.worker.js --bundle --minify --format=esm "--outfile=$(Join-Path $OutputDirectory 'editor.worker.js')"
    if ($LASTEXITCODE -ne 0) { throw 'Worker bundle failed.' }
    Copy-Item assets/editor/src/index.html $OutputDirectory
    foreach ($language in @('fr', 'de', 'es', 'it', 'pt-br', 'ru', 'ja', 'ko', 'zh-cn', 'zh-tw')) {
        Copy-Item "artifacts/monaco-dependencies/monaco/package/esm/nls.messages.$language.js" $OutputDirectory
    }
    Copy-Item artifacts/monaco-dependencies/monaco/package/LICENSE (Join-Path $OutputDirectory 'MONACO-LICENSE.txt')
    Copy-Item artifacts/monaco-dependencies/monaco/package/ThirdPartyNotices.txt (Join-Path $OutputDirectory 'MONACO-ThirdPartyNotices.txt')
} finally { Pop-Location }

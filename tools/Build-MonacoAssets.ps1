param([switch]$Offline)
$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
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
    & $builder assets/editor/src/editor.js --bundle --minify --format=iife --loader:.ttf=file '--asset-names=[name]-[hash]' --outfile=assets/editor/dist/editor.js
    if ($LASTEXITCODE -ne 0) { throw 'Editor bundle failed.' }
    & $builder artifacts/monaco-dependencies/monaco/package/esm/vs/editor/editor.worker.js --bundle --minify --format=esm --outfile=assets/editor/dist/editor.worker.js
    if ($LASTEXITCODE -ne 0) { throw 'Worker bundle failed.' }
    Copy-Item assets/editor/src/index.html assets/editor/dist/index.html
    foreach ($language in @('fr', 'de', 'es', 'it', 'pt-br', 'ru', 'ja', 'ko', 'zh-cn', 'zh-tw')) {
        Copy-Item "artifacts/monaco-dependencies/monaco/package/esm/nls.messages.$language.js" assets/editor/dist/
    }
    Copy-Item artifacts/monaco-dependencies/monaco/package/LICENSE assets/editor/dist/MONACO-LICENSE.txt
    Copy-Item artifacts/monaco-dependencies/monaco/package/ThirdPartyNotices.txt assets/editor/dist/MONACO-ThirdPartyNotices.txt
} finally { Pop-Location }

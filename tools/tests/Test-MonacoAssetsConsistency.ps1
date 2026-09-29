$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$deliveredDirectory = Join-Path $repository 'assets/editor/dist'
$cachedDependencies = Join-Path $repository 'artifacts/monaco-dependencies'
$buildScript = Join-Path $repository 'tools/Build-MonacoAssets.ps1'
$temporaryBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
$temporaryRoot = Join-Path $temporaryBase ('vbai-monaco-assets-' + [Guid]::NewGuid().ToString('N'))
$workspace = Join-Path $temporaryRoot 'workspace'
$generatedDirectory = Join-Path $temporaryRoot 'dist'
$temporaryRootCreated = $false

try {
    if (Test-Path -LiteralPath $temporaryRoot) { throw "Temporary directory already exists: $temporaryRoot" }
    New-Item -ItemType Directory -Path $temporaryRoot -ErrorAction Stop | Out-Null
    $temporaryRootCreated = $true
    foreach ($name in @('monaco', 'esbuild')) {
        $archive = Join-Path $cachedDependencies "$name/package.tgz"
        if (-not (Test-Path -LiteralPath $archive -PathType Leaf)) {
            throw "Offline asset check requires the cached archive: $archive"
        }
    }

    $sourceEditor = Join-Path $repository 'assets/editor/src'
    $workspaceEditor = Join-Path $workspace 'assets/editor'
    $workspaceDependencies = Join-Path $workspace 'artifacts/monaco-dependencies'
    $workspaceTools = Join-Path $workspace 'tools'
    New-Item -ItemType Directory -Path $workspaceEditor, $workspaceDependencies, $workspaceTools -Force | Out-Null
    Copy-Item -LiteralPath $sourceEditor -Destination $workspaceEditor -Recurse
    Copy-Item -LiteralPath (Join-Path $repository 'assets/editor/monaco.lock.json') -Destination $workspaceEditor
    Copy-Item -LiteralPath (Join-Path $repository 'assets/editor/esbuild.lock.json') -Destination $workspaceEditor
    Copy-Item -LiteralPath $buildScript -Destination $workspaceTools
    foreach ($name in @('monaco', 'esbuild')) {
        $destination = Join-Path $workspaceDependencies $name
        New-Item -ItemType Directory -Path $destination -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $cachedDependencies "$name/package.tgz") -Destination $destination
    }

    $workspaceBuildScript = Join-Path $workspaceTools 'Build-MonacoAssets.ps1'
    & $workspaceBuildScript -Offline -OutputDirectory $generatedDirectory
    if ($LASTEXITCODE -ne 0) { throw 'Monaco asset generation failed.' }

    $delivered = @{}
    Get-ChildItem -LiteralPath $deliveredDirectory -File -Recurse | ForEach-Object {
        $relativePath = $_.FullName.Substring($deliveredDirectory.Length).TrimStart('\', '/').Replace('\', '/')
        $delivered[$relativePath] = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
    }
    $generated = @{}
    Get-ChildItem -LiteralPath $generatedDirectory -File -Recurse | ForEach-Object {
        $relativePath = $_.FullName.Substring($generatedDirectory.Length).TrimStart('\', '/').Replace('\', '/')
        $generated[$relativePath] = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
    }

    $names = @($delivered.Keys + $generated.Keys | Sort-Object -Unique)
    $differences = @($names | Where-Object {
        -not $delivered.ContainsKey($_) -or -not $generated.ContainsKey($_) -or $delivered[$_] -cne $generated[$_]
    })
    if ($differences.Count -gt 0) {
        throw "Monaco assets differ from the generated output: $($differences -join ', ')"
    }
    Write-Output "Monaco assets match byte-for-byte ($($delivered.Count) files)."
} finally {
    $resolvedTemporaryRoot = [IO.Path]::GetFullPath($temporaryRoot)
    $resolvedTemporaryBase = [IO.Path]::GetFullPath($temporaryBase).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    $resolvedParent = [IO.Path]::GetDirectoryName($resolvedTemporaryRoot).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    if ($resolvedTemporaryRoot -cne $temporaryRoot -or $resolvedParent -ine $resolvedTemporaryBase) {
        throw "Refusing to remove a temporary path outside the expected root: $resolvedTemporaryRoot"
    }
    if ($temporaryRootCreated -and (Test-Path -LiteralPath $resolvedTemporaryRoot)) {
        Remove-Item -LiteralPath $resolvedTemporaryRoot -Recurse -Force
    }
}

param(
    [string]$AssemblyPath = 'bin/Debug/net48/VBAi.dll',
    [ValidateSet('Full', 'Unchanged', 'EditOne')][string]$Mode = 'Full',
    [int[]]$ModuleCounts = @(1, 10, 30, 100),
    [ValidateRange(5, 1000)][int]$Iterations = 30,
    [string]$OutputPath = 'artifacts/monaco-performance/index-baseline.json'
)
# Synthetic CPU/serialization probe only: no VBE, host application or user macro.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web.Extensions
$assemblyFile = (Resolve-Path -LiteralPath $AssemblyPath).Path
$assembly = [Reflection.Assembly]::LoadFrom($assemblyFile)
$flags = [Reflection.BindingFlags]'Static,NonPublic'
$sourceType = $assembly.GetType('VBAi.EditorSource', $true)
$build = $assembly.GetType('VBAi.EditorLanguageIndex', $true).GetMethod('Build', $flags)
$serializer = New-Object Web.Script.Serialization.JavaScriptSerializer
$serializer.MaxJsonLength = [int]::MaxValue
$cacheType = $assembly.GetType('VBAi.EditorLanguageCache')
$cacheBuild = if ($cacheType) { $cacheType.GetMethod('Build', [Reflection.BindingFlags]'Instance,NonPublic') }
function InternalField($item, [string]$name) { ,$item.GetType().GetField($name, [Reflection.BindingFlags]'Instance,NonPublic').GetValue($item) }
$results = @()
foreach ($count in $ModuleCounts) {
    if ($count -lt 1 -or $count -gt 1000) { throw 'Module count must be between 1 and 1000.' }
    $sources = [Array]::CreateInstance($sourceType, $count)
    for ($i = 0; $i -lt $count; $i++) {
        $source = [Activator]::CreateInstance($sourceType, $true)
        $source.Module = 'Module' + $i
        $source.ComponentType = 1
        $lines = @('Option Explicit')
        for ($p = 0; $p -lt 25; $p++) {
            $lines += @("Public Sub Example${p}(ByVal value As Long)", '    Dim total As Long', '    total = value + 1', 'End Sub')
        }
        $source.Text = $lines -join "`n"
        $sources.SetValue($source, $i)
    }
    $arguments = New-Object object[] 1
    $arguments[0] = $sources
    $buildTimes = @(); $jsonTimes = @()
    $cache = if ($Mode -ne 'Full') { [Activator]::CreateInstance($cacheType, $true) }
    $previousKey = $null; $knownParts = @{}; $baseFirst = $sources[0].Text
    for ($iteration = -3; $iteration -lt $Iterations; $iteration++) {
        if ($Mode -eq 'EditOne') { $sources[0].Text = $baseFirst + "`n' revision " + $iteration }
        $watch = [Diagnostics.Stopwatch]::StartNew()
        if ($Mode -eq 'Full') { $symbols = $build.Invoke($null, $arguments) }
        else {
            $snapshot = $cacheBuild.Invoke($cache, @($sources, [string[]]@(), [Threading.CancellationToken]::None))
            $symbols = InternalField $snapshot 'Symbols'
        }
        $watch.Stop(); $buildMs = $watch.Elapsed.TotalMilliseconds
        if ($Mode -eq 'Full') { $response = @{ id = 'synthetic'; module = 'Module0'; symbols = $symbols; sources = $sources } }
        else {
            $key = 'synthetic:' + (InternalField $snapshot 'Key')
            if ($previousKey -eq $key) { $response = @{ unchanged = $true; key = $key } }
            else {
                $parts = @(foreach ($part in (InternalField $snapshot 'Parts')) {
                    $name = InternalField $part 'Name'; $partKey = InternalField $part 'Key'
                    @{ name = $name; key = $partKey; symbols = $(if ($knownParts[$name] -ne $partKey) { InternalField $part 'Symbols' } else { $null }) }
                    $knownParts[$name] = $partKey
                })
                $response = @{ id = 'synthetic'; module = 'Module0'; key = $key; parts = $parts; sources = @() }
            }
            $previousKey = $key
        }
        $watch.Restart()
        $payload = $serializer.Serialize($response)
        $watch.Stop()
        if ($iteration -ge 0) { $buildTimes += $buildMs; $jsonTimes += $watch.Elapsed.TotalMilliseconds }
    }
    $buildTimes = @($buildTimes | Sort-Object); $jsonTimes = @($jsonTimes | Sort-Object)
    $p50 = [int][Math]::Ceiling($Iterations * .50) - 1
    $p95 = [int][Math]::Ceiling($Iterations * .95) - 1
    $row = [pscustomobject]@{
        mode = $Mode; modules = $count; lines = 101 * $count + $(if ($Mode -eq 'EditOne') { 1 } else { 0 }); symbols = $symbols.Count; iterations = $Iterations
        indexP50Ms = [Math]::Round($buildTimes[$p50], 2); indexP95Ms = [Math]::Round($buildTimes[$p95], 2)
        serializationP50Ms = [Math]::Round($jsonTimes[$p50], 2); serializationP95Ms = [Math]::Round($jsonTimes[$p95], 2)
        responseUtf8Bytes = [Text.Encoding]::UTF8.GetByteCount($payload)
    }
    $results += $row
    $row | Format-List | Out-Host
}
$fullOutput = [IO.Path]::GetFullPath($OutputPath)
[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($fullOutput)) | Out-Null
[pscustomobject]@{
    measuredAtUtc = [DateTime]::UtcNow.ToString('o')
    commit = (git rev-parse HEAD)
    assemblySha256 = (Get-FileHash -LiteralPath $assemblyFile -Algorithm SHA256).Hash
    mode = $Mode
    scope = 'Synthetic project declarations and JSON only; excludes native COM reads, reference libraries, queueing, WebView transport and rendering. Three warmups per case.'
    results = $results
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $fullOutput -Encoding UTF8
Write-Output $fullOutput

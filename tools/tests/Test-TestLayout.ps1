param([string]$ReportPath)

$ErrorActionPreference = 'Stop'
$repository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$productionRoot = Join-Path $repository 'src/CodexVBE'
$unitRoot = Join-Path $repository 'tests/CodexVBE.Tests/Unit'
$sources = @{}
foreach ($file in Get-ChildItem -LiteralPath $productionRoot -Filter '*.cs' -File -Recurse) {
    $relative = $file.FullName.Substring($productionRoot.Length + 1).Replace('\', '/')
    if ($relative -match '^(bin|obj)/') { continue }
    $sources[$relative] = $null
}

foreach ($file in Get-ChildItem -LiteralPath $unitRoot -Filter '*.cs' -File -Recurse) {
    $relative = $file.FullName.Substring($unitRoot.Length + 1).Replace('\', '/')
    if (-not $relative.EndsWith('.Tests.cs', [StringComparison]::Ordinal)) {
        throw "Unit test file must mirror a production file: $relative"
    }
    $source = $relative.Substring(0, $relative.Length - '.Tests.cs'.Length) + '.cs'
    if (-not $sources.ContainsKey($source)) { throw "No production counterpart for Unit/$relative" }
    $sources[$source] = 'Unit/' + $relative
}

$entries = @($sources.Keys | Sort-Object | ForEach-Object {
    [pscustomobject]@{ Source = 'src/CodexVBE/' + $_; Mirror = $sources[$_] }
})
$mirrors = @($entries | Where-Object { $null -ne $_.Mirror }).Count
Write-Output "PASS: $mirrors dedicated test mirrors for $($entries.Count) production files."
Write-Output 'Files without a dedicated mirror may be exercised by scenarios; this is not a coverage measurement.'
if ($ReportPath) {
    $report = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($report)) | Out-Null
    [IO.File]::WriteAllText($report, ($entries | ConvertTo-Json -Depth 3), (New-Object Text.UTF8Encoding($false)))
}

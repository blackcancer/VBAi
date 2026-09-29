param([ValidateRange(30, 300)] [int] $MaxSeconds = 180)

$ErrorActionPreference = 'Stop'
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'An elevated Windows session is required for the kernel registry trace.'
}

$name = 'VBAiRegistryProbe'
$etl = Join-Path $PSScriptRoot 'VBAi-registry.etl'
$status = Join-Path $PSScriptRoot 'VBAi-registry.status'
$stop = Join-Path $PSScriptRoot 'VBAi-registry.stop'
Remove-Item -LiteralPath $status, $stop -ErrorAction SilentlyContinue

try {
    & logman start $name -ets -p '{70EB4F03-C1DE-4F73-A051-33D13D5413BD}' 0x2C00 0x4 -o $etl -f bincirc -max 128
    if ($LASTEXITCODE -ne 0) { throw "logman start failed: $LASTEXITCODE" }
    Set-Content -LiteralPath $status -Value "ACTIVE $([DateTime]::UtcNow.ToString('o'))" -Encoding ASCII
    $deadline = [DateTime]::UtcNow.AddSeconds($MaxSeconds)
    while ([DateTime]::UtcNow -lt $deadline -and -not (Test-Path -LiteralPath $stop)) {
        Start-Sleep -Seconds 1
    }
}
catch {
    Set-Content -LiteralPath $status -Value "ERROR $($_.Exception.Message)" -Encoding ASCII
}
finally {
    & logman stop $name -ets | Out-Null
    Add-Content -LiteralPath $status -Value "STOPPED $([DateTime]::UtcNow.ToString('o'))" -Encoding ASCII
}

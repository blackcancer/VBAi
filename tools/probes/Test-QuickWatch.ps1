param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$excel = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($excel.ProcessName -ne 'EXCEL') { throw 'A disposable Excel process is required.' }

$client = Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-CodexVBE.ps1'
function Invoke-Vbe([hashtable] $request) {
    $payload = ConvertTo-Json -InputObject $request -Compress -Depth 8
    $result = & $client -HostProcessId $HostProcessId -RequestJson $payload | ConvertFrom-Json
    if (-not $result.Ok) { throw "$($request.Command): $($result.Error)" }
    return $result.Data
}

$initial = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
if ($initial.Code.Length -ne 0) { throw 'Use a fresh workbook with an empty ThisWorkbook module.' }

$probe = Join-Path $PSScriptRoot 'Invoke-BreakpointProbe.ps1'
$null = & $probe -HostProcessId $HostProcessId -Stage insert
$null = & $probe -HostProcessId $HostProcessId -Stage toggle
$null = & $probe -HostProcessId $HostProcessId -Stage run

$state = $null
for ($attempt = 0; $attempt -lt 20; $attempt++) {
    $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
    if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq 4) { break }
    Start-Sleep -Milliseconds 100
}
if ($state.Mode -ne 1 -or $state.Selection.StartLine -ne 4) {
    throw 'The disposable macro did not stop on line 4.'
}

$source = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
$watch = Invoke-Vbe @{ Command = 'quick_watch'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; ExpectedMode = 1; StartLine = 4;
    StartColumn = 5; EndColumn = 15; Expression = 'probeValue';
    Procedure = 'CodexBreakpointProbe' }
if ($watch.Expression -ne 'probeValue' -or $watch.Value -ne '1' -or
    $watch.Context -ne 'VBAProject.ThisWorkbook.CodexBreakpointProbe' -or
    $watch.Verification -ne 'NativeDialogReadback') {
    throw "Unexpected Quick Watch result: $($watch | ConvertTo-Json -Compress)"
}

[pscustomobject]@{
    HostProcessId = $HostProcessId
    Expression = $watch.Expression
    Value = $watch.Value
    Context = $watch.Context
    Verification = $watch.Verification
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$excel = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($excel.ProcessName -ne 'EXCEL') { throw 'A disposable Excel process is required.' }
$client = Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-VBAi.ps1'
function Invoke-Vbe([hashtable] $request) {
    $json = ConvertTo-Json -InputObject $request -Compress -Depth 8
    $reply = & $client -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $reply.Ok) { throw "$($request.Command): $($reply.Error)" }
    return $reply.Data
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

$expression = 'probeValue = 2'
$context = 'ThisWorkbook.CodexBreakpointProbe'
$null = Invoke-Vbe @{ Command = 'add_watch'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedMode = 1; Procedure = 'CodexBreakpointProbe'; Expression = $expression;
    WatchType = 'expression' }
$before = Invoke-Vbe @{ Command = 'debug_windows' }
$watchBefore = @($before.Watches.Items | Where-Object {
    $_.Expression -eq $expression -and $_.Context -eq $context })
if ($watchBefore.Count -ne 1 -or $watchBefore[0].Value -notmatch 'Faux|False') {
    throw 'The initial expression watch was not read back as false.'
}

$edit = Invoke-Vbe @{ Command = 'edit_watch'; Project = 'VBAProject'; ExpectedMode = 1;
    Expression = $expression; NewExpression = $expression; Context = $context;
    WatchType = 'break_when_true' }
if (-not $edit.Edited -or $edit.Verification -ne 'ReadbackVerified') {
    throw 'The edited watch was not read back.'
}

$source = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
$commands = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'Continuer' })
$control = $commands | Where-Object { $_.Id -eq 186 -and $_.Enabled } | Select-Object -First 1
if (-not $control) { throw 'Native Continue command 186 is unavailable.' }
$null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 4; ExpectedMode = 1;
    Action = 'continue'; ControlId = 186; ControlCaption = $control.Caption }

$after = $null
for ($attempt = 0; $attempt -lt 20; $attempt++) {
    $after = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
    if ($after.Mode -eq 1 -and $after.Selection.StartLine -eq 5) { break }
    Start-Sleep -Milliseconds 100
}
$windows = Invoke-Vbe @{ Command = 'debug_windows' }
$watchAfter = @($windows.Watches.Items | Where-Object {
    $_.Expression -eq $expression -and $_.Context -eq $context })
if ($after.Mode -ne 1 -or $after.Selection.StartLine -ne 5 -or
    $watchAfter.Count -ne 1 -or $watchAfter[0].Value -notmatch 'Vrai|True') {
    throw 'The edited break_when_true watch did not stop execution at line 5.'
}

[pscustomobject]@{
    HostProcessId = $HostProcessId
    Expression = $expression
    Before = $watchBefore[0].Value
    After = $watchAfter[0].Value
    BreakLine = $after.Selection.StartLine
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

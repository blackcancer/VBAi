param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$excel = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($excel.ProcessName -ne 'EXCEL') { throw 'A disposable Excel process is required.' }
$client = Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-CodexVBE.ps1'
function Invoke-Vbe([hashtable] $request) {
    $json = ConvertTo-Json -InputObject $request -Compress -Depth 8
    $response = & $client -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $response.Ok) { throw "$($request.Command): $($response.Error)" }
    return $response.Data
}

$before = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
if ($before.Code.Length -ne 0) { throw 'Use a fresh workbook with an empty ThisWorkbook module.' }
$code = @'
Public Sub CodexStepOuter()
    Dim value As Long
    value = 1
    CodexStepInner value
    Debug.Print value
End Sub
Private Sub CodexStepInner(ByRef value As Long)
    value = value + 5
End Sub
'@
$null = Invoke-Vbe @{ Command = 'replace_lines'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $before.Sha256; StartLine = 1; Count = 0; Text = $code }
$source = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
$null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 4 }
$breakCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'point' }) |
    Where-Object { $_.Id -eq 51 -and $_.Enabled } | Select-Object -First 1
if (-not $breakCommand) { throw 'Native Toggle Breakpoint command 51 is unavailable.' }
$null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 4; ExpectedMode = 2;
    Action = 'toggle_breakpoint'; ControlId = 51; ControlCaption = $breakCommand.Caption }
$null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 1 }
$runCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'Sub/UserForm' }) |
    Where-Object { $_.Id -eq 186 -and $_.Enabled } | Select-Object -First 1
if (-not $runCommand) { throw 'Native Run Sub command 186 is unavailable.' }
$null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 1; ExpectedMode = 2;
    Action = 'run'; ControlId = 186; ControlCaption = $runCommand.Caption }
$state = $null
for ($attempt = 0; $attempt -lt 30; $attempt++) {
    $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
    if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq 4) { break }
    Start-Sleep -Milliseconds 100
}
if ($state.Mode -ne 1 -or $state.Selection.StartLine -ne 4) {
    throw 'The outer procedure did not stop at its call line.'
}
$stepCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'pas à pas principal' }) |
    Where-Object { $_.Enabled -and $_.Caption -match 'principal' } | Select-Object -First 1
if (-not $stepCommand) { throw 'Native Step Over command is unavailable.' }
$null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 4; ExpectedMode = 1;
    Action = 'step_over'; ControlId = $stepCommand.Id; ControlCaption = $stepCommand.Caption }
for ($attempt = 0; $attempt -lt 30; $attempt++) {
    $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
    if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq 5) { break }
    Start-Sleep -Milliseconds 100
}
if ($state.Mode -ne 1 -or $state.Selection.StartLine -ne 5) {
    throw 'Step Over did not stop at the following outer line.'
}
$null = Invoke-Vbe @{ Command = 'open_debug_pane'; Action = 'locals' }
$windows = Invoke-Vbe @{ Command = 'debug_windows' }
$valueRow = @($windows.Locals.Items) |
    Where-Object { $_.Expression -eq 'value' -and $_.Value -eq '6' } | Select-Object -First 1
if (-not $valueRow) {
    throw "The called procedure's effect was not visible in Locals: $($windows.Locals | ConvertTo-Json -Compress -Depth 5)"
}
[pscustomobject]@{
    HostProcessId = $HostProcessId
    StepCommandId = $stepCommand.Id
    BreakLine = 4
    StepOverLine = $state.Selection.StartLine
    LocalValue = $valueRow.Value
} | Format-List

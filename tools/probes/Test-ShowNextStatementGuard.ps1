param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$excelProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($excelProcess.ProcessName -ne 'EXCEL') { throw 'A disposable Excel process is required.' }
$bridge = Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-CodexVBE.ps1'
function Invoke-Vbe([hashtable] $request) {
    $json = ConvertTo-Json -InputObject $request -Compress -Depth 8
    $response = & $bridge -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $response.Ok) { throw "$($request.Command): $($response.Error)" }
    return $response.Data
}

$initial = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
$code = @'
Public Sub CodexShowProbe()
    Dim value As Long
    value = 1
    value = value + 1
    Debug.Print "CodexShowProbe:" & CStr(value)
End Sub
'@
if ($initial.Code.Trim().Length -eq 0) {
    $null = Invoke-Vbe @{ Command = 'replace_lines'; Project = 'VBAProject'; Module = 'ThisWorkbook';
        ExpectedSha256 = $initial.Sha256; StartLine = 1; Count = 0; Text = $code }
} elseif (($initial.Code.Trim() -replace "`r`n", "`n") -ne $code.Trim()) {
    throw 'Use a fresh workbook or the exact CodexShowProbe source.'
}
$source = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
$null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 4 }
$breakControl = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'point' }) |
    Where-Object { $_.Id -eq 51 -and $_.Enabled } | Select-Object -First 1
if (-not $breakControl) { throw 'Native Toggle Breakpoint command 51 is unavailable.' }
$null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 4; ExpectedMode = 2;
    Action = 'toggle_breakpoint'; ControlId = 51; ControlCaption = $breakControl.Caption }
$runControl = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'Sub/UserForm' }) |
    Where-Object { $_.Id -eq 186 -and $_.Enabled } | Select-Object -First 1
if (-not $runControl) { throw 'Native Run Sub command 186 is unavailable.' }
$null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 1; ExpectedMode = 2;
    Action = 'run'; ControlId = 186; ControlCaption = $runControl.Caption }
$state = $null
for ($attempt = 0; $attempt -lt 30; $attempt++) {
    $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
    if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq 4) { break }
    Start-Sleep -Milliseconds 100
}
if ($state.Mode -ne 1 -or $state.Selection.StartLine -ne 4) {
    throw "Did not stop at line 4: $($state | ConvertTo-Json -Compress -Depth 5)"
}
$null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 2 }
$before = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
if ($before.Selection.StartLine -ne 2) { throw 'Selection did not move away from the execution line.' }
$show = Invoke-Vbe @{ Command = 'debug_global'; Project = 'VBAProject'; ExpectedMode = 1;
    Action = 'show_next_statement' }
$after = $null
for ($attempt = 0; $attempt -lt 30; $attempt++) {
    $after = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
    if ($after.Selection.StartLine -eq 4) { break }
    Start-Sleep -Milliseconds 100
}
if ($after.Mode -ne 1 -or $after.Selection.StartLine -ne 4) {
    throw "Show Next Statement did not restore line 4: $($after | ConvertTo-Json -Compress -Depth 5)"
}
$null = Invoke-Vbe @{ Command = 'debug_global'; Project = 'VBAProject'; ExpectedMode = 1;
    Action = 'reset' }
[pscustomobject]@{
    HostProcessId = $HostProcessId
    BreakLine = 4
    MovedSelectionLine = $before.Selection.StartLine
    RestoredSelectionLine = $after.Selection.StartLine
    ShowControlId = $show.ControlId
} | Format-List

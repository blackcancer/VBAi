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
Public Sub CodexClearProbe()
    Dim value As Long
    value = 1
    value = value + 1
    Debug.Print "CodexClearProbe:" & CStr(value)
End Sub
'@
$null = Invoke-Vbe @{ Command = 'replace_lines'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $before.Sha256; StartLine = 1; Count = 0; Text = $code }
$source = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
$null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 3 }
$breakCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'point' }) |
    Where-Object { $_.Id -eq 51 -and $_.Enabled } | Select-Object -First 1
if (-not $breakCommand) { throw 'Native Toggle Breakpoint command 51 is unavailable.' }
foreach ($line in @(3, 4)) {
    $null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
        ExpectedSha256 = $source.Sha256; StartLine = $line; ExpectedMode = 2;
        Action = 'toggle_breakpoint'; ControlId = 51; ControlCaption = $breakCommand.Caption }
}
$null = Invoke-Vbe @{ Command = 'open_debug_pane'; Action = 'immediate' }
$null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 1 }
$runCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'Sub/UserForm' }) |
    Where-Object { $_.Id -eq 186 -and $_.Enabled } | Select-Object -First 1
if (-not $runCommand) { throw 'Native Run Sub command 186 is unavailable.' }
function Start-Probe {
    $null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
        ExpectedSha256 = $source.Sha256; StartLine = 1 }
    $null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
        ExpectedSha256 = $source.Sha256; StartLine = 1; ExpectedMode = 2;
        Action = 'run'; ControlId = 186; ControlCaption = $runCommand.Caption }
}
Start-Probe
$state = $null
for ($attempt = 0; $attempt -lt 30; $attempt++) {
    $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
    if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq 3) { break }
    Start-Sleep -Milliseconds 100
}
if ($state.Mode -ne 1 -or $state.Selection.StartLine -ne 3) {
    throw 'The probe did not stop at its first breakpoint.'
}
$null = Invoke-Vbe @{ Command = 'debug_global'; Project = 'VBAProject'; ExpectedMode = 1;
    Action = 'reset' }
$reset = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
if ($reset.Mode -ne 2) { throw 'Reset did not restore design mode.' }
$cleared = Invoke-Vbe @{ Command = 'debug_global'; Project = 'VBAProject'; ExpectedMode = 2;
    Action = 'clear_all_breakpoints' }
Start-Probe
$text = $null
for ($attempt = 0; $attempt -lt 30; $attempt++) {
    $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
    if ($state.Mode -eq 1) { throw "Execution still stopped at line $($state.Selection.StartLine)." }
    $windows = Invoke-Vbe @{ Command = 'debug_windows' }
    $text = $windows.Immediate.Text
    if ($text -match 'CodexClearProbe:2') { break }
    Start-Sleep -Milliseconds 100
}
if ($state.Mode -ne 2 -or $text -notmatch 'CodexClearProbe:2') {
    throw "The procedure did not finish with its expected output. Mode=$($state.Mode); Immediate=$text"
}
[pscustomobject]@{
    HostProcessId = $HostProcessId
    FirstStopLine = 3
    ClearCommandId = $cleared.ControlId
    FinalMode = $state.Mode
    OutputObserved = 'CodexClearProbe:2'
} | Format-List

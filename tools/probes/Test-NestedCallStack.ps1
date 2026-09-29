param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$excel = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($excel.ProcessName -ne 'EXCEL') { throw 'A disposable Excel process is required.' }
$client = Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-VBAi.ps1'
function Invoke-Vbe([hashtable] $request) {
    $json = ConvertTo-Json -InputObject $request -Compress -Depth 8
    $response = & $client -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $response.Ok) { throw "$($request.Command): $($response.Error)" }
    return $response.Data
}

$before = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
if ($before.Code.Length -ne 0) { throw 'Use a fresh workbook with an empty ThisWorkbook module.' }
$code = @'
Public Sub CodexStackOuter()
    Dim outerValue As Long
    outerValue = 1
    CodexStackInner
    Debug.Print outerValue
End Sub
Private Sub CodexStackInner()
    Dim innerValue As Long
    innerValue = 2
    innerValue = innerValue + 1
End Sub
'@
$null = Invoke-Vbe @{ Command = 'replace_lines'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $before.Sha256; StartLine = 1; Count = 0; Text = $code }
$source = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
$null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 10 }

$breakCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'point' }) |
    Where-Object { $_.Id -eq 51 -and $_.Enabled } | Select-Object -First 1
if (-not $breakCommand) { throw 'Native Toggle Breakpoint command 51 is unavailable.' }
$null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
    ExpectedSha256 = $source.Sha256; StartLine = 10; ExpectedMode = 2;
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
for ($attempt = 0; $attempt -lt 20; $attempt++) {
    $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
    if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq 10) { break }
    Start-Sleep -Milliseconds 100
}
if ($state.Mode -ne 1 -or $state.Selection.StartLine -ne 10) {
    throw 'The inner procedure did not stop at line 10.'
}
$null = Invoke-Vbe @{ Command = 'open_debug_pane'; Action = 'locals' }
$windows = Invoke-Vbe @{ Command = 'debug_windows'; IncludeCallStack = $true }
if (-not $windows.CallStack.Available -or $windows.CallStack.Error -or
    @($windows.CallStack.Frames).Count -lt 2 -or
    -not (@($windows.CallStack.Frames) -match 'CodexStackInner') -or
    -not (@($windows.CallStack.Frames) -match 'CodexStackOuter')) {
    throw "Nested Call Stack not observed: $($windows.CallStack | ConvertTo-Json -Compress)"
}

[pscustomobject]@{
    HostProcessId = $HostProcessId
    BreakLine = $state.Selection.StartLine
    Frames = ($windows.CallStack.Frames -join ' | ')
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

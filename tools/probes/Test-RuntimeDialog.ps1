param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') { throw 'An isolated Excel process is required.' }

function Invoke-Vbe([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $response = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $json |
        ConvertFrom-Json
    if (-not $response.Ok) { throw "$($Request.Command): $($response.Error)" }
    return $response.Data
}

$project = 'VBAProject'
$module = 'ThisWorkbook'
$before = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
if ($before.Code.Length -ne 0) { throw 'Use a fresh disposable workbook with an empty ThisWorkbook module.' }
$code = "Public Sub CodexRuntimeProbe()`r`n    Err.Raise 11`r`nEnd Sub"
Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
    ExpectedSha256 = $before.Sha256; StartLine = 1; Count = 0; Text = $code } | Out-Null
$source = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
Invoke-Vbe @{ Command = 'select_code'; Project = $project; Module = $module;
    ExpectedSha256 = $source.Sha256; StartLine = 1 } | Out-Null
$control = @(Invoke-Vbe @{ Command = 'list_commands' } | Where-Object {
    $_.Id -eq 186 -and $_.Enabled -and $_.Caption -match 'Sub/UserForm'
}) | Select-Object -First 1
if (-not $control) { throw 'The native Run Sub/UserForm command is unavailable.' }
Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
    ExpectedSha256 = $source.Sha256; StartLine = 1; ExpectedMode = 2;
    Action = 'run'; ControlId = 186; ControlCaption = $control.Caption } | Out-Null

$dialog = $null
for ($attempt = 0; $attempt -lt 30; $attempt++) {
    $dialog = Invoke-Vbe @{ Command = 'debug_dialog' }
    if ($dialog.Available) { break }
    Start-Sleep -Milliseconds 100
}
if (-not $dialog.Available -or $dialog.Error -or
    $dialog.Diagnostic -notmatch "Erreur d'exécution '11'|Run-time error '11'" -or
    @($dialog.Buttons | Where-Object { $_ -eq 'OK' }).Count -ne 1) {
    throw 'The expected native run-time error dialog was not captured.'
}
$closed = Invoke-Vbe @{ Command = 'respond_debug_dialog';
    Diagnostic = $dialog.Diagnostic; Button = 'OK' }
$after = Invoke-Vbe @{ Command = 'debug_dialog' }
$state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
if ($closed.Verification -ne 'DialogClosed' -or $after.Available -or $state.Mode -ne 2)
{ throw 'The native diagnostic did not close into design mode.' }

[pscustomobject]@{
    HostProcessId = $HostProcessId
    Diagnostic = $dialog.Diagnostic
    Button = $closed.Button
    ModeAfter = $state.Mode
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [Parameter(Mandatory = $true)] [ValidateSet('insert', 'state', 'toggle', 'run', 'continue', 'step_into', 'remove')]
    [string] $Stage
)

$ErrorActionPreference = 'Stop'
$client = Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-VBAi.ps1'
$hosts = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($hosts.Count -ne 1 -or $hosts[0].Id -ne $HostProcessId) {
    throw 'The breakpoint probe requires one isolated Excel process matching HostProcessId.'
}

function Invoke-Vbe([hashtable] $request) {
    $json = ConvertTo-Json -InputObject $request -Compress -Depth 10
    $response = & $client -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $response.Ok) { throw $response.Error }
    return $response.Data
}

$project = 'VBAProject'
$module = 'ThisWorkbook'
if ($Stage -eq 'insert') {
    $before = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
    if ($before.Code.Length -ne 0) { throw 'The probe module is not empty.' }
    $code = @'
Public Sub CodexBreakpointProbe()
    Dim probeValue As Long
    probeValue = 1
    probeValue = probeValue + 1
    Debug.Print probeValue
End Sub
'@
    Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
        ExpectedSha256 = $before.Sha256; StartLine = 1; Count = 0; Text = $code }
    $after = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
    Write-Output "Inserted temporary probe: $($after.Sha256)"
    return
}

$source = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
if ($source.Code -notmatch 'CodexBreakpointProbe') { throw 'Probe procedure is missing.' }
$state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
if ($Stage -eq 'state') { $state; return }

$action = switch ($Stage) {
    'toggle' { 'toggle_breakpoint' }
    'remove' { 'toggle_breakpoint' }
    'run' { 'run' }
    'continue' { 'continue' }
    'step_into' { 'step_into' }
}
$controlId = switch ($Stage) {
    'toggle' { 51 }
    'remove' { 51 }
    'run' { 186 }
    'continue' { 186 }
    'step_into' { 188 }
}
$targetLine = if ($Stage -eq 'toggle' -or $Stage -eq 'remove') { 4 }
    elseif ($Stage -eq 'run') { 1 }
    else { [int]$state.Selection.StartLine }
$null = Invoke-Vbe @{ Command = 'select_code'; Project = $project; Module = $module;
    ExpectedSha256 = $source.Sha256; StartLine = $targetLine }
$query = switch ($Stage) {
    'toggle' { 'point' }
    'remove' { 'point' }
    'run' { 'Sub/UserForm' }
    'continue' { 'Contin' }
    'step_into' { 'pas' }
}
$commands = @(Invoke-Vbe @{ Command = 'list_commands'; Query = $query })
$control = $commands | Where-Object { $_.Id -eq $controlId -and $_.Enabled } | Select-Object -First 1
if (-not $control) { throw "Native VBE command $controlId is absent or disabled." }
Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
    ExpectedSha256 = $source.Sha256; StartLine = $targetLine;
    ExpectedMode = $state.Mode; Action = $action;
    ControlId = $controlId; ControlCaption = $control.Caption }
Invoke-Vbe @{ Command = 'debug_state'; Project = $project }

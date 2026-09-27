param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
if ((Get-Process -Id $HostProcessId -ErrorAction Stop).ProcessName -ne 'EXCEL') {
    throw 'An isolated Excel process is required.'
}

function Invoke-VbeRaw([hashtable] $Request) {
    $payload = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    return (& (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $payload |
        ConvertFrom-Json)
}

function Invoke-Vbe([hashtable] $Request) {
    $response = Invoke-VbeRaw $Request
    if (-not $response.Ok) { throw "$($Request.Command): $($response.Error)" }
    return $response.Data
}

$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2 -or $projects[0].FileName) {
    throw 'Use one unsaved disposable design-mode VBA project.'
}
$project = [string]$projects[0].Name
$module = 'CodexRangeProbe'
$created = Invoke-Vbe @{ Command = 'create_module'; Project = $project; Module = $module; ExpectedMode = 2 }
$source = "Option Explicit`r`nPublic Sub Sample()`r`n    Debug.Print 42`r`nEnd Sub"
Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
    ExpectedSha256 = $created.Sha256; StartLine = 1; Count = $created.Lines; Text = $source } | Out-Null
$before = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
$selected = Invoke-Vbe @{ Command = 'select_code_range'; Project = $project; Module = $module;
    ExpectedSha256 = $before.Sha256; StartLine = 2; StartColumn = 1; EndLine = 3; EndColumn = 11 }
if (-not $selected.Verified -or $selected.StartLine -ne 2 -or $selected.EndLine -ne 3) {
    throw 'The native multi-line selection was not verified.'
}
$pane = Invoke-Vbe @{ Command = 'code_panes' }
$selection = $pane.ActiveCodePane.Properties.Selection
if ($selection.StartLine -ne 2 -or $selection.StartColumn -ne 1 -or
    $selection.EndLine -ne 3 -or $selection.EndColumn -ne 11) {
    throw 'A separate CodePanes read did not retain the selection.'
}
$invalid = Invoke-VbeRaw @{ Command = 'select_code_range'; Project = $project; Module = $module;
    ExpectedSha256 = $before.Sha256; StartLine = 3; StartColumn = 1; EndLine = 2; EndColumn = 1 }
if ($invalid.Ok) { throw 'A reversed code range was accepted.' }
$stale = Invoke-VbeRaw @{ Command = 'select_code_range'; Project = $project; Module = $module;
    ExpectedSha256 = ('0' * 64); StartLine = 2; StartColumn = 1; EndLine = 3; EndColumn = 11 }
if ($stale.Ok) { throw 'A stale module SHA was accepted.' }
[pscustomobject]@{ HostProcessId = $HostProcessId; ModuleSha256 = $before.Sha256;
    Verified = $selected.Verified; SeparateReadVerified = $true;
    ReversedRangeRejected = -not $invalid.Ok; StaleShaRejected = -not $stale.Ok } | Format-List

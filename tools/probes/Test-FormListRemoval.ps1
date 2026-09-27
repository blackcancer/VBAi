param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') { throw 'The probe requires Excel.' }
function Invoke-Vbe([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $reply = & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $reply.Ok) { throw "$($Request.Command): $($reply.Error)" }
    return $reply.Data
}

$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) { throw 'Expected one disposable design-mode project.' }
$project = $projects[0].Name
$form = 'CodexListRemovalProbe'
$control = 'lstProbe'
if (@(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } | Where-Object { $_.Name -eq $form }).Count) {
    throw "Form $form already exists."
}
Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form } | Out-Null
$tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
    Control = $control; ControlType = 'Forms.ListBox.1'; Left = 24; Top = 24;
    Width = 140; Height = 80; ExpectedFormVersion = $tree.FormVersion } | Out-Null
foreach ($value in @('Alpha', 'Beta', 'Gamma')) {
    $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
    $append = Invoke-Vbe @{ Command = 'probe_append_form_list_item'; Project = $project;
        Form = $form; ControlPath = "Controls/$control"; Text = $value;
        ExpectedTreeVersion = $tree.TreeVersion }
    if (-not $append.Verified) { throw "AddItem not verified for $value." }
}
$stale = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
    ControlPath = "Controls/$control"; Offset = 0; Limit = 64 }
if ($stale.TotalRows -ne 3 -or -not $stale.ListVersion) { throw 'Full list version is unavailable.' }
$tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
Invoke-Vbe @{ Command = 'probe_append_form_list_item'; Project = $project;
    Form = $form; ControlPath = "Controls/$control"; Text = 'Delta';
    ExpectedTreeVersion = $tree.TreeVersion } | Out-Null
$current = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
    ControlPath = "Controls/$control"; Offset = 0; Limit = 64 }
if ($current.TotalRows -ne 4 -or $current.ListVersion -eq $stale.ListVersion) {
    throw 'The content revision did not detect the append.'
}
$staleRequest = @{ Command = 'remove_form_list_item'; Project = $project;
    Form = $form; ControlPath = "Controls/$control"; RowIndex = 1;
    ExpectedTreeVersion = $current.TreeVersion; ExpectedListVersion = $stale.ListVersion }
$staleReply = & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson (ConvertTo-Json $staleRequest -Compress) | ConvertFrom-Json
if ($staleReply.Ok -or $staleReply.Error -notmatch 'list changed') {
    throw 'A stale ListVersion was not refused before RemoveItem.'
}
$unchanged = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
    ControlPath = "Controls/$control"; Offset = 0; Limit = 64 }
if ($unchanged.ListVersion -ne $current.ListVersion) { throw 'Stale refusal changed the list.' }
$removed = Invoke-Vbe @{ Command = 'remove_form_list_item'; Project = $project;
    Form = $form; ControlPath = "Controls/$control"; RowIndex = 1;
    ExpectedTreeVersion = $current.TreeVersion; ExpectedListVersion = $current.ListVersion }
if (-not $removed.Applied -or -not $removed.Verified -or $removed.RemovedValue -ne 'Beta') {
    throw 'RemoveItem was not verified.'
}
$after = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
    ControlPath = "Controls/$control"; Offset = 0; Limit = 64 }
$values = @($after.Rows | ForEach-Object { $_.Cells[0].Value })
if ($after.TotalRows -ne 3 -or ($values -join '|') -ne 'Alpha|Gamma|Delta' -or
    $after.ListVersion -ne $removed.ListVersionAfter) {
    throw 'The requested item did not disappear exactly from the list.'
}
[pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project; Form = $form;
    StaleVersionRefused = $true; RemovedValue = $removed.RemovedValue;
    FinalValues = $values; FinalListVersion = $after.ListVersion;
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue) } | ConvertTo-Json -Depth 5

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
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) { throw 'Expected one disposable project in design mode.' }
$project = $projects[0].Name
$form = 'CodexColumnGuardProbe'
if (@(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } | Where-Object { $_.Name -eq $form }).Count) {
    throw "Form $form already exists."
}
Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form } | Out-Null
$tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
    Control = 'cboProbe'; ControlType = 'Forms.ComboBox.1'; Left = 24; Top = 24;
    Width = 140; Height = 30; ExpectedFormVersion = $tree.FormVersion } | Out-Null
$before = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
$property = @($before.Controls[0].Properties | Where-Object { $_.Name -eq 'ColumnCount' })[0]
if ($property.SetterStatus -ne 'BlockedAfterHostCrash') {
    throw "Unexpected ColumnCount status: $($property.SetterStatus)"
}
$json = ConvertTo-Json -InputObject @{ Command = 'set_form_node_property'; Project = $project;
    Form = $form; ControlPath = 'Controls/cboProbe'; Property = 'ColumnCount';
    Value = 2; ExpectedTreeVersion = $before.TreeVersion } -Compress
$reply = & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
if ($reply.Ok -or $reply.Error -notmatch 'temporarily disabled') {
    throw 'ComboBox.ColumnCount was not blocked before COM.'
}
$after = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
if ($before.TreeVersion -ne $after.TreeVersion) { throw 'The refused setter changed TreeVersion.' }
[pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project; Form = $form;
    SetterStatus = $property.SetterStatus; Refused = $true;
    TreeVersionStable = $true; ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue) } | ConvertTo-Json

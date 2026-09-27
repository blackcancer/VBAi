param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [Parameter(Mandatory = $true)] [string] $ControlType,
    [Parameter(Mandatory = $true)] [string] $Property,
    [Parameter(Mandatory = $true)] [string] $ValueJson,
    [string] $Form = 'CodexSingleSetterSurvey',
    [string] $Control = 'probeControl'
)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') { throw 'The probe requires Excel.' }
$value = ConvertFrom-Json -InputObject $ValueJson

function Invoke-Vbe([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $reply = & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $reply.Ok) { throw "$($Request.Command): $($reply.Error)" }
    return $reply.Data
}

function Read-Tree {
    Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $Form }
}

function Read-Target([object] $Tree) {
    $node = @($Tree.Controls | Where-Object { $_.Name -eq $Control })[0]
    if (-not $node) { throw "Control $Control is absent from form_tree." }
    $parts = $Property.Split('.')
    if ($parts.Count -gt 2) { throw 'Only a scalar or one object member is supported by this probe.' }
    $root = @($node.Properties | Where-Object { $_.Name -eq $parts[0] })[0]
    if (-not $root) { throw "Property $Property is absent from form_tree." }
    if ($parts.Count -eq 1) { return $root }
    $member = @($root.Members | Where-Object { $_.Name -eq $parts[1] })[0]
    if (-not $member) { throw "Object member $Property is absent from form_tree." }
    return $member
}

$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) {
    throw 'Expected exactly one disposable VBA project in design mode.'
}
$project = $projects[0].Name
if (@(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } | Where-Object { $_.Name -eq $Form }).Count) {
    throw "Form $Form already exists."
}
Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $Form } | Out-Null
$initial = Read-Tree
Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $Form;
    Control = $Control; ControlType = $ControlType; Left = 24; Top = 24; Width = 120; Height = 30;
    ExpectedFormVersion = $initial.FormVersion } | Out-Null
$before = Read-Tree
$old = Read-Target $before
Invoke-Vbe @{ Command = 'set_form_node_property'; Project = $project; Form = $Form;
    ControlPath = "Controls/$Control"; Property = $Property; Value = $value;
    ExpectedTreeVersion = $before.TreeVersion } | Out-Null
$after = Read-Tree
$new = Read-Target $after
$stable = Read-Tree
if ($after.TreeVersion -eq $before.TreeVersion -or
    [string]$old.Value -eq [string]$new.Value -or $new.Error) {
    throw "The $ControlType.$Property write was not observed as a changed, readable value in form_tree."
}
if ($stable.TreeVersion -ne $after.TreeVersion) {
    throw 'The form_tree version changed on a read-only follow-up.'
}
[pscustomobject]@{
    HostProcessId = $HostProcessId
    Project = $project
    Form = $Form
    ControlType = $ControlType
    Property = $Property
    RequestedValue = $ValueJson
    BeforeValue = $old.Value
    AfterValue = $new.Value
    BeforeSetterStatus = $old.SetterStatus
    AfterSetterStatus = $new.SetterStatus
    VersionChanged = ($after.TreeVersion -ne $before.TreeVersion)
    VersionStableOnReread = $true
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

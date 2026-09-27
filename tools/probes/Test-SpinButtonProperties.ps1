param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [string] $Project = 'VBAProject',
    [string] $Form = 'CodexSpinPropertySurvey'
)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') {
    throw "The probe requires Excel; PID $HostProcessId is $($hostProcess.ProcessName)."
}

function Invoke-Vbe([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $reply = & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $reply.Ok) { throw "$($Request.Command): $($reply.Error)" }
    return $reply.Data
}

function Read-Tree {
    Invoke-Vbe @{ Command = 'form_tree'; Project = $Project; Form = $Form }
}

function Read-SpinProperties {
    $all = @(Invoke-Vbe @{ Command = 'form_control_properties'; Project = $Project; Form = $Form; Control = 'spnProbe' })
    $values = @{}
    foreach ($property in $all) {
        if ($property.Name -in @('Min', 'Max', 'Value', 'Delay', 'SmallChange')) {
            $values[$property.Name] = $property.Value
        }
    }
    return $values
}

$existing = @(Invoke-Vbe @{ Command = 'list_forms'; Project = $Project })
if (@($existing | Where-Object { $_.Name -eq $Form }).Count) {
    throw "Form $Form already exists; use a new name in a disposable workbook."
}

Invoke-Vbe @{ Command = 'create_form'; Project = $Project; Form = $Form } | Out-Null
$initial = Read-Tree
Invoke-Vbe @{ Command = 'add_form_control'; Project = $Project; Form = $Form;
    ControlType = 'Forms.SpinButton.1'; Control = 'spnProbe';
    Left = 24; Top = 24; Width = 32; Height = 80;
    ExpectedFormVersion = $initial.FormVersion } | Out-Null
$before = Read-Tree
$original = Read-SpinProperties

$after = Read-Tree
$final = Read-SpinProperties
[pscustomobject]@{
    HostProcessId = $HostProcessId
    Form = $Form
    ReadOnlyProbe = $true
    BeforeNodeCount = $before.NodeCount
    AfterNodeCount = $after.NodeCount
    BeforeTreeVersion = $before.TreeVersion
    AfterTreeVersion = $after.TreeVersion
    OriginalProperties = ($original | ConvertTo-Json -Compress)
    FinalProperties = ($final | ConvertTo-Json -Compress)
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

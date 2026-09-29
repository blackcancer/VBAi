param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [string] $Project = 'VBAProject',
    [string] $Form = 'CodexFramePlanSurvey'
)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') {
    throw "The probe requires Excel; PID $HostProcessId is $($hostProcess.ProcessName)."
}

function Invoke-Vbe([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $reply = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $reply.Ok) { throw "$($Request.Command): $($reply.Error)" }
    return $reply.Data
}

function Read-Tree {
    Invoke-Vbe @{ Command = 'form_tree'; Project = $Project; Form = $Form }
}

$existing = @(Invoke-Vbe @{ Command = 'list_forms'; Project = $Project })
if (@($existing | Where-Object { $_.Name -eq $Form }).Count) {
    throw "Form $Form already exists; use a new name in a disposable workbook."
}

Invoke-Vbe @{ Command = 'create_form'; Project = $Project; Form = $Form } | Out-Null
$initial = Read-Tree
Invoke-Vbe @{ Command = 'add_form_control'; Project = $Project; Form = $Form;
    ControlType = 'Forms.Frame.1'; Control = 'fraOriginal'; Caption = 'Groupe';
    Left = 24; Top = 24; Width = 160; Height = 100;
    ExpectedFormVersion = $initial.FormVersion } | Out-Null
$empty = Read-Tree
Invoke-Vbe @{ Command = 'add_nested_form_control'; Project = $Project; Form = $Form;
    ParentPath = 'Controls/fraOriginal'; ControlType = 'Forms.Label.1';
    Control = 'lblChild'; Caption = 'Enfant'; Left = 8; Top = 18; Width = 60; Height = 20;
    ExpectedTreeVersion = $empty.TreeVersion } | Out-Null
$withLabel = Read-Tree
$eligible = Invoke-Vbe @{ Command = 'frame_copy_plan'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/fraOriginal'; NewName = 'fraTarget';
    ExpectedTreeVersion = $withLabel.TreeVersion }
if (-not $eligible.ReadOnly -or $eligible.MutationVerified -or
    -not $eligible.EligibleForLimitedProbe -or $eligible.DirectChildCount -ne 1 -or
    @($eligible.Children).Count -ne 1 -or
    $eligible.Children[0].ProposedPath -ne 'Controls/fraTarget/Controls/fraTarget_lblChild') {
    throw 'Frame plan did not identify the direct Label and proposed canonical path.'
}

Invoke-Vbe @{ Command = 'add_nested_form_control'; Project = $Project; Form = $Form;
    ParentPath = 'Controls/fraOriginal'; ControlType = 'Forms.TextBox.1';
    Control = 'txtChild'; Left = 8; Top = 46; Width = 70; Height = 20;
    ExpectedTreeVersion = $withLabel.TreeVersion } | Out-Null
$withTextBox = Read-Tree
$ineligible = Invoke-Vbe @{ Command = 'frame_copy_plan'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/fraOriginal'; NewName = 'fraTarget';
    ExpectedTreeVersion = $withTextBox.TreeVersion }
$afterPlan = Read-Tree
if ($ineligible.EligibleForLimitedProbe -or $ineligible.DirectChildCount -ne 2 -or
    @($ineligible.Issues).Count -eq 0 -or $afterPlan.TreeVersion -ne $withTextBox.TreeVersion) {
    throw 'Frame plan did not reject the unsupported TextBox child without mutation.'
}

[pscustomobject]@{
    HostProcessId = $HostProcessId
    Form = $Form
    EligibleWithLabel = $eligible.EligibleForLimitedProbe
    LabelProposedPath = $eligible.Children[0].ProposedPath
    EligibleWithTextBox = $ineligible.EligibleForLimitedProbe
    IssuesWithTextBox = ($ineligible.Issues -join '; ')
    ReadOnlyVersionUnchanged = ($afterPlan.TreeVersion -eq $withTextBox.TreeVersion)
    FinalNodeCount = $afterPlan.NodeCount
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

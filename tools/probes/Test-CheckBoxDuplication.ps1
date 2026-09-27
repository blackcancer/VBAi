param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [string] $Project = 'VBAProject',
    [string] $Form = 'CodexCheckBoxCopySurvey'
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

$existing = @(Invoke-Vbe @{ Command = 'list_forms'; Project = $Project })
if (@($existing | Where-Object { $_.Name -eq $Form }).Count) {
    throw "Form $Form already exists; use a new name in a disposable workbook."
}

Invoke-Vbe @{ Command = 'create_form'; Project = $Project; Form = $Form } | Out-Null
$initial = Read-Tree
Invoke-Vbe @{ Command = 'add_form_control'; Project = $Project; Form = $Form;
    ControlType = 'Forms.CheckBox.1'; Control = 'chkOriginal'; Caption = 'Choix Codex';
    Left = 24; Top = 24; Width = 120; Height = 30;
    ExpectedFormVersion = $initial.FormVersion } | Out-Null
$withCheckBox = Read-Tree
Invoke-Vbe @{ Command = 'set_form_node_property'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/chkOriginal'; Property = 'Value'; Value = $true;
    ExpectedTreeVersion = $withCheckBox.TreeVersion } | Out-Null
$before = Read-Tree

$copy = Invoke-Vbe @{ Command = 'duplicate_form_checkbox'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/chkOriginal'; NewName = 'chkCopy';
    ExpectedTreeVersion = $before.TreeVersion }
$after = Read-Tree
if ($copy.Completeness -ne 'Partial' -or $copy.NewPath -ne 'Controls/chkCopy' -or
    $after.NodeCount -ne ($before.NodeCount + 1) -or $after.TreeVersion -eq $before.TreeVersion) {
    throw 'DuplicateCheckBox did not return the expected partial copy and new tree revision.'
}

[pscustomobject]@{
    HostProcessId = $HostProcessId
    Form = $Form
    SourcePath = $copy.SourcePath
    NewPath = $copy.NewPath
    Completeness = $copy.Completeness
    CopiedProperties = ($copy.CopiedProperties -join ', ')
    BeforeTreeVersion = $before.TreeVersion
    AfterTreeVersion = $after.TreeVersion
    BeforeNodeCount = $before.NodeCount
    AfterNodeCount = $after.NodeCount
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

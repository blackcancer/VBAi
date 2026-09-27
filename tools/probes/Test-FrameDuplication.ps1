param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [string] $Project = 'VBAProject',
    [string] $Form = 'CodexFrameCopySurvey'
)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') {
    throw "The probe requires Excel; PID $HostProcessId is $($hostProcess.ProcessName)."
}

function Invoke-VbeRaw([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
}

function Invoke-Vbe([hashtable] $Request) {
    $reply = Invoke-VbeRaw $Request
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
    Left = 24; Top = 24; Width = 150; Height = 90;
    ExpectedFormVersion = $initial.FormVersion } | Out-Null
$before = Read-Tree

$copy = Invoke-Vbe @{ Command = 'duplicate_empty_form_frame'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/fraOriginal'; NewName = 'fraCopy';
    ExpectedTreeVersion = $before.TreeVersion }
$afterCopy = Read-Tree
if ($copy.Completeness -ne 'Partial' -or $copy.NewPath -ne 'Controls/fraCopy' -or
    $copy.ChildrenCopied -ne $false -or $copy.SourceChildCount -ne 0 -or
    $afterCopy.NodeCount -ne ($before.NodeCount + 1) -or $afterCopy.TreeVersion -eq $before.TreeVersion) {
    throw 'Empty Frame duplication did not produce the expected partial copy.'
}

Invoke-Vbe @{ Command = 'add_nested_form_control'; Project = $Project; Form = $Form;
    ParentPath = 'Controls/fraOriginal'; ControlType = 'Forms.Label.1';
    Control = 'lblChild'; Caption = 'Enfant'; Left = 8; Top = 16; Width = 70; Height = 18;
    ExpectedTreeVersion = $afterCopy.TreeVersion } | Out-Null
$beforeRefusal = Read-Tree
$refusal = Invoke-VbeRaw @{ Command = 'duplicate_empty_form_frame'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/fraOriginal'; NewName = 'fraRejected';
    ExpectedTreeVersion = $beforeRefusal.TreeVersion }
$afterRefusal = Read-Tree
if ($refusal.Ok -or $refusal.Error -notmatch 'child controls' -or
    $afterRefusal.TreeVersion -ne $beforeRefusal.TreeVersion -or
    $afterRefusal.NodeCount -ne $beforeRefusal.NodeCount) {
    throw 'Frame with a child was not refused without mutation.'
}

[pscustomobject]@{
    HostProcessId = $HostProcessId
    Form = $Form
    CopiedPath = $copy.NewPath
    CopiedProperties = ($copy.CopiedProperties -join ', ')
    EmptyCopyNodeCount = $afterCopy.NodeCount
    ChildRefusal = $refusal.Error
    RefusalLeftTreeUnchanged = ($afterRefusal.TreeVersion -eq $beforeRefusal.TreeVersion)
    FinalNodeCount = $afterRefusal.NodeCount
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

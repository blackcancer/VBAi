param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [string] $Project = 'VBAProject',
    [string] $Form = 'CodexFrameLabelCopySurvey'
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
    Left = 24; Top = 24; Width = 180; Height = 110;
    ExpectedFormVersion = $initial.FormVersion } | Out-Null
foreach ($item in @(
    @{ Name = 'lblOne'; Caption = 'Un'; Top = 18 },
    @{ Name = 'lblTwo'; Caption = 'Deux'; Top = 48 }
)) {
    $tree = Read-Tree
    Invoke-Vbe @{ Command = 'add_nested_form_control'; Project = $Project; Form = $Form;
        ParentPath = 'Controls/fraOriginal'; ControlType = 'Forms.Label.1';
        Control = $item.Name; Caption = $item.Caption;
        Left = 8; Top = $item.Top; Width = 80; Height = 20;
        ExpectedTreeVersion = $tree.TreeVersion } | Out-Null
}
$before = Read-Tree
$plan = Invoke-Vbe @{ Command = 'frame_copy_plan'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/fraOriginal'; NewName = 'fraClone';
    ExpectedTreeVersion = $before.TreeVersion }
if (-not $plan.EligibleForLimitedProbe -or $plan.DirectChildCount -ne 2) {
    throw 'Preflight did not find two direct Labels.'
}

$copy = Invoke-Vbe @{ Command = 'duplicate_form_frame_labels'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/fraOriginal'; NewName = 'fraClone';
    ExpectedTreeVersion = $before.TreeVersion }
$afterCopy = Read-Tree
if ($copy.Completeness -ne 'Partial' -or $copy.DirectLabelsCopied -ne 2 -or
    @($copy.CopiedChildPaths).Count -ne 2 -or
    $afterCopy.NodeCount -ne ($before.NodeCount + 3) -or
    $afterCopy.TreeVersion -eq $before.TreeVersion) {
    throw 'Frame with direct Labels was not copied and read back as expected.'
}

Invoke-Vbe @{ Command = 'add_nested_form_control'; Project = $Project; Form = $Form;
    ParentPath = 'Controls/fraOriginal'; ControlType = 'Forms.TextBox.1';
    Control = 'txtUnsupported'; Left = 8; Top = 78; Width = 80; Height = 20;
    ExpectedTreeVersion = $afterCopy.TreeVersion } | Out-Null
$beforeRefusal = Read-Tree
$refusal = Invoke-VbeRaw @{ Command = 'duplicate_form_frame_labels'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/fraOriginal'; NewName = 'fraRejected';
    ExpectedTreeVersion = $beforeRefusal.TreeVersion }
$afterRefusal = Read-Tree
if ($refusal.Ok -or $refusal.Error -notmatch 'ineligible' -or
    $afterRefusal.TreeVersion -ne $beforeRefusal.TreeVersion -or
    $afterRefusal.NodeCount -ne $beforeRefusal.NodeCount) {
    throw 'Unsupported child was not refused before mutation.'
}

[pscustomobject]@{
    HostProcessId = $HostProcessId
    Form = $Form
    CopiedFrame = $copy.NewPath
    CopiedLabels = ($copy.CopiedChildPaths -join ', ')
    BeforeNodeCount = $before.NodeCount
    AfterCopyNodeCount = $afterCopy.NodeCount
    UnsupportedChildRefusal = $refusal.Error
    RefusalLeftTreeUnchanged = ($afterRefusal.TreeVersion -eq $beforeRefusal.TreeVersion)
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

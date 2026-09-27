param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [string] $Project = 'VBAProject',
    [string] $Form = 'CodexProfiledFrameSurvey'
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
    ControlType = 'Forms.Frame.1'; Control = 'fraOriginal'; Caption = 'Profils';
    Left = 24; Top = 24; Width = 250; Height = 240;
    ExpectedFormVersion = $initial.FormVersion } | Out-Null

$items = @(
    @{ Name = 'lblOne'; Type = 'Forms.Label.1'; Caption = 'Nom'; Top = 12 },
    @{ Name = 'txtOne'; Type = 'Forms.TextBox.1'; Caption = $null; Top = 42 },
    @{ Name = 'chkOne'; Type = 'Forms.CheckBox.1'; Caption = 'Actif'; Top = 72 },
    @{ Name = 'cmdOne'; Type = 'Forms.CommandButton.1'; Caption = 'Exécuter'; Top = 102 },
    @{ Name = 'cmbOne'; Type = 'Forms.ComboBox.1'; Caption = $null; Top = 132 },
    @{ Name = 'optOne'; Type = 'Forms.OptionButton.1'; Caption = 'Choix'; Top = 162 }
)
foreach ($item in $items) {
    $tree = Read-Tree
    $request = @{ Command = 'add_nested_form_control'; Project = $Project; Form = $Form;
        ParentPath = 'Controls/fraOriginal'; ControlType = $item.Type;
        Control = $item.Name; Left = 8; Top = $item.Top; Width = 110; Height = 24;
        ExpectedTreeVersion = $tree.TreeVersion }
    if ($null -ne $item.Caption) { $request.Caption = $item.Caption }
    Invoke-Vbe $request | Out-Null
}
foreach ($mutation in @(
    @{ Path = 'Controls/fraOriginal/Controls/txtOne'; Property = 'Value'; Value = 'Texte Codex' },
    @{ Path = 'Controls/fraOriginal/Controls/chkOne'; Property = 'Value'; Value = $true },
    @{ Path = 'Controls/fraOriginal/Controls/cmbOne'; Property = 'ListWidth'; Value = '72 pt' }
)) {
    $tree = Read-Tree
    Invoke-Vbe @{ Command = 'set_form_node_property'; Project = $Project; Form = $Form;
        ControlPath = $mutation.Path; Property = $mutation.Property; Value = $mutation.Value;
        ExpectedTreeVersion = $tree.TreeVersion } | Out-Null
}
$before = Read-Tree
$plan = Invoke-Vbe @{ Command = 'frame_profile_copy_plan'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/fraOriginal'; NewName = 'fraClone';
    ExpectedTreeVersion = $before.TreeVersion }
if (-not $plan.EligibleForLimitedProbe -or $plan.DirectChildCount -ne 6 -or
    @($plan.Children).Count -ne 6) {
    throw 'The positive registry did not accept all six direct child profiles.'
}
$copy = Invoke-Vbe @{ Command = 'duplicate_form_frame_profiled'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/fraOriginal'; NewName = 'fraClone';
    ExpectedTreeVersion = $before.TreeVersion }
$afterCopy = Read-Tree
if ($copy.Completeness -ne 'Partial' -or $copy.DirectChildrenCopied -ne 6 -or
    @($copy.CopiedChildren).Count -ne 6 -or
    $afterCopy.NodeCount -ne ($before.NodeCount + 7) -or
    $afterCopy.TreeVersion -eq $before.TreeVersion) {
    throw 'Profiled Frame copy did not produce seven new canonical nodes.'
}

Invoke-Vbe @{ Command = 'add_nested_form_control'; Project = $Project; Form = $Form;
    ParentPath = 'Controls/fraOriginal'; ControlType = 'Forms.SpinButton.1';
    Control = 'spinUnsupported'; Left = 160; Top = 12; Width = 24; Height = 80;
    ExpectedTreeVersion = $afterCopy.TreeVersion } | Out-Null
$beforeRefusal = Read-Tree
$unsupportedPlan = Invoke-Vbe @{ Command = 'frame_profile_copy_plan'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/fraOriginal'; NewName = 'fraRejected';
    ExpectedTreeVersion = $beforeRefusal.TreeVersion }
$refusal = Invoke-VbeRaw @{ Command = 'duplicate_form_frame_profiled'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/fraOriginal'; NewName = 'fraRejected';
    ExpectedTreeVersion = $beforeRefusal.TreeVersion }
$afterRefusal = Read-Tree
if ($unsupportedPlan.EligibleForLimitedProbe -or $refusal.Ok -or
    $refusal.Error -notmatch 'ineligible' -or
    $afterRefusal.TreeVersion -ne $beforeRefusal.TreeVersion -or
    $afterRefusal.NodeCount -ne $beforeRefusal.NodeCount) {
    throw 'Unsupported SpinButton child was not refused before mutation.'
}

[pscustomobject]@{
    HostProcessId = $HostProcessId
    Form = $Form
    Profiles = (@($copy.CopiedChildren | ForEach-Object { $_.Type }) -join ', ')
    CopiedFrame = $copy.NewPath
    CopiedChildren = $copy.DirectChildrenCopied
    BeforeNodeCount = $before.NodeCount
    AfterCopyNodeCount = $afterCopy.NodeCount
    UnsupportedChildRefusal = $refusal.Error
    RefusalLeftTreeUnchanged = ($afterRefusal.TreeVersion -eq $beforeRefusal.TreeVersion)
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

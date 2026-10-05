param(
    [Parameter(Mandatory = $true)] [ValidateRange(1, [int]::MaxValue)] [int] $HostProcessId,
    [ValidateSet('Empty', 'Labels', 'SimpleChildren', 'Profiled')] [string] $Scenario = 'Empty',
    [ValidateNotNullOrEmpty()] [string] $Project = 'VBAProject',
    [string] $Form,
    [switch] $OwnedDisposableWorkbook
)

$ErrorActionPreference = 'Stop'
if (-not $OwnedDisposableWorkbook) {
    throw 'Confirm that the selected Excel process contains an owned disposable workbook with -OwnedDisposableWorkbook.'
}
$defaultForms = @{
    Empty = 'CodexFrameCopySurvey'
    Labels = 'CodexFrameLabelCopySurvey'
    SimpleChildren = 'CodexFrameSimpleCopySurvey'
    Profiled = 'CodexProfiledFrameSurvey'
}
if (-not $Form) { $Form = $defaultForms[$Scenario] }
if ([string]::IsNullOrWhiteSpace($Form)) { throw 'A nonempty disposable form name is required.' }

$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') {
    throw "The probe requires Excel; PID $HostProcessId is $($hostProcess.ProcessName)."
}

function Invoke-VbeRaw([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
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
switch ($Scenario) {
    'Empty' {
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
    }
    'Labels' {
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
    }
    'SimpleChildren' {
        Invoke-Vbe @{ Command = 'add_form_control'; Project = $Project; Form = $Form;
            ControlType = 'Forms.Frame.1'; Control = 'fraOriginal'; Caption = 'Groupe';
            Left = 24; Top = 24; Width = 180; Height = 110;
            ExpectedFormVersion = $initial.FormVersion } | Out-Null
        $empty = Read-Tree
        Invoke-Vbe @{ Command = 'add_nested_form_control'; Project = $Project; Form = $Form;
            ParentPath = 'Controls/fraOriginal'; ControlType = 'Forms.Label.1';
            Control = 'lblChild'; Caption = 'Nom'; Left = 8; Top = 18; Width = 60; Height = 20;
            ExpectedTreeVersion = $empty.TreeVersion } | Out-Null
        $withLabel = Read-Tree
        Invoke-Vbe @{ Command = 'add_nested_form_control'; Project = $Project; Form = $Form;
            ParentPath = 'Controls/fraOriginal'; ControlType = 'Forms.TextBox.1';
            Control = 'txtChild'; Left = 8; Top = 48; Width = 90; Height = 20;
            ExpectedTreeVersion = $withLabel.TreeVersion } | Out-Null
        $withTextBox = Read-Tree
        Invoke-Vbe @{ Command = 'set_form_node_property'; Project = $Project; Form = $Form;
            ControlPath = 'Controls/fraOriginal/Controls/txtChild'; Property = 'Value'; Value = 'Texte Codex';
            ExpectedTreeVersion = $withTextBox.TreeVersion } | Out-Null
        $before = Read-Tree

        $plan = Invoke-Vbe @{ Command = 'frame_simple_copy_plan'; Project = $Project; Form = $Form;
            ControlPath = 'Controls/fraOriginal'; NewName = 'fraClone';
            ExpectedTreeVersion = $before.TreeVersion }
        if (-not $plan.EligibleForLimitedProbe -or $plan.DirectChildCount -ne 2) {
            throw 'Simple Frame plan did not accept direct Label and TextBox children.'
        }
        $copy = Invoke-Vbe @{ Command = 'duplicate_form_frame_simple_children'; Project = $Project; Form = $Form;
            ControlPath = 'Controls/fraOriginal'; NewName = 'fraClone';
            ExpectedTreeVersion = $before.TreeVersion }
        $afterCopy = Read-Tree
        if ($copy.Completeness -ne 'Partial' -or $copy.DirectLabelsCopied -ne 1 -or
            $copy.DirectTextBoxesCopied -ne 1 -or @($copy.CopiedChildPaths).Count -ne 2 -or
            $afterCopy.NodeCount -ne ($before.NodeCount + 3) -or
            $afterCopy.TreeVersion -eq $before.TreeVersion) {
            throw 'Frame with Label/TextBox was not copied and read back as expected.'
        }

        Invoke-Vbe @{ Command = 'add_nested_form_control'; Project = $Project; Form = $Form;
            ParentPath = 'Controls/fraOriginal'; ControlType = 'Forms.CheckBox.1';
            Control = 'chkUnsupported'; Caption = 'Autre'; Left = 8; Top = 78; Width = 80; Height = 20;
            ExpectedTreeVersion = $afterCopy.TreeVersion } | Out-Null
        $beforeRefusal = Read-Tree
        $refusal = Invoke-VbeRaw @{ Command = 'duplicate_form_frame_simple_children'; Project = $Project; Form = $Form;
            ControlPath = 'Controls/fraOriginal'; NewName = 'fraRejected';
            ExpectedTreeVersion = $beforeRefusal.TreeVersion }
        $afterRefusal = Read-Tree
        if ($refusal.Ok -or $refusal.Error -notmatch 'ineligible' -or
            $afterRefusal.TreeVersion -ne $beforeRefusal.TreeVersion -or
            $afterRefusal.NodeCount -ne $beforeRefusal.NodeCount) {
            throw 'Unsupported CheckBox child was not refused before mutation.'
        }

        [pscustomobject]@{
            HostProcessId = $HostProcessId
            Form = $Form
            CopiedFrame = $copy.NewPath
            CopiedLabels = $copy.DirectLabelsCopied
            CopiedTextBoxes = $copy.DirectTextBoxesCopied
            CopiedPaths = ($copy.CopiedChildPaths -join ', ')
            BeforeNodeCount = $before.NodeCount
            AfterCopyNodeCount = $afterCopy.NodeCount
            UnsupportedChildRefusal = $refusal.Error
            RefusalLeftTreeUnchanged = ($afterRefusal.TreeVersion -eq $beforeRefusal.TreeVersion)
            ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
        } | Format-List
    }
    'Profiled' {
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
    }
}

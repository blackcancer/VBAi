param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('AddControlRollback','AddNestedRollback','ComboColumnGuard','ControlEnumChoices',
        'PropertyStatus','SingleNodeSetter','SpinProperties','SpinWriteGuard','FrameCopyPlan',
        'FrameDuplication','ControlDuplication')]
    [string] $Scenario,
    [Parameter(Mandatory = $true)] [ValidateRange(1, [int]::MaxValue)] [int] $HostProcessId,
    [string] $Project = 'VBAProject',
    [string] $Form,
    [string] $ExpectedCodeBase,
    [string] $ControlType,
    [string] $Property,
    [string] $ValueJson,
    [string] $Control = 'probeControl',
    [ValidateSet('Empty','Labels','SimpleChildren','Profiled')] [string] $FrameScenario = 'Empty',
    [switch] $OwnedDisposableWorkbook
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'VbeProbe.Common.ps1')
$scenarioParameters = @{
    'SingleNodeSetter' = @{ Allowed = @('HostProcessId','ControlType','Property','ValueJson','Form','Control','OwnedDisposableWorkbook'); Required = @('HostProcessId','ControlType','Property','ValueJson','OwnedDisposableWorkbook') }
    'AddControlRollback' = @{ Allowed = @('HostProcessId','OwnedDisposableWorkbook'); Required = @('HostProcessId','OwnedDisposableWorkbook') }
    'ComboColumnGuard' = @{ Allowed = @('HostProcessId','OwnedDisposableWorkbook'); Required = @('HostProcessId','OwnedDisposableWorkbook') }
    'ControlDuplication' = @{ Allowed = @('ControlType','HostProcessId','Project','Form','OwnedDisposableWorkbook'); Required = @('ControlType','HostProcessId','OwnedDisposableWorkbook') }
    'FrameDuplication' = @{ Allowed = @('HostProcessId','FrameScenario','Project','Form','OwnedDisposableWorkbook'); Required = @('HostProcessId','OwnedDisposableWorkbook') }
    'PropertyStatus' = @{ Allowed = @('HostProcessId','ExpectedCodeBase','Form','OwnedDisposableWorkbook'); Required = @('HostProcessId','ExpectedCodeBase','OwnedDisposableWorkbook') }
    'AddNestedRollback' = @{ Allowed = @('HostProcessId','OwnedDisposableWorkbook'); Required = @('HostProcessId','OwnedDisposableWorkbook') }
    'ControlEnumChoices' = @{ Allowed = @('HostProcessId','Form','OwnedDisposableWorkbook'); Required = @('HostProcessId','OwnedDisposableWorkbook') }
    'SpinProperties' = @{ Allowed = @('HostProcessId','Project','Form','OwnedDisposableWorkbook'); Required = @('HostProcessId','OwnedDisposableWorkbook') }
    'SpinWriteGuard' = @{ Allowed = @('HostProcessId','ExpectedCodeBase','Project','Form','OwnedDisposableWorkbook'); Required = @('HostProcessId','ExpectedCodeBase','OwnedDisposableWorkbook') }
    'FrameCopyPlan' = @{ Allowed = @('HostProcessId','Project','Form','OwnedDisposableWorkbook'); Required = @('HostProcessId','OwnedDisposableWorkbook') }
}
$contract = $scenarioParameters[$Scenario]
Assert-VbeProbeParameters -Bound $PSBoundParameters -Required $contract.Required -Allowed $contract.Allowed
if (-not $OwnedDisposableWorkbook) {
    throw 'Confirm an owned disposable Excel workbook with -OwnedDisposableWorkbook.'
}
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') {
    throw "The probe requires Excel; PID $HostProcessId is $($hostProcess.ProcessName)."
}
$VbeProbeJsonDepth = 8
$VbeProbeRawError = $false
$VbeProbeTraceRequests = $false

switch ($Scenario) {
    'AddControlRollback' {
        $ErrorActionPreference = 'Stop'

        $projects = @(Invoke-Vbe @{ Command = 'list_projects' })
        if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) { throw 'Expected one disposable project in design mode.' }
        $project = $projects[0].Name
        $form = 'CodexAddAtomicProbe'
        if (-not @(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } | Where-Object { $_.Name -eq $form }).Count) {
            Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form } | Out-Null
        }
        $initial = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
        if ($initial.NodeCount -ne 0) { throw "Form $form must be empty before this probe." }
        $invalid = Invoke-Reply @{ Command = 'add_form_control'; Project = $project; Form = $form;
            Control = 'imgRejected'; ControlType = 'Forms.Image.1'; Left = 24; Top = 24;
            Width = 90; Height = 60; Caption = 'Unsupported'; ExpectedFormVersion = $initial.FormVersion }
        if ($invalid.Ok -or $invalid.Error -notmatch 'does not expose Caption') {
            throw "Image.Caption did not produce a verified rollback: $($invalid.Error)"
        }
        $afterRefusal = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
        if ($afterRefusal.NodeCount -ne 0 -or
            $afterRefusal.TreeVersion -ne $initial.TreeVersion -or
            @($afterRefusal.Controls | Where-Object { $_.Name -eq 'imgRejected' }).Count) {
            throw 'The rejected Image changed form_tree or remains in the form.'
        }
        $image = Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
            Control = 'imgAccepted'; ControlType = 'Forms.Image.1'; Left = 24; Top = 24;
            Width = 90; Height = 60; ExpectedFormVersion = $afterRefusal.FormVersion }
        $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
        $imageNode = @($tree.Controls | Where-Object { $_.Name -eq 'imgAccepted' })[0]
        if (-not $imageNode -or $imageNode.Type -ne 'Image' -or $tree.NodeCount -ne 1) {
            throw 'Image without Caption was not created and read back.'
        }
        $label = Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
            Control = 'lblAccepted'; ControlType = 'Forms.Label.1'; Left = 24; Top = 100;
            Width = 90; Height = 30; Caption = 'Visible'; ExpectedFormVersion = $tree.FormVersion }
        $final = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
        $labelNode = @($final.Controls | Where-Object { $_.Name -eq 'lblAccepted' })[0]
        $caption = @($labelNode.Properties | Where-Object { $_.Name -eq 'Caption' })[0].Value
        if ($final.NodeCount -ne 2 -or $caption -ne 'Visible') {
            throw 'Label Caption did not survive the property preflight.'
        }
        [pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project; Form = $form;
            UnsupportedCaptionRefused = $true; RejectedControlAbsent = $true;
            ImageWithoutCaptionCreated = $true; LabelCaptionReadback = $caption;
            FinalNodeCount = $final.NodeCount;
            ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue) } | ConvertTo-Json -Depth 5
    }
    'AddNestedRollback' {
        $ErrorActionPreference = 'Stop'

        $projects = @(Invoke-Vbe @{ Command = 'list_projects' })
        if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) { throw 'Expected one disposable design-mode project.' }
        $project = $projects[0].Name
        $form = 'CodexNestedAtomicProbe'
        if (@(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } | Where-Object { $_.Name -eq $form }).Count) {
            throw "Form $form already exists."
        }
        Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form } | Out-Null
        $initial = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
        Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
            Control = 'fraProbe'; ControlType = 'Forms.Frame.1'; Caption = 'Parent';
            Left = 20; Top = 20; Width = 170; Height = 130;
            ExpectedFormVersion = $initial.FormVersion } | Out-Null
        $before = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
        $invalid = Invoke-Reply @{ Command = 'add_nested_form_control'; Project = $project; Form = $form;
            ParentPath = 'Controls/fraProbe'; Control = 'imgRejected'; ControlType = 'Forms.Image.1';
            Left = 10; Top = 20; Width = 60; Height = 50; Caption = 'Unsupported';
            ExpectedTreeVersion = $before.TreeVersion }
        if ($invalid.Ok -or $invalid.Error -notmatch 'does not expose Caption') {
            throw "Nested Image.Caption did not fail as expected: $($invalid.Error)"
        }
        $afterRefusal = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
        if ($afterRefusal.NodeCount -ne 1 -or $afterRefusal.TreeVersion -ne $before.TreeVersion -or
            @($afterRefusal.Controls[0].Children | Where-Object { $_.Name -eq 'imgRejected' }).Count) {
            throw 'The rejected nested Image changed form_tree or remains in its Frame.'
        }
        Invoke-Vbe @{ Command = 'add_nested_form_control'; Project = $project; Form = $form;
            ParentPath = 'Controls/fraProbe'; Control = 'imgAccepted'; ControlType = 'Forms.Image.1';
            Left = 10; Top = 20; Width = 60; Height = 50;
            ExpectedTreeVersion = $afterRefusal.TreeVersion } | Out-Null
        $final = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
        if ($final.NodeCount -ne 2 -or
            -not @($final.Controls[0].Children | Where-Object { $_.Name -eq 'imgAccepted' }).Count) {
            throw 'Nested Image without Caption was not created and read back.'
        }
        [pscustomobject]@{ HostProcessId = $HostProcessId; Form = $form;
            UnsupportedCaptionRefused = $true; RejectedControlAbsent = $true;
            TreeVersionRestored = $true; NestedImageCreated = $true;
            FinalNodeCount = $final.NodeCount;
            ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue) } | ConvertTo-Json
    }
    'ComboColumnGuard' {
        $ErrorActionPreference = 'Stop'

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
        $reply = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
        if ($reply.Ok -or $reply.Error -notmatch 'temporarily disabled') {
            throw 'ComboBox.ColumnCount was not blocked before COM.'
        }
        $after = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
        if ($before.TreeVersion -ne $after.TreeVersion) { throw 'The refused setter changed TreeVersion.' }
        [pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project; Form = $form;
            SetterStatus = $property.SetterStatus; Refused = $true;
            TreeVersionStable = $true; ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue) } | ConvertTo-Json
    }
    'ControlEnumChoices' {
        if (-not $Form) { $Form = 'CodexEnumChoicesSurvey' }
        $ErrorActionPreference = 'Stop'


        $projects = @(Invoke-Vbe @{ Command = 'list_projects' })
        if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) {
            throw 'Expected exactly one disposable VBA project in design mode.'
        }
        $project = $projects[0].Name
        if (@(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } | Where-Object { $_.Name -eq $Form }).Count) {
            throw "Form $Form already exists."
        }

        Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $Form } | Out-Null
        $types = @('CheckBox', 'ComboBox', 'CommandButton', 'Frame', 'Image', 'Label', 'ListBox',
            'MultiPage', 'OptionButton', 'ScrollBar', 'SpinButton', 'TabStrip', 'TextBox', 'ToggleButton')
        $total = 0
        $enums = 0
        $missing = @()
        foreach ($index in 0..($types.Count - 1)) {
            $type = $types[$index]
            $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $Form }
            $name = 'probe' + $type
            Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $Form;
                Control = $name; ControlType = "Forms.$type.1"; ExpectedFormVersion = $tree.FormVersion;
                Left = (12 + (($index % 2) * 155)); Top = (12 + ([math]::Floor($index / 2) * 42));
                Width = 140; Height = 32 } | Out-Null
            $properties = @(Invoke-Vbe @{ Command = 'form_control_properties'; Project = $project;
                    Form = $Form; Control = $name })
            $total += $properties.Count
            foreach ($property in $properties) {
                if ($property.Type -match '_fm[A-Za-z0-9]+$') {
                    $enums++
                    if (@($property.AllowedValues).Count -eq 0) {
                        $missing += "$type.$($property.Name)"
                    }
                }
            }
        }
        $first = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $Form }
        $second = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $Form }
        if ($first.NodeCount -ne 18 -or $second.TreeVersion -ne $first.TreeVersion -or $missing.Count) {
            throw "Enum catalog unstable or incomplete: NodeCount=$($first.NodeCount), Missing=$($missing -join ', ')"
        }
        $label = @($first.Controls | Where-Object { $_.Name -eq 'probeLabel' })[0]
        $textAlign = @($label.Properties | Where-Object { $_.Name -eq 'TextAlign' })[0]
        if (@($textAlign.AllowedValues).Count -lt 3) {
            throw 'Label.TextAlign enum choices were not exposed by form_tree.'
        }
        [pscustomobject]@{
            HostProcessId = $HostProcessId
            Project = $project
            Form = $Form
            ControlTypes = $types.Count
            NodeCount = $first.NodeCount
            PropertyRows = $total
            EnumProperties = $enums
            EnumChoicesMissing = $missing.Count
            LabelTextAlignChoices = (@($textAlign.AllowedValues) -join ', ')
            TreeVersionStable = ($second.TreeVersion -eq $first.TreeVersion)
            ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
        } | Format-List
    }
    'PropertyStatus' {
        if (-not $Form) { $Form = 'CodexPropertyStatusSurvey' }
        if ([string]::IsNullOrWhiteSpace($ExpectedCodeBase)) { throw 'ExpectedCodeBase is required for this scenario.' }
        $ErrorActionPreference = 'Stop'
        $actualCodeBase = (Get-ItemProperty -LiteralPath 'Registry::HKEY_CURRENT_USER\Software\Classes\CLSID\{8E854243-087F-4D6C-9E0E-8622B0E50883}\InprocServer32').CodeBase
        if ($actualCodeBase -ne $ExpectedCodeBase) { throw 'Refusing probe: COM CodeBase differs from expected build.' }



        $projects = @(Invoke-Vbe @{ Command = 'list_projects' })
        if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) {
            throw 'Expected exactly one disposable VBA project in design mode.'
        }
        $project = $projects[0].Name
        if (@(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } | Where-Object { $_.Name -eq $Form }).Count) {
            throw "Form $Form already exists."
        }
        Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $Form } | Out-Null
        foreach ($item in @(
                @{ Type = 'Label'; Name = 'lblProbe'; Left = 12 },
                @{ Type = 'SpinButton'; Name = 'spnProbe'; Left = 100 },
                @{ Type = 'ToggleButton'; Name = 'tglProbe'; Left = 160 },
                @{ Type = 'TextBox'; Name = 'txtProbe'; Left = 220 }
        )) {
            $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $Form }
            Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $Form;
                Control = $item.Name; ControlType = "Forms.$($item.Type).1";
                Left = $item.Left; Top = 24; Width = 50; Height = 30;
                ExpectedFormVersion = $tree.FormVersion } | Out-Null
        }

        $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $Form }
        $checks = @(
            @{ Control = 'lblProbe'; Property = 'Cancel'; Expected = 'BlockedNativeSetterFailure' },
            @{ Control = 'lblProbe'; Property = '_Font_Reserved'; Expected = 'GetterUnavailable' },
            @{ Control = 'tglProbe'; Property = 'Value'; Expected = 'BlockedAfterHostCrash' },
            @{ Control = 'txtProbe'; Property = 'ScrollBars'; Expected = 'BlockedAfterHostCrash' },
            @{ Control = 'spnProbe'; Property = 'Min'; Expected = 'BlockedAfterHostCrash' },
            @{ Control = 'spnProbe'; Property = 'Max'; Expected = 'BlockedAfterHostCrash' },
            @{ Control = 'spnProbe'; Property = 'Value'; Expected = 'BlockedAfterHostCrash' },
            @{ Control = 'spnProbe'; Property = 'Delay'; Expected = 'BlockedAfterHostCrash' },
            @{ Control = 'spnProbe'; Property = 'SmallChange'; Expected = 'BlockedAfterHostCrash' }
        )
        foreach ($check in $checks) {
            $node = @($tree.Controls | Where-Object { $_.Name -eq $check.Control })[0]
            $property = @($node.Properties | Where-Object { $_.Name -eq $check.Property })[0]
            $flat = @(Invoke-Vbe @{ Command = 'form_control_properties'; Project = $project;
                    Form = $Form; Control = $check.Control } | Where-Object { $_.Name -eq $check.Property })[0]
            if ($property.SetterStatus -ne $check.Expected -or $flat.SetterStatus -ne $check.Expected) {
                throw "$($check.Control).$($check.Property) status differs between form_tree and form_control_properties."
            }
        }
        $blocked = Invoke-VbeRaw @{ Command = 'set_form_node_property'; Project = $project; Form = $Form;
            ControlPath = 'Controls/lblProbe'; Property = '_Font_Reserved'; Value = 'unsafe';
            ExpectedTreeVersion = $tree.TreeVersion }
        $after = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $Form }
        if ($blocked.Ok -or $blocked.Error -notmatch 'reserved member' -or
            $after.TreeVersion -ne $tree.TreeVersion) {
            throw 'The reserved Font member was not refused before mutation.'
        }
        $scrollBlocked = Invoke-VbeRaw @{ Command = 'set_form_node_property'; Project = $project; Form = $Form;
            ControlPath = 'Controls/txtProbe'; Property = 'ScrollBars'; Value = 'Vertical';
            ExpectedTreeVersion = $tree.TreeVersion }
        $afterScroll = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $Form }
        if ($scrollBlocked.Ok -or $scrollBlocked.Error -notmatch 'temporarily disabled' -or
            $afterScroll.TreeVersion -ne $tree.TreeVersion) {
            throw 'TextBox.ScrollBars was not refused before mutation.'
        }

        [pscustomobject]@{
            HostProcessId = $HostProcessId
            CheckedStatuses = $checks.Count
            FormTreeAndControlPropertiesAgree = $true
            ReservedMemberRefused = $true
            ScrollBarsRefused = $true
            TreeVersionStable = ($afterScroll.TreeVersion -eq $tree.TreeVersion)
            ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
        } | Format-List
    }
    'SingleNodeSetter' {
        if (-not $Form) { $Form = 'CodexSingleSetterSurvey' }
        if ([string]::IsNullOrWhiteSpace($ControlType) -or [string]::IsNullOrWhiteSpace($Property) -or [string]::IsNullOrWhiteSpace($ValueJson)) { throw 'ControlType, Property and ValueJson are required.' }
        $ErrorActionPreference = 'Stop'
        $value = ConvertFrom-Json -InputObject $ValueJson


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
    }
    'SpinProperties' {
        if (-not $Form) { $Form = 'CodexSpinPropertySurvey' }
        $ErrorActionPreference = 'Stop'


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
    }
    'SpinWriteGuard' {
        if (-not $Form) { $Form = 'CodexSpinGuardSurvey' }
        if ([string]::IsNullOrWhiteSpace($ExpectedCodeBase)) { throw 'ExpectedCodeBase is required for this scenario.' }
        $ErrorActionPreference = 'Stop'
        $actualCodeBase = (Get-ItemProperty -LiteralPath 'Registry::HKEY_CURRENT_USER\Software\Classes\CLSID\{8E854243-087F-4D6C-9E0E-8622B0E50883}\InprocServer32').CodeBase
        if ($actualCodeBase -ne $ExpectedCodeBase) {
            throw "Refusing any write probe: COM CodeBase differs from guarded build. Actual=$actualCodeBase"
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
            ControlType = 'Forms.SpinButton.1'; Control = 'spnProbe';
            Left = 24; Top = 24; Width = 32; Height = 80;
            ExpectedFormVersion = $initial.FormVersion } | Out-Null
        $before = Read-Tree
        $refused = @()
        foreach ($change in @(
                @{ Name = 'Min'; Value = 2 },
                @{ Name = 'Max'; Value = 20 },
                @{ Name = 'Value'; Value = 7 },
                @{ Name = 'Delay'; Value = 100 },
                @{ Name = 'SmallChange'; Value = 2 }
        )) {
            $reply = Invoke-VbeRaw @{ Command = 'set_form_node_property'; Project = $Project; Form = $Form;
                ControlPath = 'Controls/spnProbe'; Property = $change.Name; Value = $change.Value;
                ExpectedTreeVersion = $before.TreeVersion }
            if ($reply.Ok -or $reply.Error -notmatch 'temporarily disabled') {
                throw "SpinButton.$($change.Name) was not refused before COM mutation."
            }
            $after = Read-Tree
            if ($after.TreeVersion -ne $before.TreeVersion -or $after.NodeCount -ne $before.NodeCount) {
                throw "SpinButton.$($change.Name) refusal changed the form tree."
            }
            $refused += $change.Name
        }

        [pscustomobject]@{
            HostProcessId = $HostProcessId
            GuardedCodeBase = $actualCodeBase
            Refused = ($refused -join ', ')
            TreeUnchanged = $true
            ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
        } | Format-List
    }
    'FrameCopyPlan' {
        if (-not $Form) { $Form = 'CodexFramePlanSurvey' }
        $ErrorActionPreference = 'Stop'


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
    }
    'FrameDuplication' {
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
        if (-not $Form) { $Form = $defaultForms[$FrameScenario] }
        if ([string]::IsNullOrWhiteSpace($Form)) { throw 'A nonempty disposable form name is required.' }




        function Read-Tree {
            Invoke-Vbe @{ Command = 'form_tree'; Project = $Project; Form = $Form }
        }

        $existing = @(Invoke-Vbe @{ Command = 'list_forms'; Project = $Project })
        if (@($existing | Where-Object { $_.Name -eq $Form }).Count) {
            throw "Form $Form already exists; use a new name in a disposable workbook."
        }

        Invoke-Vbe @{ Command = 'create_form'; Project = $Project; Form = $Form } | Out-Null
        $initial = Read-Tree
        switch ($FrameScenario) {
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
    }
    'ControlDuplication' {
        if ($ControlType -notin @('Label','TextBox','ComboBox','CheckBox','OptionButton','ToggleButton','CommandButton')) { throw 'Select one of the seven supported ControlType values.' }
        $ErrorActionPreference = 'Stop'
        if (-not $OwnedDisposableWorkbook) {
            throw 'Confirm that the selected Excel process contains an owned disposable workbook with -OwnedDisposableWorkbook.'
        }

        # Each entry preserves the command, seed values, revision sequence, and extra
        # response oracle from its original single-control probe.
        $definitions = @{
            Label = @{ Form = 'CodexLabelCopySurvey'; ProgId = 'Forms.Label.1'; Source = 'lblOriginal'; Copy = 'lblCopy';
                Caption = 'Label source'; Width = 120; Property = 'BackColor'; Value = '#FF8800';
                Duplicate = 'duplicate_form_label'; Error = 'DuplicateLabel did not return the expected partial copy and new tree revision.' }
            TextBox = @{ Form = 'CodexTextBoxCopySurvey'; ProgId = 'Forms.TextBox.1'; Source = 'txtOriginal'; Copy = 'txtCopy';
                Width = 120; Property = 'Value'; Value = 'Texte Codex';
                Duplicate = 'duplicate_form_textbox'; Error = 'DuplicateTextBox did not return the expected partial copy and new tree revision.' }
            ComboBox = @{ Form = 'CodexComboCopySurvey'; ProgId = 'Forms.ComboBox.1'; Source = 'cmbOriginal'; Copy = 'cmbCopy';
                Width = 120; Property = 'ListWidth'; Value = '72 pt'; ExtraFalse = @('ItemsCopied', 'BindingsCopied');
                Duplicate = 'duplicate_form_combobox'; Error = 'DuplicateComboBox did not return the expected partial copy and new tree revision.' }
            CheckBox = @{ Form = 'CodexCheckBoxCopySurvey'; ProgId = 'Forms.CheckBox.1'; Source = 'chkOriginal'; Copy = 'chkCopy';
                Caption = 'Choix Codex'; Width = 120; Property = 'Value'; Value = $true;
                Duplicate = 'duplicate_form_checkbox'; Error = 'DuplicateCheckBox did not return the expected partial copy and new tree revision.' }
            OptionButton = @{ Form = 'CodexOptionCopySurvey'; ProgId = 'Forms.OptionButton.1'; Source = 'optOriginal'; Copy = 'optCopy';
                Caption = 'Premier choix'; Width = 130; ExtraFalse = @('SelectionCopied', 'GroupCopied');
                Duplicate = 'duplicate_form_optionbutton'; Error = 'DuplicateOptionButton did not return the expected partial copy and new tree revision.' }
            ToggleButton = @{ Form = 'CodexToggleButtonCopySurvey'; ProgId = 'Forms.ToggleButton.1'; Source = 'tglOriginal'; Copy = 'tglCopy';
                Caption = 'Choix Codex'; Width = 120; RefuseValueWrite = $true;
                Duplicate = 'duplicate_form_togglebutton'; Error = 'DuplicateToggleButton did not return the expected partial copy and new tree revision.' }
            CommandButton = @{ Form = 'CodexButtonCopySurvey'; ProgId = 'Forms.CommandButton.1'; Source = 'cmdOriginal'; Copy = 'cmdCopy';
                Caption = 'Exécuter'; Width = 120; ExtraFalse = @('EventsCopied');
                Duplicate = 'duplicate_form_commandbutton'; Error = 'DuplicateCommandButton did not return the expected partial copy and new tree revision.' }
        }
        $definition = $definitions[$ControlType]
        if (-not $Form) { $Form = $definition.Form }
        if ([string]::IsNullOrWhiteSpace($Form)) { throw 'A nonempty disposable form name is required.' }



        function Read-Tree {
            Invoke-Vbe @{ Command = 'form_tree'; Project = $Project; Form = $Form }
        }

        $existing = @(Invoke-Vbe @{ Command = 'list_forms'; Project = $Project })
        if (@($existing | Where-Object { $_.Name -eq $Form }).Count) {
            throw "Form $Form already exists; use a new name in a disposable workbook."
        }

        Invoke-Vbe @{ Command = 'create_form'; Project = $Project; Form = $Form } | Out-Null
        $initial = Read-Tree
        $add = @{ Command = 'add_form_control'; Project = $Project; Form = $Form;
            ControlType = $definition.ProgId; Control = $definition.Source;
            Left = 24; Top = 24; Width = $definition.Width; Height = 30;
            ExpectedFormVersion = $initial.FormVersion }
        if ($definition.ContainsKey('Caption')) { $add.Caption = $definition.Caption }
        Invoke-Vbe $add | Out-Null

        if ($definition.ContainsKey('Property')) {
            $withControl = Read-Tree
            Invoke-Vbe @{ Command = 'set_form_node_property'; Project = $Project; Form = $Form;
                ControlPath = ('Controls/' + $definition.Source); Property = $definition.Property;
                Value = $definition.Value; ExpectedTreeVersion = $withControl.TreeVersion } | Out-Null
        }
        $before = Read-Tree

        $refusedReply = $null
        if ($definition.RefuseValueWrite) {
            $refused = @{ Command = 'set_form_node_property'; Project = $Project; Form = $Form;
                ControlPath = ('Controls/' + $definition.Source); Property = 'Value'; Value = $true;
                ExpectedTreeVersion = $before.TreeVersion }
            $refusedJson = ConvertTo-Json -InputObject $refused -Compress -Depth 8
            $refusedReply = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $refusedJson | ConvertFrom-Json
            $afterRefusal = Read-Tree
            if ($refusedReply.Ok -or $refusedReply.Error -notmatch 'temporarily disabled' -or
                $afterRefusal.TreeVersion -ne $before.TreeVersion) {
                throw 'ToggleButton.Value=true was not refused before mutation.'
            }
        }

        $copy = Invoke-Vbe @{ Command = $definition.Duplicate; Project = $Project; Form = $Form;
            ControlPath = ('Controls/' + $definition.Source); NewName = $definition.Copy;
            ExpectedTreeVersion = $before.TreeVersion }
        $after = Read-Tree
        $extraFailed = $false
        $extraProperties = if ($definition.ContainsKey('ExtraFalse')) { @($definition.ExtraFalse) } else { @() }
        foreach ($property in $extraProperties) {
            if ($copy.$property -ne $false) { $extraFailed = $true }
        }
        if ($copy.Completeness -ne 'Partial' -or $copy.NewPath -ne ('Controls/' + $definition.Copy) -or
            $extraFailed -or $after.NodeCount -ne ($before.NodeCount + 1) -or
            $after.TreeVersion -eq $before.TreeVersion) {
            throw $definition.Error
        }

        $result = [ordered]@{
            HostProcessId = $HostProcessId
            Form = $Form
            SourcePath = $copy.SourcePath
            NewPath = $copy.NewPath
            Completeness = $copy.Completeness
        }
        foreach ($property in $extraProperties) { $result[$property] = $copy.$property }
        $result.CopiedProperties = ($copy.CopiedProperties -join ', ')
        if ($definition.RefuseValueWrite) { $result.ValueWriteRefused = $refusedReply.Error }
        if ($ControlType -ne 'OptionButton') {
            $result.BeforeTreeVersion = $before.TreeVersion
            $result.AfterTreeVersion = $after.TreeVersion
        }
        $result.BeforeNodeCount = $before.NodeCount
        $result.AfterNodeCount = $after.NodeCount
        $result.ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
        [pscustomobject]$result | Format-List
    }
}

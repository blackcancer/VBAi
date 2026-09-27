param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [Parameter(Mandatory = $true)] [string] $ExpectedCodeBase,
    [string] $Form = 'CodexPropertyStatusSurvey'
)

$ErrorActionPreference = 'Stop'
$actualCodeBase = (Get-ItemProperty -LiteralPath 'Registry::HKEY_CURRENT_USER\Software\Classes\CLSID\{8E854243-087F-4D6C-9E0E-8622B0E50883}\InprocServer32').CodeBase
if ($actualCodeBase -ne $ExpectedCodeBase) { throw 'Refusing probe: COM CodeBase differs from expected build.' }
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') { throw 'The probe requires Excel.' }

function Invoke-VbeRaw([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
}

function Invoke-Vbe([hashtable] $Request) {
    $reply = Invoke-VbeRaw $Request
    if (-not $reply.Ok) { throw "$($Request.Command): $($reply.Error)" }
    return $reply.Data
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

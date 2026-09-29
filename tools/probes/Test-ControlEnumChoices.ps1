param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [string] $Form = 'CodexEnumChoicesSurvey'
)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') { throw 'The probe requires Excel.' }

function Invoke-Vbe([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $reply = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
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

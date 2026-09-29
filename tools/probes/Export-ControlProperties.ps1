param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$client = Join-Path $repo 'tools/Invoke-VBAi.ps1'
$output = Join-Path $repo 'docs/reference/excel-control-properties.csv'

function Invoke-Vbe([hashtable] $request) {
    $raw = & $client -HostProcessId $HostProcessId -RequestJson (ConvertTo-Json -InputObject $request -Depth 8 -Compress)
    $response = $raw | ConvertFrom-Json
    if (-not $response.Ok) { throw "$($request.Command): $($response.Error)" }
    return $response.Data
}

$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) {
    throw 'Expected one disposable VBA project in design mode.'
}
$project = $projects[0].Name
$form = 'CodexAllControlProperties'
$state = Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form }
foreach ($dimension in @(@{ Property = 'Width'; Value = 330 }, @{ Property = 'Height'; Value = 360 })) {
    $changed = Invoke-Vbe @{
        Command = 'set_form_property'; Project = $project; Form = $form;
        ExpectedFormVersion = $state.Version; Property = $dimension.Property; Value = $dimension.Value
    }
    $state = $changed.State
}

$types = @(
    'CheckBox', 'ComboBox', 'CommandButton', 'Frame', 'Image', 'Label', 'ListBox',
    'MultiPage', 'OptionButton', 'ScrollBar', 'SpinButton', 'TabStrip', 'TextBox', 'ToggleButton'
)
$rows = New-Object 'System.Collections.Generic.List[object]'
for ($index = 0; $index -lt $types.Count; $index++) {
    $type = $types[$index]
    $name = 'probe' + $type
    $state = Invoke-Vbe @{
        Command = 'add_form_control'; Project = $project; Form = $form;
        Control = $name; ControlType = "Forms.$type.1"; ExpectedFormVersion = $state.Version;
        Left = 12 + (($index % 2) * 155); Top = 12 + ([math]::Floor($index / 2) * 42);
        Width = 140; Height = 32
    }
    foreach ($property in @(Invoke-Vbe @{
        Command = 'form_control_properties'; Project = $project; Form = $form; Control = $name
    })) {
        $rows.Add([pscustomobject]@{
            ControlType = $type; ProgId = "Forms.$type.1"; Property = $property.Name;
            Type = $property.Type; ReadOnly = $property.ReadOnly;
            Value = $property.Value; Error = $property.Error
        })
    }
}
$rows | Sort-Object ControlType,Property | Export-Csv -LiteralPath $output -Encoding UTF8 -NoTypeInformation
Write-Output "Form=$form; Controls=$($types.Count); PropertyRows=$($rows.Count); Output=$output"

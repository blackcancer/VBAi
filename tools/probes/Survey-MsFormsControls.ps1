param([Parameter(Mandatory = $true)] [int] $HostProcessId,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $Project,
    [switch] $AcknowledgeDisposableProject)

$ErrorActionPreference = 'Stop'
if (-not $AcknowledgeDisposableProject) { throw 'Confirm the selected host contains only an owned disposable project; this probe adds a form and controls.' }
$client = Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-VBAi.ps1'
function Invoke-Vbe([hashtable] $request) {
    $raw = & $client -HostProcessId $HostProcessId -RequestJson (ConvertTo-Json -InputObject $request -Depth 8 -Compress)
    $response = $raw | ConvertFrom-Json
    if (-not $response.Ok) { throw "$($request.Command): $($response.Error)" }
    return $response.Data
}

$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2 -or $projects[0].Name -cne $Project) {
    throw 'Expected exactly the explicitly named disposable VBA project in design mode.'
}
$form = 'CodexControlSurvey2'
$state = Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form }
Write-Output "FORM $($state.Form) Width=$($state.Width) Height=$($state.Height)"

$types = @(
    'Forms.OptionButton.1', 'Forms.ListBox.1', 'Forms.SpinButton.1',
    'Forms.ScrollBar.1', 'Forms.Image.1', 'Forms.MultiPage.1',
    'Forms.TabStrip.1', 'Forms.ToggleButton.1'
)
for ($index = 0; $index -lt $types.Count; $index++) {
    $type = $types[$index]
    $name = 'probe' + ($index + 1)
    try {
        $state = Invoke-Vbe @{
            Command = 'add_form_control'; Project = $project; Form = $form;
            Control = $name; ControlType = $type; ExpectedFormVersion = $state.Version;
            Left = 12 + (($index % 2) * 120); Top = 12 + ([math]::Floor($index / 2) * 36);
            Width = 105; Height = 28
        }
        $properties = @(Invoke-Vbe @{
            Command = 'form_control_properties'; Project = $project; Form = $form; Control = $name
        })
        $errors = @($properties | Where-Object { $_.Error })
        $writable = @($properties | Where-Object { -not $_.ReadOnly })
        $names = @($properties | ForEach-Object Name)
        Write-Output "CONTROL $type Added=$name Properties=$($properties.Count) Writable=$($writable.Count) Errors=$($errors.Count) HasCaption=$($names -contains 'Caption') HasValue=$($names -contains 'Value') HasPicture=$($names -contains 'Picture')"
        foreach ($propertyError in $errors) { Write-Output "  ERROR $($propertyError.Name): $($propertyError.Error)" }
    }
    catch { Write-Output "CONTROL $type FAILED $($_.Exception.Message)" }
}
Write-Output "FINAL Controls=$($state.Controls.Count) Version=$($state.Version)"

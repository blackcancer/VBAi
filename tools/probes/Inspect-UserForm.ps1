param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Explore','Nested','ExportProperties','SurveyControls','FormProperties')]
    [string] $Scenario,
    [int] $HostProcessId,
    [string] $Project,
    [string] $Form,
    [ValidateSet('Inspect','Create','Read')] [string] $Stage = 'Inspect',
    [switch] $AcknowledgeDisposableProject,
    [switch] $OwnedDisposableWorkbook
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'VbeProbe.Common.ps1')
$scenarioParameters = @{
    'Explore' = @{ Allowed = @('Stage','OwnedDisposableWorkbook'); Required = @() }
    'FormProperties' = @{ Allowed = @('HostProcessId','Project','Form','OwnedDisposableWorkbook'); Required = @('HostProcessId','Project','Form') }
    'Nested' = @{ Allowed = @('OwnedDisposableWorkbook'); Required = @() }
    'SurveyControls' = @{ Allowed = @('HostProcessId','Project','AcknowledgeDisposableProject','OwnedDisposableWorkbook'); Required = @('HostProcessId','Project') }
    'ExportProperties' = @{ Allowed = @('HostProcessId','OwnedDisposableWorkbook'); Required = @('HostProcessId') }
}
$contract = $scenarioParameters[$Scenario]
Assert-VbeProbeParameters -Bound $PSBoundParameters -Required $contract.Required -Allowed $contract.Allowed
if ($Scenario -in @('ExportProperties','SurveyControls','FormProperties') -and $HostProcessId -le 0) {
    throw 'HostProcessId is required for this scenario.'
}
if ($Scenario -in @('SurveyControls','FormProperties') -and [string]::IsNullOrWhiteSpace($Project)) {
    throw 'Project is required for this scenario.'
}
if ($Scenario -eq 'FormProperties' -and [string]::IsNullOrWhiteSpace($Form)) {
    throw 'Form is required for this scenario.'
}
if ($Scenario -eq 'Explore' -and $Stage -eq 'Create' -and -not $OwnedDisposableWorkbook) {
    throw 'Creating a form requires -OwnedDisposableWorkbook.'
}
switch ($Scenario) {
    'Explore' {
        $ErrorActionPreference = 'Stop'
        if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit PowerShell.' }
        $hosts = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
        if ($hosts.Count -ne 1) { throw 'Expected one Excel process.' }
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        Write-Output "Excel COM=$($excel.GetType().FullName), Name=$($excel.Name), Workbooks=$($excel.Workbooks.Count)"
        $vbe = $excel.VBE
        if (-not $vbe) {
            Write-Output 'Excel.VBE returned null; trying ActiveWorkbook.VBProject.VBE.'
            $vbe = $excel.ActiveWorkbook.VBProject.VBE
        }
        if (-not $vbe) { throw 'The external Excel COM object did not expose the VBE.' }
        $project = $vbe.VBProjects.Item('VBAProject')
        Write-Output "Excel PID=$($hosts[0].Id), workbook=$($excel.ActiveWorkbook.Name), VBE visible=$($vbe.MainWindow.Visible), mode=$($project.Mode)"

        if ($Stage -eq 'Inspect') {
            foreach ($component in $project.VBComponents) {
                Write-Output "Component $($component.Name), type=$($component.Type)"
            }
            return
        }

        if ($Stage -eq 'Create') {
            foreach ($component in $project.VBComponents) {
                if ($component.Name -eq 'CodexFormProbe') { throw 'The probe form already exists.' }
            }
            $form = $project.VBComponents.Add(3)
            $form.Name = 'CodexFormProbe'
            Write-Output "Created $($form.Name), type=$($form.Type)"
        }
        else {
            $form = $project.VBComponents.Item('CodexFormProbe')
        }

        Write-Output "Designer type=$($form.Designer.GetType().FullName)"
        Write-Output "DesignerWindow type=$($form.DesignerWindow.GetType().FullName)"
        $form.DesignerWindow.Visible = $true
        Write-Output "Designer visible=$($form.DesignerWindow.Visible)"
        Write-Output "Form controls=$($form.Designer.Controls.Count)"
    }
    'Nested' {
        $ErrorActionPreference = 'Stop'
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        $vbe = $excel.VBE
        Write-Output "ExcelHwnd=$($excel.Hwnd) Workbooks=$($excel.Workbooks.Count)"
        Write-Output "Projects=$($vbe.VBProjects.Count)"
        Write-Output "Workbook=$($excel.Workbooks.Item(1).Name) WorkbookProject=$($excel.Workbooks.Item(1).VBProject.Name) WorkbookComponents=$($excel.Workbooks.Item(1).VBProject.VBComponents.Count)"
        foreach ($project in $vbe.VBProjects) {
            Write-Output "Project=$($project.Name) Mode=$($project.Mode)"
            foreach ($component in $project.VBComponents) {
                if ($component.Type -ne 3) { continue }
                Write-Output " Form=$($component.Name)"
                $designer = $component.Designer
                foreach ($control in $designer.Controls) {
                    Write-Output "  Control=$($control.Name) Type=$($control.GetType().FullName)"
                    foreach ($collectionName in @('Controls','Pages','Tabs')) {
                        try {
                            $collection = $control.$collectionName
                            Write-Output "   $collectionName Count=$($collection.Count)"
                            foreach ($item in $collection) {
                                $line = "    Item=$($item.Name)"
                                try { $line += " Caption=$($item.Caption)" } catch { }
                                try { $line += " Controls=$($item.Controls.Count)" } catch { }
                                Write-Output $line
                            }
                        }
                        catch { Write-Output "   $collectionName unavailable: $($_.Exception.Message)" }
                    }
                }
            }
        }
    }
    'ExportProperties' {
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
    }
    'SurveyControls' {
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
    }
    'FormProperties' {
        $ErrorActionPreference = 'Stop'
        $client = Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-VBAi.ps1'
        $request = @{ Command = 'form_properties'; Project = $Project; Form = $Form }
        $response = (& $client -HostProcessId $HostProcessId -RequestJson (ConvertTo-Json -InputObject $request -Compress)) | ConvertFrom-Json
        if (-not $response.Ok) { throw $response.Error }
        $response.Data | ConvertTo-Json -Depth 20
    }
}

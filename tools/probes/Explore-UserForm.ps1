param([ValidateSet('Inspect', 'Create', 'Read')] [string] $Stage = 'Inspect')

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

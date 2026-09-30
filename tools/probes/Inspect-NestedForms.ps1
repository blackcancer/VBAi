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

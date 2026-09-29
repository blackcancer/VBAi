param([switch] $NewInstance, [switch] $UseNewWorkbook)

$ErrorActionPreference = 'Stop'
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel must be closed for this isolated test.' }

$excelPath = 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE'
if (-not (Test-Path -LiteralPath $excelPath)) { throw "Excel was not found: $excelPath" }
$workbookPath = Join-Path $PSScriptRoot 'probes\VBAi-scratch.xlsx'
if (-not (Test-Path -LiteralPath $workbookPath)) { throw "Scratch workbook was not found: $workbookPath" }
$started = Get-Date
$arguments = if ($NewInstance) { @('/x', $workbookPath) } else { @($workbookPath) }
$process = Start-Process -FilePath $excelPath -ArgumentList $arguments -PassThru
Write-Output "Started native Excel PID=$($process.Id)"

$excel = $null
for ($attempt = 0; $attempt -lt 30 -and -not $excel; $attempt++) {
    Start-Sleep -Milliseconds 300
    try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
    catch { }
}
if (-not $excel) { throw 'Native Excel did not register its COM application object.' }

$excel.Visible = $true
$book = $excel.Workbooks.Item(1)
if ($UseNewWorkbook) {
    $book.Close($false)
    $book = $excel.Workbooks.Add()
}
$excel.CommandBars.ExecuteMso('VisualBasic')
Start-Sleep -Seconds 3
Write-Output "Workbook=$($book.Name)"
$log = Join-Path $env:TEMP 'VBAi-load.log'
if ((Test-Path -LiteralPath $log) -and (Get-Item -LiteralPath $log).LastWriteTime -ge $started) {
    Write-Output 'Fresh add-in log:'
    Get-Content -LiteralPath $log -Tail 8
}
else { Write-Output 'No fresh add-in log.' }
Write-Output "Excel left open for inspection: $($process.Id)"

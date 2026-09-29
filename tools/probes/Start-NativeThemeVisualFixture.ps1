param([Parameter(Mandatory=$true)][string]$StopFile)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Existing Excel session detected.' }
$env:VBAi_NATIVE_DARK_EXPERIMENT='1'
$excel=$null; $book=$null
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$true
 $book=$excel.Workbooks.Add()
 $book.VBProject.Name='NativeThemeVisualProbe'
 $excel.CommandBars.ExecuteMso('VisualBasic')
 'Fixture ready.'
 $timer=[Diagnostics.Stopwatch]::StartNew()
 while(-not (Test-Path -LiteralPath $StopFile) -and $timer.Elapsed.TotalMinutes -lt 12) { Start-Sleep -Milliseconds 500 }
} finally {
 if($excel) { $excel.VBE.MainWindow.Visible=$false }
 if($book) { $book.Close($false) }
 $book=$null
 if($excel) { $excel.Quit() }
 $excel=$null
 [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

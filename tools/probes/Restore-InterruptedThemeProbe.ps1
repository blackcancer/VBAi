param([Parameter(Mandatory=$true)][string]$RecoveryFile)
$ErrorActionPreference='Stop'
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel must be closed before recovery.' }
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$recovery=[IO.Path]::GetFullPath($RecoveryFile)
if (-not (Test-Path -LiteralPath $recovery)) { throw 'Palette recovery file not found.' }
$assembly=[Reflection.Assembly]::LoadFrom((Join-Path $root 'bin/Debug/net48/VBAi.dll'))
$change=$assembly.GetType('VBAi.VbeNativePalette').GetMethod('Change',[Reflection.BindingFlags]'NonPublic,Static')
$excel=$null; $book=$null
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$true
 $book=$excel.Workbooks.Add()
 $excel.CommandBars.ExecuteMso('VisualBasic')
 Start-Sleep -Seconds 3
 [void]$change.Invoke($null,[object[]]@($excel.VBE.PSObject.BaseObject,$false,[string]$recovery))
 'Restored and reopened verification succeeded.' | Set-Content ([IO.Path]::ChangeExtension($recovery, '.restored.txt'))
} finally {
 if($excel) { $excel.VBE.MainWindow.Visible=$false }
 if($book) { $book.Close($false) }
 $book=$null
 if($excel) { $excel.Quit() }
 $excel=$null
 [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

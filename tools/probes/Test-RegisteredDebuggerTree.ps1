param([Parameter(Mandatory=$true)][string]$AssemblyPath,[Parameter(Mandatory=$true)][string]$OutputDirectory,[switch]$UseBridge,[switch]$AllowTemporaryVbaAccess)
$ErrorActionPreference='Stop'
if (-not $UseBridge -or @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Registered bridge and an isolated Excel session are required.' }
$assembly=[Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$outputRoot=[IO.Path]::GetFullPath($OutputDirectory);[IO.Directory]::CreateDirectory($outputRoot) | Out-Null
$securityPath='HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$initial=Get-ItemProperty -LiteralPath $securityPath
$hadAccess=$null -ne $initial.PSObject.Properties['AccessVBOM'];$initialAccess=$initial.AccessVBOM
@{HadAccessVBOM=$hadAccess;AccessVBOM=$initialAccess} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $outputRoot 'security-before.json') -Encoding UTF8
if($initialAccess -ne 1 -and -not $AllowTemporaryVbaAccess){throw 'Explicit temporary AccessVBOM opt-in is required.'}
$excel=$null;$book=$null;$vbe=$null;$probeProcess=$null;$toolbarName=$null;$option=$null;$optionChanged=$false;$report=[ordered]@{}
function Invoke-Bridge([hashtable]$fields){
 $response=& (Join-Path $PSScriptRoot '../Invoke-VBAi.ps1') -HostProcessId $script:probeProcess.Id -RequestJson ($fields|ConvertTo-Json -Compress -Depth 12) -ResponseTimeoutSeconds 30 | ConvertFrom-Json
 if(-not $response.Ok){throw "$($fields.Command): $($response.Error)"};return $response.Data
}
function Open-ProbeExcel{
 if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Excel isolation was lost before launch.'}
 $scratch=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'VBAi-scratch.xlsx'))
 $script:probeProcess=Start-Process -FilePath 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE' -ArgumentList @('/x',('"'+$scratch+'"')) -WindowStyle Hidden -PassThru
 $script:excel=$null
 for($i=0;$i -lt 150 -and $null -eq $script:excel;$i++){Start-Sleep -Milliseconds 200;try{$script:excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')}catch{}}
 if($null -eq $script:excel -or @(Get-Process EXCEL).Count -ne 1 -or (Get-Process EXCEL).Id -ne $script:probeProcess.Id){throw 'The owned Excel process is not uniquely attached.'}
 if($script:excel.Workbooks.Count -ne 1 -or $script:excel.Workbooks.Item(1).FullName -ne $scratch){throw 'Unexpected workbook in the isolated process.'}
 $script:excel.Visible=$true;$script:vbe=$script:excel.GetType().InvokeMember('VBE',[Reflection.BindingFlags]::GetProperty,$null,$script:excel,$null);$script:vbe.MainWindow.Visible=$true
 $status=Invoke-Bridge @{Command='status'}
 if($status.HostProcessId -ne $script:probeProcess.Id -or $status.AssemblyModuleVersionId -ne $assembly.ManifestModule.ModuleVersionId.ToString('D')){throw 'Bridge process/assembly mismatch.'}
 $script:excel.Workbooks.Item(1).Close($false);return $status
}
function Close-ProbeExcel{
 if($null -ne $script:book){$script:book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:book);$script:book=$null}
 if($null -ne $script:excel){$script:excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:excel);$script:excel=$null;$script:vbe=$null}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($null -ne $script:probeProcess -and -not $script:probeProcess.WaitForExit(5000)){throw 'The owned Excel process has not exited; no next launch is allowed.'}
}
try{
 if($initialAccess -ne 1){New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value 1 -PropertyType DWord -Force|Out-Null}
 $report.Identity=Open-ProbeExcel
 $book=$excel.Workbooks.Add();$document=Join-Path $outputRoot ('Debugger-'+[Guid]::NewGuid().ToString('N')+'.xlsm');$book.SaveAs($document,52)
 $module=$book.VBProject.VBComponents.Add(1);$module.Name='LargeTreeProbe'
 $module.CodeModule.AddFromString(@"
Option Explicit
Public Sub InspectLargeTree()
    Dim values(1 To 1000) As Long
    Dim number As Double, money As Currency, stamp As Date
    Dim text As String, flag As Boolean, absent As Variant, index As Long
    For index = 1 To 1000
        values(index) = index * 3
    Next index
    number = 12.5: money = 7.25: stamp = DateSerial(2026, 9, 28)
    text = "VBAi debugger": flag = True: absent = Null
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "Paused"
    Stop
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = values(1000)
End Sub
"@)
 Invoke-Bridge @{Command='open_debug_pane';Action='locals'}|Out-Null
 $source=Invoke-Bridge @{Command='read_module';Project=$document;Module=$module.Name}
 $report.Start=Invoke-Bridge @{Command='run_sub';Project=$document;Module=$module.Name;Procedure='InspectLargeTree';ExpectedMode=2;ExpectedSha256=$source.Sha256}
 for($i=0;$i -lt 30;$i++){Start-Sleep -Milliseconds 200;$mode=Invoke-Bridge @{Command='debug_state';Project=$document};if($mode.Mode -eq 1){break}}
 if($mode.Mode -ne 1 -or $book.Worksheets.Item(1).Range('A1').Value2 -ne 'Paused'){throw 'Owned fixture did not pause at Stop.'}
 $before=Invoke-Bridge @{Command='debug_windows'};$report.Before=$before
 $row=@($before.Locals.Items|Where-Object Expression -eq 'values')
 if($row.Count -ne 1 -or @($row[0].PathSegments).Count -eq 0){throw 'Large-array root was not uniquely exposed.'}
 $report.Expand=Invoke-Bridge @{Command='debug_item';Pane='locals';Action='expand';PathSegments=@($row[0].PathSegments)}
 $expanded=Invoke-Bridge @{Command='debug_windows'};$report.Expanded=$expanded
 $children=@($expanded.Locals.Items|Where-Object { $_.PathSegments.Count -gt $row[0].PathSegments.Count -and $_.PathSegments[0] -eq $row[0].PathSegments[0] })
 if($children.Count -eq 0){throw 'No array child exposed after expansion.'}
 $verifiedValues=0
 foreach($child in $children){
  if($child.Expression -notmatch '^values\(([0-9]+)\)$' -or [int]$Matches[1] -lt 1 -or [int]$Matches[1] -gt 1000 -or $child.Value -ne [string]([int]$Matches[1]*3)){throw 'Array child value differs from its deterministic fixture.'}
  $verifiedValues++
 }
 if(@($children|Group-Object Expression|Where-Object Count -ne 1).Count){throw 'Duplicate array observations remain.'}
 $report.Array=@{ExpectedTotal=1000;ExposedChildren=$children.Count;VerifiedValues=$verifiedValues;Coverage=$expanded.Locals.Coverage;Exhaustive=($children.Count -eq 1000)}
 $report.Collapse=Invoke-Bridge @{Command='debug_item';Pane='locals';Action='collapse';PathSegments=@($row[0].PathSegments)}
 $collapsed=Invoke-Bridge @{Command='debug_windows'};$report.Collapsed=$collapsed
 if(@($collapsed.Locals.Items|Where-Object { $_.PathSegments.Count -gt $row[0].PathSegments.Count -and $_.PathSegments[0] -eq $row[0].PathSegments[0] }).Count){throw 'Array children remain exposed after collapse.'}
 foreach($name in @('number','money','stamp','text','flag','absent','index')){
  if(@($before.Locals.Items|Where-Object { $_.Expression -eq $name -and $_.Type -and $null -ne $_.Value }).Count -ne 1){throw ('Scalar row missing: '+$name)}
 }
 $resume=$vbe.CommandBars.FindControl(1,186)
 $report.Resume=Invoke-Bridge @{Command='invoke_debug';Project=$document;Module=$module.Name;ExpectedMode=1;ExpectedSha256=$source.Sha256;StartLine=13;Action='continue';ControlId=186;ControlCaption=$resume.Caption}
 Start-Sleep -Milliseconds 300
 $after=Invoke-Bridge @{Command='debug_state';Project=$document}
 if($after.Mode -ne 2 -or $book.Worksheets.Item(1).Range('A1').Value2 -ne 3000){throw 'Fixture did not resume exactly once.'}
 $read=Invoke-Bridge @{Command='read_module';Project=$document;Module=$module.Name}
 if($read.Sha256 -ne $source.Sha256){throw 'Debugger actions altered the source.'}
 $report.RuntimeMarker=3000;$report.CodeUnchanged=$true
}catch{$report|ConvertTo-Json -Depth 30|Set-Content -LiteralPath (Join-Path $outputRoot 'partial-results.json') -Encoding UTF8;throw}
finally{
 try{
  if($null -ne $book){$state=Invoke-Bridge @{Command='debug_state';Project=$document};if($state.Mode -eq 1){Invoke-Bridge @{Command='debug_global';Project=$document;ExpectedMode=1;Action='reset'}|Out-Null}}
  Close-ProbeExcel
 }finally{
  if($hadAccess){New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value $initialAccess -PropertyType DWord -Force|Out-Null}else{Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue}
 }
}
$report|ConvertTo-Json -Depth 30|Set-Content -LiteralPath (Join-Path $outputRoot 'registered-debugger-tree.json') -Encoding UTF8
Write-Output 'PASS native large array expand/collapse, scalar types and resume; exposed-row count recorded.'

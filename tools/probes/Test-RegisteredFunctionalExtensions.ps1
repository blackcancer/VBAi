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
$excel=$null;$book=$null;$vbe=$null;$probeProcess=$null;$toolbarName=$null;$option=$null;$optionChanged=$false;$report=[ordered]@{};$cleanupErrors=@()
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
 if($initialAccess -ne 1){New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null}
 $report.Identity=Open-ProbeExcel
 $book=$excel.Workbooks.Add();$document=Join-Path $outputRoot ('Arguments-'+[Guid]::NewGuid().ToString('N')+'.xlsm');$book.SaveAs($document,52)
 $module=$book.VBProject.VBComponents.Add(1);$module.Name='ArgumentsProbe'
 $module.CodeModule.AddFromString(@"
Option Explicit
Public Sub AcceptArguments(ByVal text As String, ByVal number As Double, ByVal flag As Boolean, ByVal absent As Variant)
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = text
    ThisWorkbook.Worksheets(1).Range("B1").Value2 = number
    ThisWorkbook.Worksheets(1).Range("C1").Value2 = flag
    ThisWorkbook.Worksheets(1).Range("D1").Value2 = IsNull(absent)
End Sub
Public Function ReturnResult(ByVal number As Double) As Double
    ReturnResult = number * 2
    ThisWorkbook.Worksheets(1).Range("E1").Value2 = ReturnResult
End Function
"@)
 $source=Invoke-Bridge @{Command='read_module';Project=$document;Module='ArgumentsProbe'};$runs=@()
 foreach($procedure in @('AcceptArguments','ReturnResult')){
  $arguments=if($procedure -eq 'AcceptArguments'){@('VBAi "quoted" français',12.5,$true,$null)}else{@(21)}
  $queued=Invoke-Bridge @{Command='run_procedure';Project=$document;Module='ArgumentsProbe';Procedure=$procedure;ExpectedMode=2;ExpectedSha256=$source.Sha256;Arguments=@($arguments)}
  for($i=0;$i -lt 30;$i++){Start-Sleep -Milliseconds 200;$state=Invoke-Bridge @{Command='procedure_run_status';Project=$document;Query=$queued.Query};if(-not $state.Pending){break}}
  if($state.State -ne 'Delivered'){throw ('Procedure was not delivered: '+($state|ConvertTo-Json -Depth 8 -Compress))};$runs+=$state
 }
 if($book.Worksheets.Item(1).Range('A1').Value2 -cne 'VBAi "quoted" français' -or $book.Worksheets.Item(1).Range('B1').Value2 -ne 12.5 -or -not $book.Worksheets.Item(1).Range('C1').Value2 -or -not $book.Worksheets.Item(1).Range('D1').Value2 -or $book.Worksheets.Item(1).Range('E1').Value2 -ne 42){throw 'Independent worksheet readback differs from the fixture.'}
 $report.Procedures=@{Runs=$runs;RuntimeReadbackVerified=$true;FunctionResult=42}

 $matrix=@();$catalog=Invoke-Bridge @{Command='list_form_control_types'}
 foreach($type in @('CheckBox','ComboBox','CommandButton','Frame','Image','Label','ListBox','MultiPage','OptionButton','ScrollBar','SpinButton','TabStrip','TextBox','ToggleButton')){
  $form=$book.VBProject.VBComponents.Add(3);$form.Name='Matrix'+$type
  $control=$form.Designer.Controls.Add(('Forms.'+$type+'.1'),'ProbeControl',$true)
  $properties=Invoke-Bridge @{Command='form_control_properties';Project=$document;Form=$form.Name;Control='ProbeControl'}
  foreach($property in @('Left','Enabled','Caption')){
   $descriptor=@($properties|Where-Object Name -eq $property)
   if($descriptor.Count -ne 1 -or $descriptor[0].ReadOnly -or $descriptor[0].Error){$matrix+=@{Type=$type;Property=$property;State='NOT_APPLICABLE';Reason='Property absent, read-only or unreadable'};continue}
   $original=$control.$property
   if($null -eq $original -and $property -eq 'Caption'){$control.Caption='Baseline caption';$original=$control.Caption}
   if($null -eq $original){$matrix+=@{Type=$type;Property=$property;State='NOT_APPLICABLE';Reason='Null original value cannot be restored by the scalar setter'};continue}
   $value=if($property -eq 'Left'){[double]$original+9.5}elseif($property -eq 'Enabled'){-not [bool]$original}else{'VBAi property fixture'}
   $tree=Invoke-Bridge @{Command='form_tree';Project=$document;Form=$form.Name};$node=@($tree.Controls|Where-Object Name -eq 'ProbeControl')[0]
   @{Type=$type;Property=$property;Node=$node;TreeVersion=$tree.TreeVersion;Value=$value;Original=$original}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $outputRoot 'current-property-trial.json') -Encoding UTF8
   Invoke-Bridge @{Command='set_form_node_property';Project=$document;Form=$form.Name;ControlPath=$node.Path;Property=$property;Value=$value;ExpectedTreeVersion=$tree.TreeVersion}|Out-Null
   if($control.$property -ne $value){throw "Independent property readback failed: $type/$property"}
   $tree=Invoke-Bridge @{Command='form_tree';Project=$document;Form=$form.Name}
   Invoke-Bridge @{Command='set_form_node_property';Project=$document;Form=$form.Name;ControlPath=$node.Path;Property=$property;Value=$original;ExpectedTreeVersion=$tree.TreeVersion}|Out-Null
   if($control.$property -ne $original){throw "Property restoration failed: $type/$property"}
   $matrix+=@{Type=$type;Property=$property;State='PASS';Before=$original;Changed=$value;Restored=$true}
  }
 }

 $activex=@()
 foreach($candidate in @($catalog|Where-Object { $_.Source -eq 'COM CATID_Control x64' -and $_.ProgId -like 'MSComctlLib.*' })){
  $form=$book.VBProject.VBComponents.Add(3);$form.Name='ActiveXProbe'+$activex.Count
  $state=Invoke-Bridge @{Command='form_state';Project=$document;Form=$form.Name}
  try{
   $addedControl=Invoke-Bridge @{Command='add_form_control';Project=$document;Form=$form.Name;ControlType=$candidate.ProgId;Control='HostedControl';Left=10;Top=10;Width=150;Height=40;ExpectedFormVersion=$state.Version}
   if($form.Designer.Controls.Count -ne 1){throw 'Native control count differs from the expected hosting fixture.'}
   $activex+=@{ProgId=$candidate.ProgId;State='HOSTED';NativeCount=$form.Designer.Controls.Count}
  }catch{
   $activex+=@{ProgId=$candidate.ProgId;State='HOSTING_REFUSED';Error=$_.Exception.Message;ControlsRemaining=$form.Designer.Controls.Count}
  }
 }
 $report.ActiveXHosting=$activex
 $activex|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $outputRoot 'activex-hosting.json') -Encoding UTF8
 $report.ControlProperties=$matrix;$report.InstalledActiveX=$catalog
 $matrix|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $outputRoot 'control-property-matrix.json') -Encoding UTF8
 $before=Invoke-Bridge @{Command='read_vbe_options'};$before | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $outputRoot 'options-before.json') -Encoding UTF8
 foreach($tab in $before.Tabs){foreach($control in $tab.Controls){
  if($tab.Tab -match '^(Editor|Éditeur|Editeur)$' -and $control.Type -eq 'ControlType.CheckBox' -and $control.Name.Replace('&','').Trim() -match '^(Auto Syntax Check|Vérification automatique de la syntaxe)$' -and $control.Value -in @('On','Off')){$option=@{Pane=$tab.Tab;Property=$control.Name;Original=($control.Value -eq 'On')}}
 }}
 if($null -eq $option){throw 'The exact supported syntax-check option could not be identified.'}
 $optionChanged=$true
 $changed=Invoke-Bridge @{Command='set_vbe_option';Pane=$option.Pane;Property=$option.Property;Value=(-not $option.Original);ExpectedOptionsVersion=$before.OptionsVersion}
 $after=Invoke-Bridge @{Command='read_vbe_options'}
 $actual=@($after.Tabs|Where-Object Tab -eq $option.Pane|ForEach-Object Controls|Where-Object Name -eq $option.Property);$expected=if($option.Original){'Off'}else{'On'}
 if($actual.Count -ne 1 -or $actual[0].Value -ne $expected){throw 'Preference did not persist after reopening.'}
 Invoke-Bridge @{Command='set_vbe_option';Pane=$option.Pane;Property=$option.Property;Value=$option.Original;ExpectedOptionsVersion=$after.OptionsVersion}|Out-Null
 $restoredRead=Invoke-Bridge @{Command='read_vbe_options'};if($restoredRead.OptionsVersion -ne $before.OptionsVersion){throw 'Full options state was not restored.'};$optionChanged=$false
 $report.Options=@{Changed=$changed;ReopenedValue=$actual[0].Value;OriginalStateRestored=$true}
 $optionMatrix=@()
 $baseline=Invoke-Bridge @{Command='read_vbe_options'}
 foreach($tab in $baseline.Tabs){
  if($tab.Tab -notin @('Éditeur','Editor','Général','General')){continue}
  $originalRadio=@($tab.Controls|Where-Object { $_.Type -eq 'ControlType.RadioButton' -and $_.Value -eq $true })
  foreach($control in $tab.Controls){
   if($control.Type -eq 'ControlType.CheckBox' -and $control.Name -match '^(Vérification automatique de la syntaxe|Déclaration des variables obligatoire|Complément automatique des instructions|Info express automatique|Info-bulles automatiques|Retrait automatique|Compilation sur demande|Compilation en arrière-plan|Auto Syntax Check|Require Variable Declaration|Auto List Members|Auto Quick Info|Auto Data Tips|Auto Indent|Compile on Demand|Background Compile)$'){
    $values=@(-not ($control.Value -eq 'On'));$restoreProperty=$control.Name;$restoreValue=($control.Value -eq 'On')
   }elseif($control.Type -eq 'ControlType.Edit' -and $control.Name -match '^(Largeur de la tabulation|Tab Width)'){
    $values=@(1,32);$restoreProperty=$control.Name;$restoreValue=[int]$control.Value
   }elseif($control.Type -eq 'ControlType.RadioButton'){
    if($originalRadio.Count -ne 1){throw 'Original error-trapping radio is ambiguous.'}
    $values=@($true);$restoreProperty=$originalRadio[0].Name;$restoreValue=$true
   }else{continue}
   foreach($value in $values){
    $before=Invoke-Bridge @{Command='read_vbe_options'}
    $option=@{Pane=$tab.Tab;Property=$restoreProperty;Original=$restoreValue};$optionChanged=$true
    Invoke-Bridge @{Command='set_vbe_option';Pane=$tab.Tab;Property=$control.Name;Value=$value;ExpectedOptionsVersion=$before.OptionsVersion}|Out-Null
    $after=Invoke-Bridge @{Command='read_vbe_options'}
    $actual=@($after.Tabs|Where-Object Tab -eq $tab.Tab|ForEach-Object Controls|Where-Object { $_.Name -eq $control.Name -and $_.Type -eq $control.Type })
    $expected=if($control.Type -eq 'ControlType.CheckBox'){if($value){'On'}else{'Off'}}elseif($control.Type -eq 'ControlType.Edit'){[string]$value}else{$true}
    if($actual.Count -ne 1 -or $actual[0].Value -ne $expected){throw ('Reopened preference differs: '+$control.Name)}
    Invoke-Bridge @{Command='set_vbe_option';Pane=$option.Pane;Property=$option.Property;Value=$option.Original;ExpectedOptionsVersion=$after.OptionsVersion}|Out-Null
    $restoredRead=Invoke-Bridge @{Command='read_vbe_options'}
    if($restoredRead.OptionsVersion -ne $before.OptionsVersion){throw ('Full preference state not restored: '+$control.Name)}
    $optionChanged=$false;$optionMatrix+=@{Tab=$tab.Tab;Property=$control.Name;Value=$value;State='PASS';Reopened=$true;FullStateRestored=$true}
   }
  }
 }
 $report.OptionMatrix=$optionMatrix
 $optionMatrix|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $outputRoot 'option-matrix.json') -Encoding UTF8
 $bars=Invoke-Bridge @{Command='list_toolbars'};$created=Invoke-Bridge @{Command='create_toolbar';ObjectName=('Persistence-'+[Guid]::NewGuid().ToString('N'));Temporary=$false;ExpectedToolbarCollectionVersion=$bars.ToolbarCollectionVersion};$toolbarName=$created.ObjectName
 $commands=Invoke-Bridge @{Command='toolbar_controls';ObjectName=$toolbarName};$native=$vbe.CommandBars.FindControl(1,186)
 $added=Invoke-Bridge @{Command='add_toolbar_command';ObjectName=$toolbarName;ControlId=186;ControlCaption=$native.Caption;Temporary=$false;ExpectedToolbarControlsVersion=$commands.ToolbarControlsVersion};if(-not $added.Verified){throw ('Persistent button creation was not verified: '+($added|ConvertTo-Json -Depth 12 -Compress))}
 $bars=Invoke-Bridge @{Command='list_toolbars'};$layout=@($bars.Toolbars|Where-Object { $_.Properties.Name -eq $toolbarName })
 if($layout.Count -ne 1){throw 'Owned toolbar layout is ambiguous.'}
 $floating=Invoke-Bridge @{Command='set_toolbar_position';ObjectName=$toolbarName;Action='float';ExpectedToolbarLayoutVersion=$layout[0].ToolbarLayoutVersion}
 if(-not $floating.Verified){throw 'Toolbar floating mode was not verified.'}
 $bars=Invoke-Bridge @{Command='list_toolbars'};$layout=@($bars.Toolbars|Where-Object { $_.Properties.Name -eq $toolbarName })
 $placed=Invoke-Bridge @{Command='set_toolbar_placement';ObjectName=$toolbarName;Action='float';ToolbarLeft=320;ToolbarTop=220;ExpectedToolbarLayoutVersion=$layout[0].ToolbarLayoutVersion}
 if(-not $placed.Verified){throw 'Toolbar floating coordinates were not verified.'}
 $report.ToolbarPlacement=$placed
 $report.ToolbarBefore=$added;Close-ProbeExcel;$report.RestartIdentity=Open-ProbeExcel;$report.RestartSourceCaption=$vbe.CommandBars.FindControl(1,186).Caption
 $persisted=Invoke-Bridge @{Command='toolbar_controls';ObjectName=$toolbarName};$report.RestartToolbar=$persisted;$report.RestartToolbarList=Invoke-Bridge @{Command='list_toolbars'};if(@($persisted.Controls).Count -ne 1 -or $persisted.Controls[0].Tag -ne $added.Tag -or $persisted.Controls[0].Id -ne 186){throw 'Toolbar/button ownership did not persist across native restart.'}
 $restartLayout=@($report.RestartToolbarList.Toolbars|Where-Object { $_.Properties.Name -eq $toolbarName })
 if($restartLayout.Count -ne 1 -or $restartLayout[0].Geometry.Position -ne 4 -or $restartLayout[0].Geometry.Left -ne 320 -or $restartLayout[0].Geometry.Top -ne 220){throw 'Toolbar floating geometry did not persist across restart.'}
 $report.Toolbar=@{Created=$created;Added=$added;AfterRestart=$persisted;PersistenceVerified=$true;FloatingPlacementVerified=$true}
}catch{$report|ConvertTo-Json -Depth 20|Set-Content -LiteralPath (Join-Path $outputRoot 'partial-results.json') -Encoding UTF8;$_|Out-String|Set-Content -LiteralPath (Join-Path $outputRoot 'failure.txt') -Encoding UTF8;throw}
finally{
 if($optionChanged -and $null -ne $excel){try{$state=Invoke-Bridge @{Command='read_vbe_options'};Invoke-Bridge @{Command='set_vbe_option';Pane=$option.Pane;Property=$option.Property;Value=$option.Original;ExpectedOptionsVersion=$state.OptionsVersion}|Out-Null;$restoredRead=Invoke-Bridge @{Command='read_vbe_options'};if($restoredRead.OptionsVersion -ne $before.OptionsVersion){throw 'Original preferences were not restored.'}}catch{$cleanupErrors+="Preference restoration failed: $_";Write-Warning $cleanupErrors[-1]}}
 if($null -ne $toolbarName -and $null -ne $vbe){try{
 $state=Invoke-Bridge @{Command='toolbar_controls';ObjectName=$toolbarName}
 foreach($button in @($state.Controls)){
  Invoke-Bridge @{Command='remove_toolbar_command';ObjectName=$toolbarName;ControlId=$button.Id;ControlCaption=$button.Caption;InsertIndex=$button.Index;ExpectedToolbarControlsVersion=$state.ToolbarControlsVersion}|Out-Null
  $state=Invoke-Bridge @{Command='toolbar_controls';ObjectName=$toolbarName}
 }
 $bars=Invoke-Bridge @{Command='list_toolbars'}
 Invoke-Bridge @{Command='remove_toolbar';ObjectName=$toolbarName;ExpectedToolbarControlsVersion=$state.ToolbarControlsVersion;ExpectedToolbarCollectionVersion=$bars.ToolbarCollectionVersion}|Out-Null
}catch{$cleanupErrors+="Toolbar cleanup failed: $_";Write-Warning $cleanupErrors[-1]}}
 try{Close-ProbeExcel}catch{$cleanupErrors+="Excel cleanup failed: $_";Write-Warning $cleanupErrors[-1]}
 if($hadAccess){New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value $initialAccess -PropertyType DWord -Force|Out-Null}else{Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue}
 $restoredSecurity=Get-ItemProperty -LiteralPath $securityPath
 if(($null -ne $restoredSecurity.PSObject.Properties['AccessVBOM']) -ne $hadAccess -or ($hadAccess -and $restoredSecurity.AccessVBOM -ne $initialAccess)){$cleanupErrors+='AccessVBOM restoration differs from its original state.'}
 $report.Cleanup=@{Errors=@($cleanupErrors);Verified=($cleanupErrors.Count -eq 0)}
 $report.Cleanup|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $outputRoot 'cleanup.json') -Encoding UTF8
 if($cleanupErrors.Count){throw ($cleanupErrors -join '; ')}
}
$report|ConvertTo-Json -Depth 20|Set-Content -LiteralPath (Join-Path $outputRoot 'registered-functional-extensions.json') -Encoding UTF8
Write-Output 'PASS registered procedure arguments, native option reopening and toolbar persistence.'

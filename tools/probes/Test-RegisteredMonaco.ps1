param([Parameter(Mandatory=$true)][string]$AssemblyPath,[Parameter(Mandatory=$true)][string]$OutputDirectory,[switch]$UseBridge,[switch]$AllowTemporaryVbaAccess)
$ErrorActionPreference='Stop'
if (-not $UseBridge -or @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'An isolated Excel session is required.' }
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes,System.Drawing,Accessibility
Add-Type -ReferencedAssemblies Accessibility @'
using System;
using System.Runtime.InteropServices;
public static class MonacoProbeMouse {
 [StructLayout(LayoutKind.Sequential)] public struct Point { public int X,Y; }
 [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left,Top,Right,Bottom; }
 [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr window, out Rect rect);
 [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr window, out Rect rect);
 [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr window);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr window, System.Text.StringBuilder name, int count);
 [DllImport("user32.dll")] public static extern bool ScreenToClient(IntPtr window, ref Point point);
 [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);
 [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);
 [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object accessible);
 public static Point FindModule(IntPtr tree, string name) {
  var nodes = new System.Collections.Generic.Queue<IntPtr>();
  nodes.Enqueue(SendMessage(tree, 0x110A, IntPtr.Zero, IntPtr.Zero));
  for(int n=0; nodes.Count>0 && n<300; n++) {
   var node=nodes.Dequeue(); if(node==IntPtr.Zero) continue;
   SendMessage(tree,0x1102,new IntPtr(2),node);
   nodes.Enqueue(SendMessage(tree,0x110A,new IntPtr(4),node));
   nodes.Enqueue(SendMessage(tree,0x110A,new IntPtr(1),node));
  }
  Guid iid=new Guid("618736E0-3C3D-11CF-810C-00AA00389B71"); object value;
  Marshal.ThrowExceptionForHR(AccessibleObjectFromWindow(tree,0xFFFFFFFC,ref iid,out value));
  try {
   var root=(Accessibility.IAccessible)value;
   for(int i=1;i<=root.accChildCount;i++) {
    if(root.get_accName(i)!=name) continue;
    root.accSelect(3,i); int x,y,w,h; root.accLocation(out x,out y,out w,out h,i);
    if(w<=0 || h<=0) throw new InvalidOperationException("Module bounds unavailable.");
    return new Point {X=x+w/2,Y=y+h/2};
   }
   throw new InvalidOperationException("Module not found in native accessibility tree.");
  } finally {Marshal.ReleaseComObject(value);}
 }
}
'@
$assembly=[Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$outputRoot=[IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($outputRoot) | Out-Null
$excel=$null;$book=$null;$probeProcess=$null
$securityPath='HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$security=Get-ItemProperty -LiteralPath $securityPath
$hadAccess=$null -ne $security.PSObject.Properties['AccessVBOM'];$originalAccess=$security.AccessVBOM
@{HadAccessVBOM=$hadAccess;AccessVBOM=$originalAccess}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $outputRoot 'security-before.json') -Encoding UTF8
if($originalAccess -ne 1 -and -not $AllowTemporaryVbaAccess){throw 'Explicit temporary VBA access opt-in required.'}
function Wait-Condition([scriptblock]$condition) {
 $watch=[Diagnostics.Stopwatch]::StartNew()
 do { $value=& $condition; if($value){return $value}; Start-Sleep -Milliseconds 200 } while($watch.Elapsed.TotalSeconds -lt 40)
 throw 'Timed out waiting for the registered Monaco editor.'
}
function Find-Control([string]$id) {
 $conditions=[Windows.Automation.AndCondition]::new(
  [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,$script:probeProcess.Id),
  [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::AutomationIdProperty,$id))
 [Windows.Automation.AutomationElement]::RootElement.FindFirst([Windows.Automation.TreeScope]::Descendants,$conditions)
}
function Close-ModuleTab {
 $tabs=Wait-Condition {Find-Control 'tabs'}
 $tab=Wait-Condition {@($tabs.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::TabItem)) | Where-Object {$_.Current.Name -like '*MonacoDockProbe*'})[0]}
 $rect=$tab.Current.BoundingRectangle
 $point=[MonacoProbeMouse+Point]::new(); $point.X=[int]($rect.Right-13); $point.Y=[int]($rect.Top+$rect.Height/2)
 $handle=[IntPtr]$tabs.Current.NativeWindowHandle
 [void][MonacoProbeMouse]::ScreenToClient($handle,[ref]$point)
 $position=[IntPtr](($point.Y -shl 16) -bor ($point.X -band 65535))
 [void][MonacoProbeMouse]::PostMessage($handle,0x201,[IntPtr]1,$position)
 [void][MonacoProbeMouse]::PostMessage($handle,0x202,[IntPtr]0,$position)
}
try {
 if($originalAccess -ne 1){New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null}
 $scratch=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'CodexVBE-scratch.xlsx'))
 $probeProcess=Start-Process -FilePath 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE' -ArgumentList @('/x',('"'+$scratch+'"')) -WindowStyle Hidden -PassThru
 $excel=Wait-Condition {try{[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')}catch{}}
 if(@(Get-Process EXCEL).Count -ne 1 -or (Get-Process EXCEL).Id -ne $probeProcess.Id){throw 'Unexpected Excel ownership.'}
 if($excel.Workbooks.Count -ne 1 -or $excel.Workbooks.Item(1).FullName -ne $scratch){throw 'Unexpected workbook.'}
 $excel.Visible=$true;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false
 $vbe=$excel.GetType().InvokeMember('VBE',[Reflection.BindingFlags]::GetProperty,$null,$excel,$null)
 $vbe.MainWindow.Visible=$true
 $response=& (Join-Path $PSScriptRoot '../Invoke-CodexVBE.ps1') -HostProcessId $probeProcess.Id -RequestJson '{"Command":"status"}' -ResponseTimeoutSeconds 30 | ConvertFrom-Json
 if(-not $response.Ok -or $response.Data.AssemblyModuleVersionId -ne $assembly.ManifestModule.ModuleVersionId.ToString('D')){throw 'Loaded assembly mismatch.'}
 $book=$excel.Workbooks.Add();$module=$book.VBProject.VBComponents.Add(1);$module.Name='MonacoDockProbe'
 $module.CodeModule.AddFromString("Option Explicit`r`nPublic Sub Probe()`r`n    Debug.Print 42`r`nEnd Sub")
 $module.CodeModule.CodePane.Show()
 $button=$vbe.CommandBars.FindControl(1,[Type]::Missing,'CodexVBE.ModernEditor',$false)
 if($null -eq $button){throw 'Monaco View menu is absent.'}
 $button.Execute()
 Write-Output 'Menu executed; locating editor shell.'
 $editor=Wait-Condition {Find-Control 'ModernEditorWindow'}
 $editorCaption=$editor.Current.Name; $before=$editor.Current.BoundingRectangle
 if($before.Width -lt 700 -or $before.Height -lt 400){throw 'Unexpected editor bounds.'}
 # UI Automation waits for the real browser document, not only the WinForms shell.
 $browser=Wait-Condition {$editor.FindFirst([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::Document))}
 Write-Output 'Closing module tab.'
 Close-ModuleTab
 @($vbe.Windows | ForEach-Object {@{Caption=$_.Caption;Type=$_.Type;Visible=$_.Visible}}) | ConvertTo-Json | Set-Content (Join-Path $outputRoot 'windows-after-close.json')
 [void](Wait-Condition {if(@($vbe.Windows | Where-Object {$_.Type -eq 0 -and $_.Caption -like '*MonacoDockProbe*'}).Count -eq 0){return $true}})
 Write-Output 'Native window closed; locating project tree.'
 $project= @($vbe.Windows | Where-Object {$_.Type -eq 6})[0];$project.Visible=$true;$project.SetFocus()
 $findTrees=$assembly.GetType('CodexVBE.EditorProjectNavigation').GetMethod('FindProjectTrees',[Reflection.BindingFlags]'Static,NonPublic')
 $treeHandle=Wait-Condition {@($findTrees.Invoke($null,@([int]$probeProcess.Id,[string]$project.Caption)))[0]}
 $tree=[Windows.Automation.AutomationElement]::FromHandle([IntPtr]$treeHandle)
 Write-Output 'Project tree found; locating module node.'
 $treeHandle=[IntPtr]$treeHandle
 $point=[MonacoProbeMouse]::FindModule($treeHandle,'MonacoDockProbe')
 Write-Output 'Module located by MSAA; sending double-click.'
 [void][MonacoProbeMouse]::ScreenToClient($treeHandle,[ref]$point)
 $position=[IntPtr](($point.Y -shl 16) -bor ($point.X -band 65535))
 # Direct mouse messages to the verified project tree; never global input or VBE shortcuts.
 [void][MonacoProbeMouse]::PostMessage($treeHandle,0x201,[IntPtr]1,$position)
 [void][MonacoProbeMouse]::PostMessage($treeHandle,0x202,[IntPtr]0,$position)
 Write-Output 'Double-click delivered; checking Monaco tab.'
 [void][MonacoProbeMouse]::PostMessage($treeHandle,0x203,[IntPtr]1,$position)
 [void][MonacoProbeMouse]::PostMessage($treeHandle,0x202,[IntPtr]0,$position)
 [void](Wait-Condition {@($editor.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::TabItem)) | Where-Object {$_.Current.Name -like '*MonacoDockProbe*'})[0]})
 if(@($vbe.Windows | Where-Object {$_.Type -eq 0 -and $_.Caption -like '*MonacoDockProbe*'}).Count -ne 1){throw 'Native backing window must remain alive behind Monaco.'}
 Write-Output 'Project double-click opened Monaco and retained its backing code window.'
 $parent=[MonacoProbeMouse]::GetParent([IntPtr]$editor.Current.NativeWindowHandle)
 $parentClass=[Text.StringBuilder]::new(128)
 [void][MonacoProbeMouse]::GetClassName($parent,$parentClass,$parentClass.Capacity)
 if($parentClass.ToString() -ne 'MDIClient'){throw 'Monaco must live in the native document workspace.'}
 $client=[MonacoProbeMouse+Rect]::new()
 [void][MonacoProbeMouse]::GetClientRect($parent,[ref]$client)
 $editorRect=$editor.Current.BoundingRectangle
 if([Math]::Abs($editorRect.Width-($client.Right-$client.Left)) -gt 2 -or [Math]::Abs($editorRect.Height-($client.Bottom-$client.Top)) -gt 2){throw 'Monaco does not fill the native document workspace.'}
 $originalWidth=$vbe.MainWindow.Width; $originalHeight=$vbe.MainWindow.Height
 try {
  $vbe.MainWindow.Width=[Math]::Max(950,$originalWidth-120)
  $vbe.MainWindow.Height=[Math]::Max(650,$originalHeight-80)
  [void](Wait-Condition {
   $client=[MonacoProbeMouse+Rect]::new(); [void][MonacoProbeMouse]::GetClientRect($parent,[ref]$client)
   $bounds=$editor.Current.BoundingRectangle
   [Math]::Abs($bounds.Width-($client.Right-$client.Left)) -le 2 -and [Math]::Abs($bounds.Height-($client.Bottom-$client.Top)) -le 2
  })
 } finally {$vbe.MainWindow.Width=$originalWidth; $vbe.MainWindow.Height=$originalHeight}
 $treeSnapshot=& (Join-Path $PSScriptRoot 'Inspect-VbeWindowTree.ps1') -HostProcessId $probeProcess.Id
 $treeSnapshot | Set-Content (Join-Path $outputRoot 'window-tree.json')
 $rootRect=[MonacoProbeMouse+Rect]::new(); [void][MonacoProbeMouse]::GetWindowRect([IntPtr]$vbe.MainWindow.HWnd,[ref]$rootRect)
 $rootBitmap=[Drawing.Bitmap]::new($rootRect.Right-$rootRect.Left,$rootRect.Bottom-$rootRect.Top)
 $rootGraphics=[Drawing.Graphics]::FromImage($rootBitmap)
 try {$rootGraphics.CopyFromScreen($rootRect.Left,$rootRect.Top,0,0,$rootBitmap.Size);$rootBitmap.Save((Join-Path $outputRoot 'vbe-workspace.png'))} finally {$rootGraphics.Dispose();$rootBitmap.Dispose()}

 $rect=$editor.Current.BoundingRectangle
 $bitmap=[Drawing.Bitmap]::new([int]$rect.Width,[int]$rect.Height)
 $graphics=[Drawing.Graphics]::FromImage($bitmap)
 try{$graphics.CopyFromScreen([int]$rect.X,[int]$rect.Y,0,0,$bitmap.Size);$bitmap.Save((Join-Path $outputRoot 'monaco-host.png'))}finally{$graphics.Dispose();$bitmap.Dispose()}
 Write-Output 'Closing module tab.'
 Close-ModuleTab
 @($vbe.Windows | ForEach-Object {@{Caption=$_.Caption;Type=$_.Type;Visible=$_.Visible}}) | ConvertTo-Json | Set-Content (Join-Path $outputRoot 'windows-after-close.json')
 [void](Wait-Condition {if(@($vbe.Windows | Where-Object {$_.Type -eq 0 -and $_.Caption -like '*MonacoDockProbe*'}).Count -eq 0){return $true}})
 @{State='PASS';HostProcessId=$probeProcess.Id;Assembly=$response.Data.AssemblyModuleVersionId;Menu=$true;BrowserDocument=$true;DocumentWorkspace=$true;FillsWorkspace=$true;NativeResize=$true;TabClose=$true;ProjectDoubleClick=$true;NativeWindowLifetime=$true;MacroExecuted=$false}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $outputRoot 'monaco-host.json') -Encoding UTF8
 # The workspace remains owned by the VBE until the host closes.
} catch {
 ($_ | Out-String) + $_.ScriptStackTrace | Set-Content -LiteralPath (Join-Path $outputRoot 'failure.txt') -Encoding UTF8
 Write-Output $_.ScriptStackTrace
 throw
} finally {
 if($originalAccess -ne 1){if($hadAccess){Set-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value $originalAccess}else{Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM}}
 $restored=Get-ItemProperty -LiteralPath $securityPath
 if(($null -ne $restored.PSObject.Properties['AccessVBOM']) -ne $hadAccess -or $restored.AccessVBOM -ne $originalAccess){throw 'VBA access restoration failed.'}
 if($null -ne $book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
 if($null -ne $excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 if($null -ne $probeProcess -and -not $probeProcess.WaitForExit(10000)){throw 'The owned Excel process did not exit.'}
}

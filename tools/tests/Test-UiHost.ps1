<#
.SYNOPSIS
Explicit registered-control, disposable Excel or selected IDE checks.
ChatControlActivation exercises the registered COM class. ExcelFormLayouts
requires no pre-existing Excel process and owns only its disposable workbook.
VisualStudioWinForms requires the selected PID to have this checkout open.
.DESCRIPTION
Select exactly one scenario. No aggregate scenario is provided. Assembly and
output defaults remain specific to each scenario; pass an explicit candidate
assembly when required. Shared helper import opens no window or host.
.EXAMPLE
.\Test-UiHost.ps1 -Scenario ChatControlActivation
.EXAMPLE
.\Test-UiHost.ps1 -Scenario VisualStudioWinForms -VisualStudioProcessId 1234
#>
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('ChatControlActivation','ExcelFormLayouts','VisualStudioWinForms')][string]$Scenario,
    [string]$AssemblyPath, [string]$OutputDirectory,
    [switch]$AllowTemporaryVbaAccess, [int]$VisualStudioProcessId
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'UiProbe.psm1') -ErrorAction Stop

function Invoke-ChatControlActivation {
    if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit Windows PowerShell.' }

    $type = [Type]::GetTypeFromProgID('VBAi.ChatToolWindow', $true)
    $control = [Activator]::CreateInstance($type)
    try {
        if ($control.GetType().FullName -ne 'VBAi.ChatToolWindow') {
            throw 'The chat control ProgID activated an unexpected class.'
        }
        Write-Output 'PASS chat control COM activation in the current user session'
    }
    finally {
        if ($control -is [IDisposable]) { $control.Dispose() }
        elseif ([Runtime.InteropServices.Marshal]::IsComObject($control)) {
            [Runtime.InteropServices.Marshal]::ReleaseComObject($control) | Out-Null
        }
    }
}

function Invoke-ExcelFormLayouts {
    param(
        [Parameter(Mandatory = $true)][string]$AssemblyPath,
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [switch]$AllowTemporaryVbaAccess
    )
    if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close existing Excel instances before this isolated test.' }
    $assembly = Open-UiAssembly $AssemblyPath
    $directory = [IO.Path]::GetFullPath($OutputDirectory)
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    $document = Join-Path $directory ('Vbai-EditorProbe-' + [Guid]::NewGuid().ToString('N') + '.xlsm')
    $securityPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    $prior = Get-ItemProperty -LiteralPath $securityPath -ErrorAction Stop
    $hadAccess = $null -ne $prior.PSObject.Properties['AccessVBOM']
    $priorAccess = $prior.AccessVBOM
    if ($priorAccess -ne 1 -and -not $AllowTemporaryVbaAccess) { throw 'Explicit -AllowTemporaryVbaAccess is required to temporarily enable trusted VBA access.' }
    $probeProcess = $null
    $otherBook = $null
    $excel = $null; $book = $null; $form = $null; $session = $null
    function Invoke-Session([hashtable]$Fields) {
        $request = New-Object VBAi.Request
        foreach ($key in $Fields.Keys) { $request.$key = $Fields[$key] }
        $response = $session.Execute($request)
        if (-not $response.Ok) { throw $response.Error }
        return $response.Data
    }
    try {
        if ($priorAccess -ne 1) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null }
        $excel = New-Object -ComObject Excel.Application
        $probeProcess = Get-Process EXCEL -ErrorAction Stop
        if (@($probeProcess).Count -ne 1) { throw "Excel isolation was lost." }
        $excel.Visible = $true
        $book = $excel.Workbooks.Add()
        $vbe = $excel.GetType().InvokeMember('VBE', [Reflection.BindingFlags]::GetProperty, $null, $excel, $null)
        $vbe.MainWindow.Visible = $true
        $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
        $project = $book.VBProject
        $actions = @('align_left','align_right','align_top','align_bottom','align_centers','align_middles','same_width','same_height','same_size','center_horizontal','center_vertical','snap_grid','distribute_horizontal','distribute_vertical','space_horizontal','space_vertical','increase_horizontal_spacing','increase_vertical_spacing','decrease_horizontal_spacing','decrease_vertical_spacing')
        $proof = @(); $index = 0
        function Read-Boxes($Container) {
            return @(0..2 | ForEach-Object { $c = $Container.Controls.Item("Button$_"); [pscustomobject]@{ Left = [double]$c.Left; Top = [double]$c.Top; Width = [double]$c.Width; Height = [double]$c.Height } })
        }
        function Assert-Near([double]$Actual, [double]$Expected, [string]$Name) {
            if ([Math]::Abs($Actual - $Expected) -gt 0.1) { throw "$Name : expected $Expected, got $Actual" }
        }
        foreach ($kind in @('form', 'frame', 'page')) {
            foreach ($action in $actions) {
                $form = $project.VBComponents.Add(3); $form.Name = ('LayoutProbe' + $index); $index++
                $form.Properties.Item('Width').Value = 500; $form.Properties.Item('Height').Value = 400
                $container = $form.Designer; $prefix = 'Controls/'
                if ($kind -eq 'frame') {
                    $container = $form.Designer.Controls.Add('Forms.Frame.1','Frame1',$true)
                    $container.Width = 400; $container.Height = 300
                    $prefix = 'Controls/Frame1/Controls/'
                } elseif ($kind -eq 'page') {
                    $multi = $form.Designer.Controls.Add('Forms.MultiPage.1','MultiPage1',$true)
                    $multi.Width = 400; $multi.Height = 300
                    $container = $multi.Pages.Item(0); $container.Name = 'Page1'
                    $prefix = 'Controls/MultiPage1/Pages/Page1/Controls/'
                }
                $positions = @(@(53,53,60,30), @(143,113,40,20), @(233,203,50,25))
                foreach ($i in 0..2) {
                    $control = $container.Controls.Add('Forms.CommandButton.1',"Button$i",$true)
                    $control.Left = $positions[$i][0]; $control.Top = $positions[$i][1]
                    $control.Width = $positions[$i][2]; $control.Height = $positions[$i][3]
                }
                $tree = Invoke-Session @{ Command = 'form_tree'; Project = $project.Name; Form = $form.Name }
                $paths = [string[]]@(0..2 | ForEach-Object { $prefix + "Button$_" })
                $result = Invoke-Session @{ Command = 'apply_form_layout'; Project = $project.Name; Form = $form.Name; ExpectedTreeVersion = $tree.TreeVersion; Items = $paths; Action = $action; Width = 10 }
                if (-not $result.Verified) { throw "Unverified layout $kind/$action" }
                $boxes = Read-Boxes $container
                foreach ($i in 0..2) {
                    $box = $boxes[$i]
                    switch ($action) {
                        'align_left' { Assert-Near $box.Left 53 $action }
                        'align_right' { Assert-Near ($box.Left + $box.Width) 113 $action }
                        'align_top' { Assert-Near $box.Top 53 $action }
                        'align_bottom' { Assert-Near ($box.Top + $box.Height) 83 $action }
                        'align_centers' { Assert-Near ($box.Left + $box.Width / 2) 83 $action }
                        'align_middles' { Assert-Near ($box.Top + $box.Height / 2) 68 $action }
                        'same_width' { Assert-Near $box.Width 60 $action }
                        'same_height' { Assert-Near $box.Height 30 $action }
                        'same_size' { Assert-Near $box.Width 60 $action; Assert-Near $box.Height 30 $action }
                        'snap_grid' { Assert-Near $box.Left (@(50,140,230)[$i]) $action; Assert-Near $box.Top (@(50,110,200)[$i]) $action }
                    }
                }
                if ($action -eq 'center_horizontal') { Assert-Near (($boxes[0].Left + $boxes[2].Left + $boxes[2].Width)/2) ($container.InsideWidth/2) $action }
                if ($action -eq 'center_vertical') { Assert-Near (($boxes[0].Top + $boxes[2].Top + $boxes[2].Height)/2) ($container.InsideHeight/2) $action }
                if ($action -eq 'distribute_horizontal') { Assert-Near ($boxes[1].Left-$boxes[0].Left-$boxes[0].Width) 40 $action; Assert-Near ($boxes[2].Left-$boxes[1].Left-$boxes[1].Width) 40 $action }
                if ($action -eq 'distribute_vertical') { Assert-Near ($boxes[1].Top-$boxes[0].Top-$boxes[0].Height) 50 $action; Assert-Near ($boxes[2].Top-$boxes[1].Top-$boxes[1].Height) 50 $action }
                if ($action -match '^(space_|increase_|decrease_)') {
                    $horizontal = $action.Contains('horizontal')
                    $originalGaps = if ($horizontal) { @(30,50) } else { @(30,70) }
                    foreach ($i in 1..2) {
                        $expectedGap = if ($action.StartsWith('space_')) { 10 } elseif ($action.StartsWith('increase_')) { $originalGaps[$i-1] + 10 } else { $originalGaps[$i-1] - 10 }
                        $actualGap = if ($horizontal) { $boxes[$i].Left - $boxes[$i-1].Left - $boxes[$i-1].Width } else { $boxes[$i].Top - $boxes[$i-1].Top - $boxes[$i-1].Height }
                        Assert-Near $actualGap $expectedGap $action
                    }
                }
                $proof += [pscustomobject]@{ Form = $form.Name; Container = $kind; Action = $action; BeforeSave = $boxes; AfterReopen = $null; Verified = $true; PersistenceVerified = $false }
            }
        }
        $book.SaveAs($document, 52); $book.Close($false); $book = $null
        $book = $excel.Workbooks.Open($document); $project = $book.VBProject
        foreach ($row in $proof) {
            $form = $project.VBComponents.Item($row.Form); $container = $form.Designer
            if ($row.Container -eq 'frame') { $container = $container.Controls.Item('Frame1') }
            if ($row.Container -eq 'page') { $container = $container.Controls.Item('MultiPage1').Pages.Item('Page1') }
            $after = Read-Boxes $container
            foreach ($i in 0..2) { foreach ($property in @('Left','Top','Width','Height')) { Assert-Near $after[$i].$property $row.BeforeSave[$i].$property ($row.Form + ':' + $property) } }
            $row.AfterReopen = $after; $row.PersistenceVerified = $true
        }
        [pscustomobject]@{ Document = $document; Cases = $proof; Count = $proof.Count; Invocation = 'Current VbeSession against native Excel COM' } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath ($document + '.proof.json') -Encoding UTF8
        Write-Output "Verified $($proof.Count) native layouts after save/reopen. Evidence: $document.proof.json"
    } finally {
        # Restore trust before Quit, which can block in some COM teardown paths.
        if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
        else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
        if ($null -ne $otherBook) { $otherBook.Close($false) }
        if ($null -ne $book) { $book.Close($false) }
        if ($null -ne $excel) { $excel.Quit() }
        if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) {
            # This process was created after the no-Excel precondition and contains only
            # the disposable workbook. Quit can leave it alive because of COM references.
            Stop-Process -Id $probeProcess.Id -Force
        }
    }
}

function Invoke-VisualStudioWinForms {
    param([Parameter(Mandatory=$true)][int]$VisualStudioProcessId,
        [string]$OutputDirectory = 'artifacts/designer-compatibility/visual-studio')
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
    Add-Type @"
using System; using System.Runtime.InteropServices; using System.Runtime.InteropServices.ComTypes;
public static class DesignerRot {
 [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int reserved,out IRunningObjectTable table);
 [DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved,out IBindCtx context);
 public static object Find(string suffix) {
  IRunningObjectTable table; GetRunningObjectTable(0,out table); IEnumMoniker iter; table.EnumRunning(out iter);
  IBindCtx context; CreateBindCtx(0,out context); var item=new IMoniker[1];
  while(iter.Next(1,item,IntPtr.Zero)==0) { string name; item[0].GetDisplayName(context,null,out name);
   if(name.StartsWith("!VisualStudio.DTE.") && name.EndsWith(suffix)) { object result; table.GetObject(item[0],out result); return result; }
  } return null;
 }
}
"@
    $process = Get-Process -Id $VisualStudioProcessId
    $publicAssemblies = Join-Path (Split-Path $process.Path) 'PublicAssemblies'
    $envdtePath = Join-Path $publicAssemblies 'envdte.dll'
    $interopPath = Join-Path $publicAssemblies 'Microsoft.VisualStudio.Interop.dll'
    Add-Type -Path $interopPath
    Add-Type -Path $envdtePath
    Add-Type -AssemblyName System.Windows.Forms,System.Design
    Add-Type -ReferencedAssemblies @($envdtePath,$interopPath,'System.Windows.Forms','System.Design') -TypeDefinition @"
using System; using System.Text; using System.Threading; using System.ComponentModel.Design; using EnvDTE;
public static class VisualStudioFormValidation {
 [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool ShowWindow(IntPtr window,int command);
 public static string Verify(object value,string root) {
  var dte=(DTE)value;
  if(!string.Equals(dte.Solution.FullName,System.IO.Path.Combine(root,"VBAi.sln"),StringComparison.OrdinalIgnoreCase))
   throw new InvalidOperationException("The selected Visual Studio must have this checkout's solution open.");
  var output=new StringBuilder();
  bool wasVisible=dte.MainWindow.Visible; dte.MainWindow.Visible=true;
  try {
  foreach(var rel in new[]{"Llm/Chat/ChatWindow.cs","Editor/ModernEditorWindow.cs","Updates/UpdateWindow.cs","Updates/UpdateProgressWindow.cs","Git/GitWindow.cs","Llm/Settings/LlmSettingsWindow.cs"}) {
   var path=System.IO.Path.Combine(root,"src/VBAi",rel.Replace('/','\\'));
   var item=dte.Solution.FindProjectItem(path);
   if(item==null)throw new InvalidOperationException("Project item not loaded: "+rel);
   VerifyItem(item,output);
  }
  foreach(var rel in new[]{"Editor/ModernEditorWindow.Debug.cs","Editor/ModernEditorWindow.Language.cs","Editor/ModernEditorWindow.Save.cs","Editor/ModernEditorWindow.Tools.cs","Git/GitWindow.Review.cs","Git/GitWindow.Views.cs","Llm/Settings/LlmSettingsWindow.Views.cs"}) {
   var item=dte.Solution.FindProjectItem(System.IO.Path.Combine(root,"src/VBAi",rel.Replace('/','\\')));
   if(item==null || !string.Equals((string)item.Properties.Item("SubType").Value,"Code",StringComparison.OrdinalIgnoreCase))
    throw new InvalidOperationException("Partial file is not classified as Code: "+rel);
   output.AppendLine("PASS Code: "+rel);
  }
  return output.ToString();
  } finally { if(!wasVisible) ShowWindow((IntPtr)dte.MainWindow.HWnd,0); }
 }
 static void VerifyItem(ProjectItem item,StringBuilder output) {
  Window window=item.Open(Constants.vsViewKindDesigner);
   window.Activate(); string expected=System.IO.Path.GetFileNameWithoutExtension(item.Name);
   var deadline=DateTime.UtcNow.AddSeconds(30); string loaded=null; string observed="no host";
   while(DateTime.UtcNow<deadline) {
    try {
     System.Windows.Forms.Application.DoEvents();
     var host=window.Object as IDesignerHost;
     observed=host==null?"no host": "loading="+host.Loading+" root="+(host.RootComponent==null?"null":host.RootComponent.Site.Name);
     if(host!=null && !host.Loading && host.RootComponent!=null && host.RootComponent.Site.Name==expected) { loaded=expected; break; }
    } catch(System.Runtime.InteropServices.COMException) { }
    System.Threading.Thread.Sleep(200);
   }
   if(loaded==null) {
    throw new InvalidOperationException("Designer did not load a named root: "+item.ContainingProject.Name+"/"+item.Name+" "+observed);
   }
   output.AppendLine("PASS Designer: "+item.ContainingProject.Name+"/"+item.Name); Console.WriteLine("PASS Designer: "+item.ContainingProject.Name+"/"+item.Name);
 }
}
"@
    $dte = [DesignerRot]::Find(':' + $VisualStudioProcessId)
    if ($null -eq $dte) { throw 'The specified Visual Studio process is not registered in ROT.' }
    $result = [VisualStudioFormValidation]::Verify($dte,$repository)
    $directory = [IO.Path]::GetFullPath((Join-Path $repository $OutputDirectory))
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    $result | Set-Content -LiteralPath (Join-Path $directory 'designer-load.txt') -Encoding UTF8
    Write-Output $result
}

$commands = @{
    'ChatControlActivation' = 'Invoke-ChatControlActivation'
    'ExcelFormLayouts' = 'Invoke-ExcelFormLayouts'
    'VisualStudioWinForms' = 'Invoke-VisualStudioWinForms'
}
$command = Get-Command $commands[$Scenario] -CommandType Function
$arguments = Get-UiScenarioParameters $command $PSBoundParameters
& $command @arguments

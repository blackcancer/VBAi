param([Parameter(Mandatory=$true)][int]$VisualStudioProcessId,
    [string]$OutputDirectory = 'artifacts/designer-compatibility/visual-studio')
$ErrorActionPreference = 'Stop'
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
  if(!string.Equals(dte.Solution.FullName,System.IO.Path.Combine(root,"CodexVBE.sln"),StringComparison.OrdinalIgnoreCase))
   throw new InvalidOperationException("The selected Visual Studio must have this checkout's solution open.");
  var output=new StringBuilder();
  bool wasVisible=dte.MainWindow.Visible; dte.MainWindow.Visible=true;
  try {
  foreach(var rel in new[]{"Llm/Chat/ChatWindow.cs","Editor/ModernEditorWindow.cs","Updates/UpdateWindow.cs","Updates/UpdateProgressWindow.cs","Git/GitWindow.cs","Llm/Settings/LlmSettingsWindow.cs"}) {
   var path=System.IO.Path.Combine(root,"src/CodexVBE",rel.Replace('/','\\'));
   var item=dte.Solution.FindProjectItem(path);
   if(item==null)throw new InvalidOperationException("Project item not loaded: "+rel);
   VerifyItem(item,output);
  }
  foreach(var rel in new[]{"Editor/ModernEditorWindow.Debug.cs","Editor/ModernEditorWindow.Language.cs","Editor/ModernEditorWindow.Save.cs","Editor/ModernEditorWindow.Tools.cs","Git/GitWindow.Review.cs","Git/GitWindow.Views.cs","Llm/Settings/LlmSettingsWindow.Views.cs"}) {
   var item=dte.Solution.FindProjectItem(System.IO.Path.Combine(root,"src/CodexVBE",rel.Replace('/','\\')));
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

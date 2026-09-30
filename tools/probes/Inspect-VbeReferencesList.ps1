param([Parameter(Mandatory=$true)][long]$ListHandle)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$root=[System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$ListHandle)
if(-not $root){throw 'ListBox UI Automation root absent'}
$items=$root.FindAll([System.Windows.Automation.TreeScope]::Children,[System.Windows.Automation.Condition]::TrueCondition)
Write-Output "ListChildren=$($items.Count)"
for($i=0;$i -lt [Math]::Min($items.Count,40);$i++){
    $item=$items.Item($i)
    $toggle=$null
    $hasToggle=$item.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern,[ref]$toggle)
    $state=if($hasToggle){$toggle.Current.ToggleState}else{'n/a'}
    Write-Output "ITEM Type=$($item.Current.ControlType.ProgrammaticName) Name=$($item.Current.Name) Toggle=$state"
}
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class ReferenceListWin32 {
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr hwnd,uint msg,IntPtr wp,IntPtr lp);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr hwnd,uint msg,IntPtr wp,StringBuilder result);
  public static string[] Read(IntPtr hwnd){int count=SendMessage(hwnd,0x018B,IntPtr.Zero,IntPtr.Zero).ToInt32();int max=Math.Min(count,30);var rows=new string[max+1];rows[0]="LB_COUNT="+count;for(int i=0;i<max;i++){int length=SendMessage(hwnd,0x018A,new IntPtr(i),IntPtr.Zero).ToInt32();var text=new StringBuilder(Math.Max(1,length+1));SendMessage(hwnd,0x0189,new IntPtr(i),text);rows[i+1]=i+" "+text;}return rows;}
}
'@
[ReferenceListWin32]::Read([IntPtr]$ListHandle)

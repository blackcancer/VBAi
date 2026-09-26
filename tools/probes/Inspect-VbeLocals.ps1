param([Parameter(Mandatory = $true)] [int] $HostProcessId,
    [string] $PaneName = 'Variables locales')

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class VbeLocalsWindows {
    public delegate bool EnumWindowsProc(IntPtr handle, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr data);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr handle, StringBuilder text, int capacity);
}
'@

$script:vbeWindow = [IntPtr]::Zero
$callback = [VbeLocalsWindows+EnumWindowsProc]{
    param($handle, $data)
    [uint32]$owner = 0
    [void][VbeLocalsWindows]::GetWindowThreadProcessId($handle, [ref]$owner)
    if ($owner -eq $HostProcessId) {
        $class = [Text.StringBuilder]::new(128)
        [void][VbeLocalsWindows]::GetClassName($handle, $class, $class.Capacity)
        if ($class.ToString() -eq 'wndclass_desked_gsk') { $script:vbeWindow = $handle }
    }
    return $true
}
[void][VbeLocalsWindows]::EnumWindows($callback, [IntPtr]::Zero)
if ($script:vbeWindow -eq [IntPtr]::Zero) { throw 'VBE window not found.' }

$root = [Windows.Automation.AutomationElement]::FromHandle($script:vbeWindow)
$condition = [Windows.Automation.PropertyCondition]::new(
    [Windows.Automation.AutomationElement]::NameProperty, $PaneName)
$panes = $root.FindAll([Windows.Automation.TreeScope]::Descendants, $condition)
Write-Output "Pane [$PaneName] elements: $($panes.Count)"
foreach ($pane in $panes) {
    Write-Output "PANE $($pane.Current.ControlType.ProgrammaticName) HWND=$($pane.Current.NativeWindowHandle)"
    $children = $pane.FindAll([Windows.Automation.TreeScope]::Descendants,
        [Windows.Automation.Condition]::TrueCondition)
    Write-Output "Descendants: $($children.Count)"
    for ($index = 0; $index -lt [Math]::Min($children.Count, 80); $index++) {
        $element = $children.Item($index)
        try {
            $type = $element.Current.ControlType.ProgrammaticName
            $name = $element.Current.Name
            $id = $element.Current.AutomationId
            $value = ''
            $pattern = $null
            if ($element.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) {
                $value = ([Windows.Automation.ValuePattern]$pattern).Current.Value
            }
            $text = ''
            $pattern = $null
            if ($element.TryGetCurrentPattern([Windows.Automation.TextPattern]::Pattern, [ref]$pattern)) {
                $text = ([Windows.Automation.TextPattern]$pattern).DocumentRange.GetText(-1)
            }
            Write-Output "ITEM $index $type name=[$name] id=[$id] value=[$value] text=[$text]"
        }
        catch { Write-Output "ITEM $index ERROR $($_.Exception.Message)" }
    }
}

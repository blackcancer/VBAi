param([Parameter(Mandatory = $true)] [int] $HostProcessId,
    [string] $NamePattern = 'Pile|Calls|Espion|Watch|Variables|Locals',
    [switch] $Invoke,
    [switch] $ClickNativeButton)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class DebugTreeWindows {
    public delegate bool EnumWindowsProc(IntPtr handle, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr data);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr handle, StringBuilder text, int capacity);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);
}
'@
$rootHandle = [IntPtr]::Zero
$callback = [DebugTreeWindows+EnumWindowsProc]{
    param($handle, $data)
    [uint32]$owner = 0
    [void][DebugTreeWindows]::GetWindowThreadProcessId($handle, [ref]$owner)
    if ($owner -eq $HostProcessId) {
        $class = [Text.StringBuilder]::new(128)
        [void][DebugTreeWindows]::GetClassName($handle, $class, $class.Capacity)
        if ($class.ToString() -eq 'wndclass_desked_gsk') { $script:rootHandle = $handle }
    }
    return $true
}
[void][DebugTreeWindows]::EnumWindows($callback, [IntPtr]::Zero)
if ($rootHandle -eq [IntPtr]::Zero) { throw 'VBE window not found.' }
$root = [Windows.Automation.AutomationElement]::FromHandle($rootHandle)
$all = $root.FindAll([Windows.Automation.TreeScope]::Descendants, [Windows.Automation.Condition]::TrueCondition)
$hits = [System.Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $all.Count; $i++) {
    $item = $all.Item($i)
    try {
        $name = $item.Current.Name
        if ($name -notmatch $NamePattern) { continue }
        $pattern = $null
        $canInvoke = $item.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)
        $className = [Text.StringBuilder]::new(128)
        $handle = [IntPtr]$item.Current.NativeWindowHandle
        if ($handle -ne [IntPtr]::Zero) { [void][DebugTreeWindows]::GetClassName($handle, $className, $className.Capacity) }
        $hits.Add([pscustomobject]@{ Index = $i; Name = $name; Type = $item.Current.ControlType.ProgrammaticName;
            Handle = $item.Current.NativeWindowHandle; Class = $className.ToString();
            Enabled = $item.Current.IsEnabled; CanInvoke = $canInvoke })
        if ($Invoke -and $canInvoke -and $item.Current.IsEnabled) {
            ([Windows.Automation.InvokePattern]$pattern).Invoke()
            break
        }
        if ($ClickNativeButton -and $className.ToString() -eq 'Button' -and $item.Current.IsEnabled) {
            if (-not [DebugTreeWindows]::PostMessage($handle, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)) {
                throw 'BM_CLICK was rejected.'
            }
            break
        }
    }
    catch { }
}
$hits.ToArray() | ConvertTo-Json -Depth 5

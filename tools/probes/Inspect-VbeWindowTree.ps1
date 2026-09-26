param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class CodexVbeWindowTree {
    public delegate bool EnumProc(IntPtr handle, IntPtr state);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc callback, IntPtr state);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr state);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr handle, StringBuilder text, int length);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr handle, StringBuilder text, int length);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr handle);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr handle);
}
'@

$root = [IntPtr]::Zero
$findRoot = [CodexVbeWindowTree+EnumProc]{
    param($handle, $state)
    [uint32]$windowProcessId = 0
    [void][CodexVbeWindowTree]::GetWindowThreadProcessId($handle, [ref]$windowProcessId)
    if ($windowProcessId -ne $HostProcessId) { return $true }
    $class = [Text.StringBuilder]::new(256)
    [void][CodexVbeWindowTree]::GetClassName($handle, $class, $class.Capacity)
    if ($class.ToString() -eq 'wndclass_desked_gsk') {
        $script:root = $handle
        return $false
    }
    return $true
}
[void][CodexVbeWindowTree]::EnumWindows($findRoot, [IntPtr]::Zero)
if ($root -eq [IntPtr]::Zero) { throw "VBE window not found for host PID $HostProcessId." }

$rows = [Collections.Generic.List[object]]::new()
$inspect = {
    param([IntPtr]$handle)
    $class = [Text.StringBuilder]::new(256)
    $title = [Text.StringBuilder]::new(512)
    [void][CodexVbeWindowTree]::GetClassName($handle, $class, $class.Capacity)
    [void][CodexVbeWindowTree]::GetWindowText($handle, $title, $title.Capacity)
    $rows.Add([pscustomobject]@{
        Handle = $handle.ToInt64()
        Parent = [CodexVbeWindowTree]::GetParent($handle).ToInt64()
        Class = $class.ToString()
        Title = $title.ToString()
        Visible = [CodexVbeWindowTree]::IsWindowVisible($handle)
    })
}
& $inspect $root
$visit = [CodexVbeWindowTree+EnumProc]{
    param($handle, $state)
    & $inspect $handle
    return $true
}
[void][CodexVbeWindowTree]::EnumChildWindows($root, $visit, [IntPtr]::Zero)
$rows | ConvertTo-Json -Depth 4 -Compress

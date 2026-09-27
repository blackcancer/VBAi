param([Parameter(Mandatory = $true)] [int] $HostProcessId,
    [string[]] $PaneNames = @('Variables locales', ('Ex' + [char]0x00E9 + 'cution')))

$ErrorActionPreference = 'Stop'
Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class CodexVbePaneWindows {
    public delegate bool EnumWindowsProc(IntPtr handle, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr data);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr handle, StringBuilder text, int capacity);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr handle, StringBuilder text, int capacity);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr handle);
    public static string Class(IntPtr handle) { var b = new StringBuilder(256); GetClassName(handle,b,b.Capacity); return b.ToString(); }
    public static string Title(IntPtr handle) { var b = new StringBuilder(256); GetWindowText(handle,b,b.Capacity); return b.ToString(); }
    public static IntPtr FindRoot(uint pid) {
        IntPtr found = IntPtr.Zero;
        EnumWindows((h,d) => { uint owner; GetWindowThreadProcessId(h,out owner);
            if (owner == pid && Class(h) == "wndclass_desked_gsk" && IsWindowVisible(h)) { found=h; return false; }
            return true; }, IntPtr.Zero);
        return found;
    }
    public static IntPtr[] Children(IntPtr parent) {
        var result = new List<IntPtr>();
        EnumChildWindows(parent,(h,d) => { result.Add(h); return true; },IntPtr.Zero);
        return result.ToArray();
    }
}
'@
$root = [CodexVbePaneWindows]::FindRoot([uint32]$HostProcessId)
if ($root -eq [IntPtr]::Zero) { throw 'Visible VBE root was not found in the exact host process.' }
$all = [CodexVbePaneWindows]::Children($root)
foreach ($name in $PaneNames) {
    $matches = @($all | Where-Object { [CodexVbePaneWindows]::Class($_) -eq 'VbaWindow' -and
        [CodexVbePaneWindows]::Title($_) -eq $name -and [CodexVbePaneWindows]::IsWindowVisible($_) })
    Write-Output "PANE [$name] count=$($matches.Count)"
    foreach ($pane in $matches) {
        Write-Output "ROOT HWND=$($pane.ToInt64()) class=$([CodexVbePaneWindows]::Class($pane))"
        $children = [CodexVbePaneWindows]::Children($pane)
        Write-Output "CHILDREN count=$($children.Count)"
        foreach ($child in $children) {
            Write-Output "HWND=$($child.ToInt64()) class=$([CodexVbePaneWindows]::Class($child)) title=[$([CodexVbePaneWindows]::Title($child))] visible=$([CodexVbePaneWindows]::IsWindowVisible($child))"
        }
    }
}

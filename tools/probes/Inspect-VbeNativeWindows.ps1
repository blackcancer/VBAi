param(
    [Parameter(Mandatory = $true)] [ValidateRange(1, [int]::MaxValue)] [int] $HostProcessId,
    [ValidateSet('Tree', 'Windows', 'Children', 'Panes')] [string] $View = 'Tree',
    [long] $WindowHandle,
    [string[]] $PaneNames = @('Variables locales', ('Ex' + [char]0x00E9 + 'cution'))
)

$ErrorActionPreference = 'Stop'
if ($View -eq 'Children' -and -not $WindowHandle) { throw 'Children requires an explicit owned WindowHandle.' }
if ($View -ne 'Children' -and $WindowHandle) { throw 'WindowHandle is supported only for Children.' }
Get-Process -Id $HostProcessId -ErrorAction Stop | Out-Null
if (-not ('VBAi.Tools.NativeWindowInventory' -as [type])) {
    Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi.Tools
{
    public sealed class NativeWindowRow
    {
        public long Handle { get; set; }
        public long Parent { get; set; }
        public string Class { get; set; }
        public string Title { get; set; }
        public bool Visible { get; set; }
    }

    public static class NativeWindowInventory
    {
        private delegate bool EnumProc(IntPtr handle, IntPtr state);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc callback, IntPtr state);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr state);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint pid);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr handle, StringBuilder text, int length);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr handle, StringBuilder text, int length);
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr handle);

        private static NativeWindowRow Read(IntPtr handle)
        {
            var kind = new StringBuilder(256);
            var title = new StringBuilder(512);
            GetClassName(handle, kind, kind.Capacity);
            GetWindowText(handle, title, title.Capacity);
            return new NativeWindowRow { Handle = handle.ToInt64(), Parent = GetParent(handle).ToInt64(),
                Class = kind.ToString(), Title = title.ToString(), Visible = IsWindowVisible(handle) };
        }

        private static void RequireOwner(IntPtr handle, int expected)
        {
            uint actual;
            if (handle == IntPtr.Zero || !IsWindow(handle) || GetWindowThreadProcessId(handle, out actual) == 0 || actual != expected)
                throw new InvalidOperationException("The selected native window no longer belongs to the explicit host process.");
        }

        public static NativeWindowRow[] Inspect(int pid, string view, long selected)
        {
            var rows = new List<NativeWindowRow>();
            if (view == "Windows")
            {
                EnumWindows((handle, state) => { uint owner; GetWindowThreadProcessId(handle, out owner);
                    if (owner == pid) rows.Add(Read(handle)); return true; }, IntPtr.Zero);
                return rows.ToArray();
            }
            IntPtr root = new IntPtr(selected);
            if (view == "Tree")
            {
                var roots = new List<IntPtr>();
                EnumWindows((handle, state) => { uint owner; GetWindowThreadProcessId(handle, out owner);
                    if (owner == pid && Read(handle).Class == "wndclass_desked_gsk") roots.Add(handle);
                    return true; }, IntPtr.Zero);
                if (roots.Count != 1) throw new InvalidOperationException("One exact VBE root is required for the selected host process.");
                root = roots[0];
            }
            else if (view != "Children") throw new ArgumentException("Unknown inventory view.");
            RequireOwner(root, pid);
            if (view == "Tree") rows.Add(Read(root));
            EnumChildWindows(root, (handle, state) => { uint owner; GetWindowThreadProcessId(handle, out owner);
                if (owner == pid) rows.Add(Read(handle)); return true; }, IntPtr.Zero);
            RequireOwner(root, pid);
            return rows.ToArray();
        }
    }
}
'@
}
$nativeView = if ($View -eq 'Panes') { 'Tree' } else { $View }
$rows = @([VBAi.Tools.NativeWindowInventory]::Inspect($HostProcessId, $nativeView, $WindowHandle))
if ($View -eq 'Panes') {
    foreach ($name in $PaneNames) {
        $panes = @($rows | Where-Object { $_.Class -eq 'VbaWindow' -and $_.Title -eq $name -and $_.Visible })
        Write-Output "PANE [$name] count=$($panes.Count)"
        foreach ($pane in $panes) {
            Write-Output "ROOT HWND=$($pane.Handle) class=$($pane.Class)"
            $children = @([VBAi.Tools.NativeWindowInventory]::Inspect($HostProcessId, 'Children', $pane.Handle))
            Write-Output "CHILDREN count=$($children.Count)"
            foreach ($child in $children) {
                Write-Output "HWND=$($child.Handle) class=$($child.Class) title=[$($child.Title)] visible=$($child.Visible)"
            }
        }
    }
    return
}
ConvertTo-Json -InputObject $rows -Depth 4 -Compress

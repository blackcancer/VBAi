param([Parameter(Mandatory=$true)][int]$HostProcessId)
$ErrorActionPreference = 'Stop'
Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class AssistantHandles {
    public delegate bool EnumProc(IntPtr hwnd, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc proc, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc proc, IntPtr data);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int size);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder text, int size);
    public static string[] Read(int pid) {
        var lines = new List<string>();
        EnumWindows((hwnd, data) => {
            uint owner; GetWindowThreadProcessId(hwnd, out owner);
            if (owner != pid) return true;
            var title = new StringBuilder(200); GetWindowText(hwnd, title, title.Capacity);
            if (!title.ToString().StartsWith("CodexVBE") || !title.ToString().Contains("Assistant")) return true;
            lines.Add("Window=" + hwnd + " " + title);
            EnumChildWindows(hwnd, (child, unused) => {
                var kind = new StringBuilder(100); GetClassName(child, kind, kind.Capacity);
                var name = new StringBuilder(300); GetWindowText(child, name, name.Capacity);
                lines.Add(child + " Class=" + kind + " Text=" + name);
                return true;
            }, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);
        return lines.ToArray();
    }
}
'@
[AssistantHandles]::Read($HostProcessId)

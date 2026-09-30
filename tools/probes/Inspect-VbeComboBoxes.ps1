param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class CodexVbeCombos
{
    private delegate bool EnumProc(IntPtr hwnd, IntPtr data);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc callback, IntPtr data);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr data);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int length);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder text, int length);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern IntPtr SendMessage(IntPtr hwnd, int message, IntPtr param, StringBuilder output);

    public static string[] Read(int targetPid)
    {
        var result = new List<string>();
        IntPtr assistant = IntPtr.Zero;
        EnumWindows((hwnd, data) => {
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);
            if (pid == targetPid) {
                var title = new StringBuilder(200);
                GetWindowText(hwnd, title, title.Capacity);
                if (title.ToString().StartsWith("CodexVBE") && title.ToString().Contains("Assistant")) assistant = hwnd;
            }
            return true;
        }, IntPtr.Zero);
        if (assistant == IntPtr.Zero) throw new InvalidOperationException("Assistant HWND not found.");
        EnumChildWindows(assistant, (hwnd, data) => {
            var kind = new StringBuilder(200);
            GetClassName(hwnd, kind, kind.Capacity);
            if (kind.ToString().IndexOf("COMBOBOX", StringComparison.OrdinalIgnoreCase) >= 0) {
                int count = SendMessage(hwnd, 0x146, IntPtr.Zero, null).ToInt32();
                result.Add("COMBO HWND=" + hwnd + " Count=" + count);
                for (int index = 0; index < count && index < 100; index++) {
                    int length = SendMessage(hwnd, 0x149, new IntPtr(index), null).ToInt32();
                    var label = new StringBuilder(Math.Max(1, length + 1));
                    SendMessage(hwnd, 0x148, new IntPtr(index), label);
                    result.Add("  " + index + " " + label);
                }
            }
            return true;
        }, IntPtr.Zero);
        return result.ToArray();
    }
}
'@
[CodexVbeCombos]::Read($HostProcessId)

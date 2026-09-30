param([Parameter(Mandatory=$true)][int]$HostProcessId, [string]$PromptText = 'Réponds exactement : TEST CODEXVBE OK', [switch]$AcknowledgeDisposableSession)
$ErrorActionPreference = 'Stop'
if (-not $AcknowledgeDisposableSession) { throw 'Confirm this is an owned disposable assistant session: its current context and the supplied prompt will be sent to its configured provider.' }
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class AssistantSmoke {
    public delegate bool EnumProc(IntPtr hwnd, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc proc, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc proc, IntPtr data);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int size);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder text, int size);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr wparam, string lparam);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr wparam, IntPtr lparam);
    public static string Send(int pid, string prompt) {
        IntPtr assistant = IntPtr.Zero;
        EnumWindows((hwnd, data) => {
            uint owner; GetWindowThreadProcessId(hwnd, out owner);
            if (owner == pid) { var title = new StringBuilder(200); GetWindowText(hwnd, title, title.Capacity);
                if (title.ToString().StartsWith("CodexVBE") && title.ToString().Contains("Assistant")) assistant = hwnd; }
            return true;
        }, IntPtr.Zero);
        if (assistant == IntPtr.Zero) throw new InvalidOperationException("Assistant absent");
        IntPtr input = IntPtr.Zero, send = IntPtr.Zero;
        int edits = 0;
        EnumChildWindows(assistant, (hwnd, data) => {
            var kind = new StringBuilder(100); GetClassName(hwnd, kind, kind.Capacity);
            var title = new StringBuilder(100); GetWindowText(hwnd, title, title.Capacity);
            if (kind.ToString().Contains(".EDIT.")) { edits++; if (edits == 2) input = hwnd; }
            if (kind.ToString().Contains(".BUTTON.") && title.ToString() == "Envoyer") send = hwnd;
            return true;
        }, IntPtr.Zero);
        if (input == IntPtr.Zero || send == IntPtr.Zero) throw new InvalidOperationException("Contrôles non trouvés");
        SendMessage(input, 0x000C, IntPtr.Zero, prompt);
        SendMessage(send, 0x00F5, IntPtr.Zero, IntPtr.Zero);
        return "Sent via named button and editor HWND";
    }
}
'@
[AssistantSmoke]::Send($HostProcessId, $PromptText)

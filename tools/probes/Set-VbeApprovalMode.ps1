param([Parameter(Mandatory=$true)][int]$HostProcessId, [Parameter(Mandatory=$true)][ValidateSet('Lecture seule','Demander à chaque action','Automatique')][string]$Mode, [switch]$AcknowledgePersistentApprovalChange)
$ErrorActionPreference = 'Stop'
if (-not $AcknowledgePersistentApprovalChange) { throw 'Explicit acknowledgement is required: this changes and saves the selected host approval mode.' }
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class ApprovalUi {
    public delegate bool EnumProc(IntPtr hwnd, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc proc, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc proc, IntPtr data);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int size);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder text, int size);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr wparam, IntPtr lparam);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr wparam, StringBuilder lparam);
    [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr hwnd);
    public static string Select(int pid, string mode) {
        IntPtr dialog = IntPtr.Zero;
        EnumWindows((hwnd, data) => { uint owner; GetWindowThreadProcessId(hwnd, out owner);
            if (owner == pid) { var title = new StringBuilder(200); GetWindowText(hwnd, title, title.Capacity);
                if (title.ToString().StartsWith("CodexVBE") && title.ToString().Contains("Configuration LLM")) dialog = hwnd; }
            return true; }, IntPtr.Zero);
        if (dialog == IntPtr.Zero) throw new InvalidOperationException("Settings dialog absent");
        IntPtr picker = IntPtr.Zero, save = IntPtr.Zero; int selected = -1;
        EnumChildWindows(dialog, (hwnd, data) => {
            var kind = new StringBuilder(100); GetClassName(hwnd, kind, kind.Capacity);
            if (kind.ToString().Contains(".COMBOBOX.")) {
                int count = SendMessage(hwnd, 0x146, IntPtr.Zero, IntPtr.Zero).ToInt32();
                for (int i=0; i<count; i++) { int length=SendMessage(hwnd,0x149,new IntPtr(i),IntPtr.Zero).ToInt32();
                    var label=new StringBuilder(length+1); SendMessage(hwnd,0x148,new IntPtr(i),label);
                    if(label.ToString()==mode) {picker=hwnd;selected=i;} }
            }
            if (kind.ToString().Contains(".BUTTON.")) { var label=new StringBuilder(100); GetWindowText(hwnd,label,label.Capacity);
                if(label.ToString()=="Enregistrer") save=hwnd; }
            return true;
        }, IntPtr.Zero);
        if (picker==IntPtr.Zero || save==IntPtr.Zero) throw new InvalidOperationException("Approval picker or save button absent");
        SendMessage(picker,0x14E,new IntPtr(selected),IntPtr.Zero);
        int controlId=GetDlgCtrlID(picker);
        IntPtr parent=GetParent(picker);
        SendMessage(parent,0x0111,new IntPtr((1<<16)|(controlId&0xFFFF)),picker);
        SendMessage(save,0x00F5,IntPtr.Zero,IntPtr.Zero);
        return "Selected "+mode+" and saved via named controls";
    }
}
'@
[ApprovalUi]::Select($HostProcessId, $Mode)

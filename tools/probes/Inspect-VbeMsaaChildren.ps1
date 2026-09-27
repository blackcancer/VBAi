param([Parameter(Mandatory = $true)] [int] $HostProcessId,
    [string] $PaneName = 'Variables locales')

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Accessibility
Add-Type -ReferencedAssemblies 'Accessibility' @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class CodexVbeMsaaProbe {
    public delegate bool EnumWindowsProc(IntPtr handle, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr data);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr handle, StringBuilder text, int capacity);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr handle, StringBuilder text, int capacity);
    [DllImport("oleacc.dll")] public static extern int AccessibleObjectFromWindow(IntPtr handle, uint objectId,
        ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object accessible);
    [DllImport("oleacc.dll")] private static extern int AccessibleChildren(Accessibility.IAccessible container,
        int start, int count, IntPtr variants, out int obtained);
    [DllImport("oleaut32.dll")] private static extern int VariantClear(IntPtr variant);
    private static string Class(IntPtr h) { var b = new StringBuilder(128); GetClassName(h,b,b.Capacity); return b.ToString(); }
    private static string Title(IntPtr h) { var b = new StringBuilder(256); GetWindowText(h,b,b.Capacity); return b.ToString(); }
    public static IntPtr FindPane(uint pid, string title) {
        IntPtr root = IntPtr.Zero, pane = IntPtr.Zero;
        EnumWindows((h,d) => { uint owner; GetWindowThreadProcessId(h,out owner);
            if (owner == pid && Class(h) == "wndclass_desked_gsk") { root=h; return false; } return true; }, IntPtr.Zero);
        if (root != IntPtr.Zero) EnumChildWindows(root,(h,d) => {
            if (Class(h) == "VbaWindow" && Title(h) == title) { pane=h; return false; } return true; }, IntPtr.Zero);
        return pane;
    }
    public static string[] Read(IntPtr handle) {
        var lines = new List<string>();
        Guid iid = new Guid("618736e0-3c3d-11cf-810c-00aa00389b71"); object raw;
        int hr = AccessibleObjectFromWindow(handle, 4294967292U, ref iid, out raw);
        if (hr != 0) throw new COMException("AccessibleObjectFromWindow", hr);
        var node = (Accessibility.IAccessible)raw;
        Walk(node, "ROOT", 0, lines);
        return lines.ToArray();
    }
    private static void Walk(Accessibility.IAccessible node, string prefix, int depth, List<string> lines) {
        if (depth > 3 || lines.Count >= 140) return;
        int count = Math.Min(node.accChildCount, 128);
        lines.Add(prefix + " role=" + node.get_accRole(0) + " name=[" + node.get_accName(0) +
            "] value=[" + node.get_accValue(0) + "] children=" + count);
        if (count == 0) return;
        int stride = IntPtr.Size == 8 ? 24 : 16;
        IntPtr buffer = Marshal.AllocCoTaskMem(count * stride);
        for (int i=0; i<count*stride; i++) Marshal.WriteByte(buffer,i,0);
        int obtained = 0;
        try {
            int hr = AccessibleChildren(node, 0, count, buffer, out obtained);
            lines.Add(prefix + " AccessibleChildren HRESULT=" + hr + " obtained=" + obtained);
            if (hr < 0) return;
            for (int i=0; i<obtained; i++) {
                if (lines.Count >= 140) break;
                IntPtr item = IntPtr.Add(buffer,i*stride);
                try {
                    object child = Marshal.GetObjectForNativeVariant(item);
                    if (child is int) {
                        int id=(int)child;
                        try { lines.Add(prefix + " ID=" + id + " role=" + node.get_accRole(id) +
                            " name=[" + node.get_accName(id) + "] value=[" + node.get_accValue(id) + "]"); }
                        catch (Exception ex) { lines.Add(prefix + " ID=" + id + " error=" + ex.Message); }
                    } else if (child is Accessibility.IAccessible) {
                        var sub=(Accessibility.IAccessible)child;
                        Walk(sub,prefix + "." + i,depth+1,lines);
                    } else lines.Add(prefix + " TYPE=" + (child == null ? "null" : child.GetType().FullName));
                } catch (Exception ex) { lines.Add(prefix + " INDEX=" + i + " error=" + ex.Message); }
            }
        } finally {
            for (int i=0; i<Math.Min(obtained,count); i++) VariantClear(IntPtr.Add(buffer,i*stride));
            Marshal.FreeCoTaskMem(buffer);
        }
    }
}
'@
$pane = [CodexVbeMsaaProbe]::FindPane([uint32]$HostProcessId, $PaneName)
if ($pane -eq [IntPtr]::Zero) { throw "Native pane not found: $PaneName" }
Write-Output "PANE [$PaneName] HWND=$($pane.ToInt64())"
[CodexVbeMsaaProbe]::Read($pane)

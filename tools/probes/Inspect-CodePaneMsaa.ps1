param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName Accessibility
Add-Type -ReferencedAssemblies ([Accessibility.IAccessible].Assembly.Location) @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class CodexCodePaneMsaa {
    [DllImport("oleacc.dll")]
    public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId,
        ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object result);
    public static string[] Inspect(IntPtr handle) {
        Guid iid = new Guid("618736e0-3c3d-11cf-810c-00aa00389b71");
        object raw;
        int hr = AccessibleObjectFromWindow(handle, 4294967292, ref iid, out raw);
        if (hr != 0) throw new COMException("AccessibleObjectFromWindow failed", hr);
        var accessible = (Accessibility.IAccessible)raw;
        var rows = new List<string>();
        rows.Add("root=" + accessible.get_accName(0) + ";count=" + accessible.accChildCount);
        for (int index = 1; index <= Math.Min(accessible.accChildCount, 100); index++) {
            try {
                rows.Add(index + ":" + accessible.get_accName(index) + ";role=" +
                    accessible.get_accRole(index) + ";state=" + accessible.get_accState(index));
            } catch (Exception error) {
                try {
                    object child = accessible.get_accChild(index);
                    var childAccess = child as Accessibility.IAccessible;
                    rows.Add(index + ":child=" + (child == null ? "null" : child.GetType().FullName) +
                        ";name=" + (childAccess == null ? "unavailable" : childAccess.get_accName(0)));
                } catch (Exception childError) {
                    rows.Add(index + ":error=" + error.Message + ";childError=" + childError.Message);
                }
            }
        }
        return rows.ToArray();
    }
}
'@

$root = [Windows.Automation.AutomationElement]::RootElement
$windows = $root.FindAll([Windows.Automation.TreeScope]::Descendants,
    [Windows.Automation.PropertyCondition]::new(
        [Windows.Automation.AutomationElement]::NameProperty, 'ThisWorkbook (Code)'))
$code = $null
foreach ($candidate in $windows) {
    if ($candidate.Current.ProcessId -eq $HostProcessId -and
        $candidate.Current.ControlType -eq [Windows.Automation.ControlType]::Window) {
        $code = $candidate
        break
    }
}
if ($null -eq $code) { throw 'Code window unavailable.' }
[CodexCodePaneMsaa]::Inspect([IntPtr]$code.Current.NativeWindowHandle) |
    ConvertTo-Json -Compress

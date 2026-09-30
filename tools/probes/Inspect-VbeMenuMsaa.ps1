param([Parameter(Mandatory = $true)] [long] $MenuBarHandle, [string] $OpenMenu, [string] $InvokeItem, [switch] $AcknowledgeMenuAction)

$ErrorActionPreference = 'Stop'
if (($OpenMenu -or $InvokeItem) -and -not $AcknowledgeMenuAction) { throw 'Explicit acknowledgement is required before opening or invoking a menu in the selected window.' }
Add-Type -AssemblyName Accessibility
Add-Type -ReferencedAssemblies Accessibility @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Accessibility;

public static class VbeMsaa
{
    [DllImport("oleacc.dll", PreserveSig = false)]
    private static extern void AccessibleObjectFromWindow(IntPtr hwnd, uint objectId,
        ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out IAccessible accessible);
    [DllImport("oleacc.dll")]
    private static extern int AccessibleChildren(IAccessible container, int start, int count,
        [Out, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.Struct)] object[] children,
        out int obtained);

    public static string[] Inspect(IntPtr hwnd, string openMenu, string invokeItem)
    {
        Guid iid = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");
        IAccessible accessible;
        AccessibleObjectFromWindow(hwnd, 0xFFFFFFFC, ref iid, out accessible);
        var names = new List<string>();
        names.Add("Root=" + accessible.get_accName(0) + " Children=" + accessible.accChildCount);
        object[] children = new object[accessible.accChildCount];
        int obtained;
        int hr = AccessibleChildren(accessible, 0, children.Length, children, out obtained);
        names.Add("AccessibleChildren HR=" + hr.ToString("X8") + " Obtained=" + obtained);
        for (int index = 0; index < obtained; index++)
        {
            try {
                IAccessible child = children[index] as IAccessible;
                if (child != null) {
                    string name = child.get_accName(0);
                    names.Add((index + 1) + " OBJECT " + name + " Role=" + child.get_accRole(0));
                    if (string.Equals(name, openMenu, StringComparison.OrdinalIgnoreCase)) {
                        child.accDoDefaultAction(0);
                        names.Add("OPENED " + name + " Children=" + child.accChildCount);
                        object[] menuItems = new object[child.accChildCount];
                        int menuObtained;
                        int menuHr = AccessibleChildren(child, 0, menuItems.Length, menuItems, out menuObtained);
                        names.Add("MenuChildren HR=" + menuHr.ToString("X8") + " Obtained=" + menuObtained);
                        for (int menuIndex = 0; menuIndex < menuObtained; menuIndex++) {
                            IAccessible menuChild = menuItems[menuIndex] as IAccessible;
                            if (menuChild != null) {
                                string itemName = menuChild.get_accName(0);
                                names.Add("  " + (menuIndex + 1) + " " + itemName);
                                if (!string.IsNullOrEmpty(invokeItem) && itemName.StartsWith(invokeItem, StringComparison.OrdinalIgnoreCase)) {
                                    menuChild.accDoDefaultAction(0);
                                    names.Add("INVOKED " + itemName);
                                }
                            }
                            else names.Add("  " + (menuIndex + 1) + " ID=" + menuItems[menuIndex]);
                        }
                    }
                }
                else names.Add((index + 1) + " ID " + children[index] + " " + accessible.get_accName(children[index]));
            }
            catch (Exception error) { names.Add(index + " ERROR " + error.Message); }
        }
        return names.ToArray();
    }
}
'@
[VbeMsaa]::Inspect([IntPtr]$MenuBarHandle, $OpenMenu, $InvokeItem)

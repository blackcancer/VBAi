param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('CodePane', 'CodePaneMsaa', 'CodeSelection', 'DebugTree', 'Locals',
        'ObjectBrowser', 'Menu', 'MsaaChildren', 'References')]
    [string] $Target,
    [Parameter(Mandatory = $true)] [ValidateRange(1, [int]::MaxValue)] [int] $HostProcessId,
    [Alias('ListHandle', 'MenuBarHandle')] [long] $WindowHandle,
    [string] $CodeWindowName = 'ThisWorkbook (Code)',
    [string] $PaneName = 'Variables locales',
    [string] $NamePattern = 'Pile|Calls|Espion|Watch|Variables|Locals',
    [string] $BrowserName = "Explorateur d'objets",
    [string] $Class,
    [string] $Member,
    [string] $OpenMenu,
    [string] $InvokeItem,
    [switch] $Invoke,
    [switch] $ClickNativeButton,
    [switch] $AcknowledgeUiAction
)

$ErrorActionPreference = 'Stop'
Get-Process -Id $HostProcessId -ErrorAction Stop | Out-Null
if (($Invoke -or $ClickNativeButton) -and $Target -ne 'DebugTree') {
    throw 'Invoke and ClickNativeButton are supported only for DebugTree.'
}
if ($Invoke -and $ClickNativeButton) { throw 'Select exactly one activation method.' }
if (($OpenMenu -or $InvokeItem) -and $Target -ne 'Menu') {
    throw 'Menu actions are supported only for Menu.'
}
if ($InvokeItem -and -not $OpenMenu) { throw 'InvokeItem requires an explicit OpenMenu.' }
if (($Invoke -or $ClickNativeButton -or $OpenMenu -or $InvokeItem -or $Target -eq 'ObjectBrowser') -and
    -not $AcknowledgeUiAction) {
    throw 'Use -AcknowledgeUiAction before changing a selection or invoking a control.'
}
if ($Target -eq 'ObjectBrowser' -and (-not $Class -or -not $Member)) {
    throw 'ObjectBrowser requires Class and Member.'
}
if ($Target -in @('Menu', 'References') -and -not $WindowHandle) {
    throw 'Menu and References require an explicit owned WindowHandle.'
}
if ($WindowHandle -and $Target -notin @('Menu', 'References')) {
    throw 'WindowHandle is supported only for Menu and References.'
}
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, Accessibility
if ($WindowHandle) {
    $owned = [Windows.Automation.AutomationElement]::FromHandle([IntPtr]$WindowHandle)
    if (-not $owned -or $owned.Current.ProcessId -ne $HostProcessId) {
        throw 'WindowHandle does not belong to the selected host.'
    }
}

function Add-InspectionType([string] $Name, [string] $Definition) {
    if (-not ($Name -as [type])) {
        Add-Type -TypeDefinition $Definition -ReferencedAssemblies ([Accessibility.IAccessible].Assembly.Location)
    }
}

function Get-VbeRows {
    if (-not $script:inspectionRows) {
        $script:inspectionRows = @(& (Join-Path $PSScriptRoot 'Inspect-VbeNativeWindows.ps1') -HostProcessId $HostProcessId -View Tree | ConvertFrom-Json)
    }
    return $script:inspectionRows
}

function Get-VbeAutomationRoot {
    $matches = @(Get-VbeRows | Where-Object { $_.Class -eq 'wndclass_desked_gsk' })
    if ($matches.Count -ne 1) { throw 'One exact VBE root is required.' }
    $root = [Windows.Automation.AutomationElement]::FromHandle([IntPtr]$matches[0].Handle)
    if (-not $root -or $root.Current.ProcessId -ne $HostProcessId) { throw 'VBE ownership changed.' }
    return $root
}

Add-InspectionType 'VBAiInspectionNative' @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class VBAiInspectionNative {
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr handle, StringBuilder text, int capacity);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr handle);
    [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr handle);
}
'@

function Read-CodePane {
    $vbe = Get-VbeAutomationRoot
    $code = $vbe.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.AndCondition]::new(
            [System.Windows.Automation.PropertyCondition]::new(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                [System.Windows.Automation.ControlType]::Document),
            [System.Windows.Automation.PropertyCondition]::new(
                [System.Windows.Automation.AutomationElement]::NameProperty, $CodeWindowName)))
    if ($null -eq $code) { throw 'ThisWorkbook code document not found.' }
    $walker = [System.Windows.Automation.TreeWalker]::RawViewWalker
    $container = $walker.GetParent($code)
    if ($null -eq $container) { throw 'Code document parent unavailable.' }
    $items = $container.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    $rows = [Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt [Math]::Min($items.Count, 200); $index++) {
        $item = $items.Item($index)
        try {
            $rows.Add([pscustomobject]@{
                    Name = $item.Current.Name
                    Type = $item.Current.ControlType.ProgrammaticName
                    Class = $item.Current.ClassName
                    AutomationId = $item.Current.AutomationId
                    Handle = $item.Current.NativeWindowHandle
                    Patterns = @($item.GetSupportedPatterns() | ForEach-Object { $_.ProgrammaticName })
            })
        }
        catch { $rows.Add([pscustomobject]@{ Error = $_.Exception.Message }) }
    }
    [pscustomobject]@{
        HostProcessId = $HostProcessId
        ContainerName = $container.Current.Name
        ContainerType = $container.Current.ControlType.ProgrammaticName
        ContainerHandle = $container.Current.NativeWindowHandle
        TotalDescendants = $items.Count
        Items = $rows.ToArray()
    } | ConvertTo-Json -Depth 6 -Compress
}

function Read-CodePaneMsaa {
    Add-InspectionType 'VBAiCodePaneMsaa' @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class VBAiCodePaneMsaa {
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

    $root = Get-VbeAutomationRoot
    $condition = [System.Windows.Automation.AndCondition]::new(
        [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, $CodeWindowName),
        [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window))
    $matches = @($root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition) | Where-Object { $_.Current.ProcessId -eq $HostProcessId })
    if ($matches.Count -ne 1) { throw 'One exact owned code window is required.' }
    $code = $matches[0]
    [VBAiCodePaneMsaa]::Inspect([IntPtr]$code.Current.NativeWindowHandle) |
    ConvertTo-Json -Compress
}

function Read-CodeSelection {
    $root = Get-VbeAutomationRoot
    $all = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    $rows = [Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $all.Count; $i++) {
        $element = $all.Item($i)
        try {
            $patternObject = $null
            if (-not $element.TryGetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern,
                    [ref]$patternObject)) { continue }
            $pattern = [System.Windows.Automation.TextPattern]$patternObject
            $selection = @($pattern.GetSelection())
            $selectedText = if ($selection.Count -gt 0) { $selection[0].GetText(300) } else { '' }
            $range = if ($selection.Count -gt 0) { $selection[0] } else { $pattern.DocumentRange }
            $foreground = $range.GetAttributeValue([System.Windows.Automation.TextPattern]::ForegroundColorAttribute)
            $background = $range.GetAttributeValue([System.Windows.Automation.TextPattern]::BackgroundColorAttribute)
            $snippet = $pattern.DocumentRange.GetText(300)
            $rows.Add([pscustomobject]@{
                    Index = $i
                    Name = $element.Current.Name
                    Class = $element.Current.ClassName
                    NativeWindowHandle = $element.Current.NativeWindowHandle
                    DocumentSnippet = $snippet
                    SelectedText = $selectedText
                    Foreground = "$foreground"
                    ForegroundType = $foreground.GetType().FullName
                    ForegroundNotSupported = [object]::ReferenceEquals($foreground,
                        [System.Windows.Automation.AutomationElement]::NotSupported)
                    ForegroundMixed = [object]::ReferenceEquals($foreground,
                        [System.Windows.Automation.TextPattern]::MixedAttributeValue)
                    Background = "$background"
                    BackgroundType = $background.GetType().FullName
                    BackgroundNotSupported = [object]::ReferenceEquals($background,
                        [System.Windows.Automation.AutomationElement]::NotSupported)
                    BackgroundMixed = [object]::ReferenceEquals($background,
                        [System.Windows.Automation.TextPattern]::MixedAttributeValue)
            })
            if ($rows.Count -ge 30) { break }
        }
        catch {
            $rows.Add([pscustomobject]@{ Index = $i; Error = $_.Exception.Message })
        }
    }
    $rows.ToArray() | ConvertTo-Json -Depth 4 -Compress
}

function Read-DebugTree {
    $root = Get-VbeAutomationRoot
    $all = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
    $hits = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $all.Count; $i++) {
        $item = $all.Item($i)
        try {
            if ($item.Current.ProcessId -ne $HostProcessId) { continue }
            $name = $item.Current.Name
            if ($name -notmatch $NamePattern) { continue }
            $pattern = $null
            $canInvoke = $item.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)
            $className = [Text.StringBuilder]::new(128)
            $handle = [IntPtr]$item.Current.NativeWindowHandle
            if ($handle -ne [IntPtr]::Zero) { [void][VBAiInspectionNative]::GetClassName($handle, $className, $className.Capacity) }
            $hits.Add([pscustomobject]@{ Index = $i; Name = $name; Type = $item.Current.ControlType.ProgrammaticName;
                    Handle = $item.Current.NativeWindowHandle; Class = $className.ToString();
                    Enabled = $item.Current.IsEnabled; CanInvoke = $canInvoke })
            if ($Invoke -and $canInvoke -and $item.Current.IsEnabled) {
                ([System.Windows.Automation.InvokePattern]$pattern).Invoke()
                break
            }
            if ($ClickNativeButton -and $className.ToString() -eq 'Button' -and $item.Current.IsEnabled) {
                if (-not [VBAiInspectionNative]::PostMessage($handle, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)) {
                    throw 'BM_CLICK was rejected.'
                }
                break
            }
        }
        catch {
            # An action can fail after being accepted. Do not try another match.
            if ($Invoke -or $ClickNativeButton) { throw }
        }
    }
    $hits.ToArray() | ConvertTo-Json -Depth 5
}

function Read-Locals {
    $root = Get-VbeAutomationRoot
    $condition = [System.Windows.Automation.PropertyCondition]::new(
        [System.Windows.Automation.AutomationElement]::NameProperty, $PaneName)
    $panes = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
    Write-Output "Pane [$PaneName] elements: $($panes.Count)"
    foreach ($pane in $panes) {
        Write-Output "PANE $($pane.Current.ControlType.ProgrammaticName) HWND=$($pane.Current.NativeWindowHandle)"
        $children = $pane.FindAll([System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.Condition]::TrueCondition)
        Write-Output "Descendants: $($children.Count)"
        for ($index = 0; $index -lt [Math]::Min($children.Count, 80); $index++) {
            $element = $children.Item($index)
            try {
                $type = $element.Current.ControlType.ProgrammaticName
                $name = $element.Current.Name
                $id = $element.Current.AutomationId
                $value = ''
                $pattern = $null
                if ($element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) {
                    $value = ([System.Windows.Automation.ValuePattern]$pattern).Current.Value
                }
                $text = ''
                $pattern = $null
                if ($element.TryGetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern, [ref]$pattern)) {
                    $text = ([System.Windows.Automation.TextPattern]$pattern).DocumentRange.GetText(-1)
                }
                Write-Output "ITEM $index $type name=[$name] id=[$id] value=[$value] text=[$text]"
            }
            catch { Write-Output "ITEM $index ERROR $($_.Exception.Message)" }
        }
    }
}

function Read-ObjectBrowser {


    function Find-ChildByName($parent, [string] $name) {
        $condition = [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::NameProperty, $name)
        return $parent.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
    }

    function Select-ListItem($list, $item) {
        if (-not $item) { throw 'The requested Object Browser item was not found.' }
        $selection = $item.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
        $selection.Select()
        $handle = [IntPtr] $list.Current.NativeWindowHandle
        if ($handle -eq [IntPtr]::Zero) { throw 'Object Browser list has no native window handle.' }
        $parent = [VBAiInspectionNative]::GetParent($handle)
        $controlId = [VBAiInspectionNative]::GetDlgCtrlID($handle)
        $notification = [IntPtr] ($controlId -bor (1 -shl 16)) # WM_COMMAND / LBN_SELCHANGE
        if (-not [VBAiInspectionNative]::PostMessage($parent, 0x111, $notification, $handle)) {
            throw 'The Object Browser did not accept the list selection notification.'
        }
    }

    $root = Get-VbeAutomationRoot
    $browserCondition = [System.Windows.Automation.AndCondition]::new(
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Window),
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::NameProperty, $BrowserName))
    $browser = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $browserCondition)
    if (-not $browser) { throw 'The native Object Browser is not open in the selected Excel process.' }

    $listCondition = [System.Windows.Automation.PropertyCondition]::new(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::List)
    $lists = $browser.FindAll([System.Windows.Automation.TreeScope]::Descendants, $listCondition)
    $classes = $null
    $members = $null
    foreach ($list in $lists) {
        $first = $list.FindFirst([System.Windows.Automation.TreeScope]::Children,
            [System.Windows.Automation.Condition]::TrueCondition)
        if (-not $first) { continue }
        if ($first.Current.Name -like 'Classes *') { $classes = $list }
        if ($first.Current.Name -like 'Membres de *') { $members = $list }
    }
    if (-not $classes -or -not $members) { throw 'The native class or member list was not identified.' }

    Select-ListItem $classes (Find-ChildByName $classes "Classes $Class")
    Start-Sleep -Milliseconds 200
    $memberItem = Find-ChildByName $members "Membres de '$Class' $Member"
    Select-ListItem $members $memberItem
    Start-Sleep -Milliseconds 200

    $documentCondition = [System.Windows.Automation.PropertyCondition]::new(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Document)
    $document = $browser.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $documentCondition)
    if (-not $document) { throw 'The Object Browser detail pane was not found.' }
    $textPattern = $document.GetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern)
    $detail = $textPattern.DocumentRange.GetText(-1).Trim()
    [pscustomobject]@{
        HostProcessId = $HostProcessId
        Class = $Class
        Member = $Member
        MemberCount = $members.FindAll([System.Windows.Automation.TreeScope]::Children,
            [System.Windows.Automation.Condition]::TrueCondition).Count
        Detail = $detail
        Source = 'Native VBE Object Browser via UI Automation'
    } | ConvertTo-Json -Compress
}

function Read-Menu {
    if (($OpenMenu -or $InvokeItem) -and -not $AcknowledgeUiAction) { throw 'Explicit acknowledgement is required before opening or invoking a menu in the selected window.' }
    Add-InspectionType 'VBAiMenuMsaa' @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Accessibility;

public static class VBAiMenuMsaa
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
                    if (!string.IsNullOrEmpty(openMenu) && string.Equals(name, openMenu, StringComparison.OrdinalIgnoreCase)) {
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
    [VBAiMenuMsaa]::Inspect([IntPtr]$WindowHandle, $OpenMenu, $InvokeItem)
}

function Read-MsaaChildren {
    Add-InspectionType 'VBAiMsaaProbe' @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class VBAiMsaaProbe {
    [DllImport("oleacc.dll")] public static extern int AccessibleObjectFromWindow(IntPtr handle, uint objectId,
        ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object accessible);
    [DllImport("oleacc.dll")] private static extern int AccessibleChildren(Accessibility.IAccessible container,
        int start, int count, IntPtr variants, out int obtained);
    [DllImport("oleaut32.dll")] private static extern int VariantClear(IntPtr variant);
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
    $matches = @(Get-VbeRows | Where-Object { $_.Class -eq 'VbaWindow' -and $_.Title -eq $PaneName })
    if ($matches.Count -ne 1) { throw 'One exact native pane is required.' }
    $pane = [IntPtr]$matches[0].Handle
    if ($pane -eq [IntPtr]::Zero) { throw "Native pane not found: $PaneName" }
    Write-Output "PANE [$PaneName] HWND=$($pane.ToInt64())"
    [VBAiMsaaProbe]::Read($pane)
}

function Read-References {
    $root=[System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$WindowHandle)
    if(-not $root){throw 'ListBox UI Automation root absent'}
    $items=$root.FindAll([System.Windows.Automation.TreeScope]::Children,[System.Windows.Automation.Condition]::TrueCondition)
    Write-Output "ListChildren=$($items.Count)"
    for($i=0;$i -lt [Math]::Min($items.Count,40);$i++){
        $item=$items.Item($i)
        $toggle=$null
        $hasToggle=$item.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern,[ref]$toggle)
        $state=if($hasToggle){$toggle.Current.ToggleState}else{'n/a'}
        Write-Output "ITEM Type=$($item.Current.ControlType.ProgrammaticName) Name=$($item.Current.Name) Toggle=$state"
    }
    Add-InspectionType 'VBAiReferenceList' @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class VBAiReferenceList {
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr hwnd,uint msg,IntPtr wp,IntPtr lp);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr hwnd,uint msg,IntPtr wp,StringBuilder result);
  public static string[] Read(IntPtr hwnd){int count=SendMessage(hwnd,0x018B,IntPtr.Zero,IntPtr.Zero).ToInt32();if(count<0)throw new InvalidOperationException("List count unavailable");int max=Math.Min(count,30);var rows=new string[max+1];rows[0]="LB_COUNT="+count;for(int i=0;i<max;i++){int length=SendMessage(hwnd,0x018A,new IntPtr(i),IntPtr.Zero).ToInt32();if(length<0||length>4096)throw new InvalidOperationException("List text length unavailable or exceeds 4096");var text=new StringBuilder(Math.Max(1,length+1));SendMessage(hwnd,0x0189,new IntPtr(i),text);rows[i+1]=i+" "+text;}return rows;}
}
'@
    [VBAiReferenceList]::Read([IntPtr]$WindowHandle)
}

# One target per invocation; these are compiled source functions, not script strings.
# Drop a stale cached inventory if the command is invoked again in the same session.
$script:inspectionRows = $null
& ("Read-" + $Target)

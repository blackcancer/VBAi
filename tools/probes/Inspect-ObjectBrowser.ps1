param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [Parameter(Mandatory = $true)] [string] $Class,
    [Parameter(Mandatory = $true)] [string] $Member
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class ObjectBrowserListNotify
{
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr handle);
    [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr handle);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);
}
'@

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
    $parent = [ObjectBrowserListNotify]::GetParent($handle)
    $controlId = [ObjectBrowserListNotify]::GetDlgCtrlID($handle)
    $notification = [IntPtr] ($controlId -bor (1 -shl 16)) # WM_COMMAND / LBN_SELCHANGE
    if (-not [ObjectBrowserListNotify]::PostMessage($parent, 0x111, $notification, $handle)) {
        throw 'The Object Browser did not accept the list selection notification.'
    }
}

$process = Get-Process -Id $HostProcessId
$root = [System.Windows.Automation.AutomationElement]::FromHandle($process.MainWindowHandle)
$browserCondition = [System.Windows.Automation.AndCondition]::new(
    [System.Windows.Automation.PropertyCondition]::new(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Window),
    [System.Windows.Automation.PropertyCondition]::new(
        [System.Windows.Automation.AutomationElement]::NameProperty, "Explorateur d'objets"))
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

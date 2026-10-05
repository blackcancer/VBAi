param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Chat', 'Settings', 'ComboBoxes', 'Menu')]
    [string] $Area,
    [Parameter(Mandatory = $true)] [ValidateRange(1, [int]::MaxValue)] [int] $HostProcessId,
    [long] $WindowHandle
)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
if ($Area -eq 'ComboBoxes' -and -not $WindowHandle) {
    throw 'ComboBoxes requires the exact ChatWindow WindowHandle.'
}
$handle = if ($WindowHandle) { [IntPtr]$WindowHandle } else { $hostProcess.MainWindowHandle }
$root = [System.Windows.Automation.AutomationElement]::FromHandle($handle)
if (-not $root) { throw 'VBE window unavailable.' }
if ($root.Current.ProcessId -ne $HostProcessId) { throw 'Selected window does not belong to HostProcessId.' }

if ($Area -eq 'ComboBoxes') {
    if ($root.Current.AutomationId -cne 'ChatWindow' -or
        $root.Current.ControlType -ne [System.Windows.Automation.ControlType]::Window) {
        throw 'ComboBoxes requires the exact owned ChatWindow handle.'
    }
    # UIA verifies ownership; the original combo projection uses read-only Win32 messages.
    if (-not ('VBAi.Tools.VbeComboInspector' -as [type])) {
        Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi.Tools
{
    public static class VbeComboInspector
    {
        private delegate bool EnumProc(IntPtr hwnd, IntPtr data);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr data);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder text, int length);
        [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern IntPtr SendMessage(IntPtr hwnd, int message, IntPtr param, StringBuilder output);

        public static string[] Read(IntPtr assistant, int targetPid)
        {
            const int maxItems = 100;
            const int maxTextLength = 4096;
            uint owner;
            GetWindowThreadProcessId(assistant, out owner);
            if (owner != (uint)targetPid) throw new InvalidOperationException("Assistant HWND belongs to a different process.");
            var result = new List<string>();
            string error = null;
            EnumChildWindows(assistant, (hwnd, data) => {
                uint pid;
                GetWindowThreadProcessId(hwnd, out pid);
                if (pid != (uint)targetPid) return true;
                var kind = new StringBuilder(200);
                GetClassName(hwnd, kind, kind.Capacity);
                if (kind.ToString().IndexOf("COMBOBOX", StringComparison.OrdinalIgnoreCase) >= 0) {
                    int count = SendMessage(hwnd, 0x146, IntPtr.Zero, null).ToInt32();
                    if (count < 0) { error = "Combo item count unavailable."; return false; }
                    result.Add("COMBO HWND=" + hwnd + " Count=" + count);
                    for (int index = 0; index < count && index < maxItems; index++) {
                        int length = SendMessage(hwnd, 0x149, new IntPtr(index), null).ToInt32();
                        if (length < 0 || length > maxTextLength) {
                            error = "Combo item text length is unavailable or exceeds 4096 characters.";
                            return false;
                        }
                        var label = new StringBuilder(Math.Max(1, length + 1));
                        SendMessage(hwnd, 0x148, new IntPtr(index), label);
                        result.Add("  " + index + " " + label);
                    }
                }
                return true;
            }, IntPtr.Zero);
            if (error != null) throw new InvalidOperationException(error);
            return result.ToArray();
        }
    }
}
'@
    }
    [VBAi.Tools.VbeComboInspector]::Read($handle, $HostProcessId)
    return
}

if ($Area -eq 'Menu') {
    Write-Output "Root=$($root.Current.Name)"
    $items = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($item in $items) {
        if ($item.Current.ProcessId -ne $HostProcessId) { continue }
        $name = $item.Current.Name
        if ($name -match 'Affichage|Outils|Codex|VBAi|View|Tools|Barre de menus') {
            $pattern = $null
            $expandable = $item.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$pattern)
            $invokable = $item.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)
            Write-Output "ITEM $name Type=$($item.Current.ControlType.ProgrammaticName) HWnd=$($item.Current.NativeWindowHandle) Offscreen=$($item.Current.IsOffscreen) Expandable=$expandable Invokable=$invokable"
        }
    }
    return
}

$automationId = if ($Area -eq 'Chat') { 'ChatWindow' } else { 'LlmSettingsWindow' }
if ($WindowHandle) {
    if ($root.Current.AutomationId -cne $automationId -or
        $root.Current.ControlType -ne [System.Windows.Automation.ControlType]::Window) {
        throw "Explicit WindowHandle is not the exact $automationId window."
    }
    $selected = $root
} else {
    $condition = [System.Windows.Automation.AndCondition]::new(
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Window),
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::AutomationIdProperty, $automationId))
    $windows = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
    $matching = @()
    foreach ($window in $windows) {
        if ($window.Current.ProcessId -eq $HostProcessId) { $matching += $window }
    }
    if ($matching.Count -ne 1) { throw "Expected one owned $automationId descendant; found $($matching.Count)." }
    $selected = $matching[0]
}
if ($Area -eq 'Chat') { Write-Output "Assistant=$($selected.Current.Name)" }
else { Write-Output "Settings=$($selected.Current.Name)" }
$children = $selected.FindAll([System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition)
$limit = if ($Area -eq 'Chat') { 150 } else { 160 }
foreach ($child in $children) {
    if ($child.Current.ProcessId -ne $HostProcessId) { continue }
    $name = $child.Current.Name.Replace("`r", ' ').Replace("`n", ' ')
    if ($name.Length -gt $limit) { $name = $name.Substring(0, $limit) }
    Write-Output "UI $($child.Current.ControlType.ProgrammaticName) Enabled=$($child.Current.IsEnabled) Name=$name"
}

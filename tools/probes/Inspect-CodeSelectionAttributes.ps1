param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class CodexCodeSelectionWindows {
    public delegate bool EnumProc(IntPtr handle, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc callback, IntPtr data);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr handle, StringBuilder text, int capacity);
}
'@

$script:rootHandle = [IntPtr]::Zero
$find = [CodexCodeSelectionWindows+EnumProc]{
    param($handle, $data)
    [uint32]$windowProcessId = 0
    [void][CodexCodeSelectionWindows]::GetWindowThreadProcessId($handle, [ref]$windowProcessId)
    if ($windowProcessId -ne $HostProcessId) { return $true }
    $class = [Text.StringBuilder]::new(128)
    [void][CodexCodeSelectionWindows]::GetClassName($handle, $class, $class.Capacity)
    if ($class.ToString() -eq 'wndclass_desked_gsk') {
        $script:rootHandle = $handle
        return $false
    }
    return $true
}
[void][CodexCodeSelectionWindows]::EnumWindows($find, [IntPtr]::Zero)
if ($script:rootHandle -eq [IntPtr]::Zero) { throw 'Visible VBE window not found.' }

$root = [Windows.Automation.AutomationElement]::FromHandle($script:rootHandle)
$all = $root.FindAll([Windows.Automation.TreeScope]::Descendants,
    [Windows.Automation.Condition]::TrueCondition)
$rows = [Collections.Generic.List[object]]::new()
for ($i = 0; $i -lt $all.Count; $i++) {
    $element = $all.Item($i)
    try {
        $patternObject = $null
        if (-not $element.TryGetCurrentPattern([Windows.Automation.TextPattern]::Pattern,
            [ref]$patternObject)) { continue }
        $pattern = [Windows.Automation.TextPattern]$patternObject
        $selection = @($pattern.GetSelection())
        $selectedText = if ($selection.Count -gt 0) { $selection[0].GetText(300) } else { '' }
        $range = if ($selection.Count -gt 0) { $selection[0] } else { $pattern.DocumentRange }
        $foreground = $range.GetAttributeValue([Windows.Automation.TextPattern]::ForegroundColorAttribute)
        $background = $range.GetAttributeValue([Windows.Automation.TextPattern]::BackgroundColorAttribute)
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
                [Windows.Automation.AutomationElement]::NotSupported)
            ForegroundMixed = [object]::ReferenceEquals($foreground,
                [Windows.Automation.TextPattern]::MixedAttributeValue)
            Background = "$background"
            BackgroundType = $background.GetType().FullName
            BackgroundNotSupported = [object]::ReferenceEquals($background,
                [Windows.Automation.AutomationElement]::NotSupported)
            BackgroundMixed = [object]::ReferenceEquals($background,
                [Windows.Automation.TextPattern]::MixedAttributeValue)
        })
        if ($rows.Count -ge 30) { break }
    }
    catch {
        $rows.Add([pscustomobject]@{ Index = $i; Error = $_.Exception.Message })
    }
}
$rows.ToArray() | ConvertTo-Json -Depth 4 -Compress

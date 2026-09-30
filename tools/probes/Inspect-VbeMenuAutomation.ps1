param([Parameter(Mandatory = $true)] [int] $HostProcessId, [long] $WindowHandle)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$hostProcess = Get-Process -Id $HostProcessId
$handle = if ($WindowHandle) { [IntPtr]$WindowHandle } else { $hostProcess.MainWindowHandle }
$root = [System.Windows.Automation.AutomationElement]::FromHandle($handle)
if (-not $root) { throw 'VBE main window unavailable.' }
Write-Output "Root=$($root.Current.Name)"
$items = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition)
foreach ($item in $items) {
    $name = $item.Current.Name
    if ($name -match 'Affichage|Outils|Codex|View|Tools|Barre de menus') {
        $pattern = $null
        $expandable = $item.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$pattern)
        $invokable = $item.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)
        Write-Output "ITEM $name Type=$($item.Current.ControlType.ProgrammaticName) HWnd=$($item.Current.NativeWindowHandle) Offscreen=$($item.Current.IsOffscreen) Expandable=$expandable Invokable=$invokable"
    }
}

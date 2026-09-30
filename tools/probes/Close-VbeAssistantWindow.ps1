param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$hostProcess = Get-Process -Id $HostProcessId
$root = [System.Windows.Automation.AutomationElement]::FromHandle($hostProcess.MainWindowHandle)
$windows = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.PropertyCondition]::new(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Window))
$assistant = $null
foreach ($window in $windows) {
    if ($window.Current.Name -match '^CodexVBE.*Assistant$') { $assistant = $window; break }
}
if (-not $assistant) { Write-Output 'Assistant window is not open.'; return }
$pattern = $null
if (-not $assistant.TryGetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern, [ref]$pattern)) {
    throw 'Assistant window does not expose WindowPattern.'
}
$pattern.Close()
Write-Output 'Closed assistant window through UI Automation WindowPattern.'

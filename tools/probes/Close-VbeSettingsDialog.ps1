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
$dialog = $null
foreach ($window in $windows) {
    if ($window.Current.Name -match '^CodexVBE.*Configuration LLM$') { $dialog = $window; break }
}
if (-not $dialog) { Write-Output 'Settings dialog is not open.'; return }
$pattern = $null
if (-not $dialog.TryGetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern, [ref]$pattern)) {
    throw 'Settings dialog does not expose WindowPattern.'
}
$pattern.Close()
Write-Output 'Closed settings dialog through UI Automation WindowPattern.'

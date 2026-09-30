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
$settings = $null
foreach ($window in $windows) {
    if ($window.Current.Name -match '^CodexVBE.*Configuration LLM$') { $settings = $window; break }
}
if (-not $settings) { throw 'Settings dialog not found in the VBE accessibility tree.' }
Write-Output "Settings=$($settings.Current.Name)"
$children = $settings.FindAll([System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition)
foreach ($child in $children) {
    $name = $child.Current.Name.Replace("`r", ' ').Replace("`n", ' ')
    if ($name.Length -gt 160) { $name = $name.Substring(0, 160) }
    Write-Output "UI $($child.Current.ControlType.ProgrammaticName) Enabled=$($child.Current.IsEnabled) Name=$name"
}

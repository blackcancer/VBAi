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
$chat = $null
foreach ($window in $windows) {
    if ($window.Current.Name -match '^CodexVBE.*Assistant$') { $chat = $window; break }
}
if (-not $chat) { throw 'Assistant not found in the VBE accessibility tree.' }
Write-Output "Assistant=$($chat.Current.Name)"
$children = $chat.FindAll([System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition)
foreach ($child in $children) {
    $name = $child.Current.Name.Replace("`r", ' ').Replace("`n", ' ')
    if ($name.Length -gt 150) { $name = $name.Substring(0, 150) }
    Write-Output "UI $($child.Current.ControlType.ProgrammaticName) Enabled=$($child.Current.IsEnabled) Name=$name"
}

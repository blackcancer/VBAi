param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$root = [Windows.Automation.AutomationElement]::RootElement
$processCondition = [Windows.Automation.PropertyCondition]::new(
    [Windows.Automation.AutomationElement]::ProcessIdProperty, $HostProcessId)
$windows = $root.FindAll([Windows.Automation.TreeScope]::Children, $processCondition)
$vbe = $null
foreach ($window in $windows) {
    if ($window.Current.Name -like 'Microsoft Visual Basic*') { $vbe = $window; break }
}
if ($null -eq $vbe) { throw 'Visible VBE window not found.' }

$code = $vbe.FindFirst([Windows.Automation.TreeScope]::Descendants,
    [Windows.Automation.AndCondition]::new(
        [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::ControlTypeProperty,
            [Windows.Automation.ControlType]::Document),
        [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::NameProperty, 'ThisWorkbook (Code)')))
if ($null -eq $code) { throw 'ThisWorkbook code document not found.' }
$walker = [Windows.Automation.TreeWalker]::RawViewWalker
$container = $walker.GetParent($code)
if ($null -eq $container) { throw 'Code document parent unavailable.' }
$items = $container.FindAll([Windows.Automation.TreeScope]::Descendants,
    [Windows.Automation.Condition]::TrueCondition)
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

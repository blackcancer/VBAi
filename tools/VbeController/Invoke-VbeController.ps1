param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [Parameter(Mandatory = $true)]
    [ValidateSet('state', 'windows', 'find_text', 'list_symbols', 'locals', 'immediate')]
    [string] $Command,
    [string] $Project,
    [string] $Query,
    [switch] $CaseSensitive
)

$ErrorActionPreference = 'Stop'

Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class VbeWindowApi
{
    public delegate bool EnumWindowsProc(IntPtr handle, IntPtr parameter);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr parameter);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr handle, StringBuilder text, int capacity);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr handle, StringBuilder text, int capacity);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr handle);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
}
'@

function Invoke-Bridge {
    param([hashtable] $Request)
    $json = ConvertTo-Json -InputObject $Request -Compress
    $pipe = [System.IO.Pipes.NamedPipeClientStream]::new('.', "CodexVBE.$HostProcessId", [System.IO.Pipes.PipeDirection]::InOut)
    try {
        $pipe.Connect(3000)
        $writer = [System.IO.StreamWriter]::new($pipe, [System.Text.UTF8Encoding]::new($false), 4096, $true)
        $reader = [System.IO.StreamReader]::new($pipe, [System.Text.UTF8Encoding]::new($false), $false, 4096, $true)
        try {
            $writer.AutoFlush = $true
            $writer.WriteLine($json)
            $line = $reader.ReadLine()
            if (-not $line) { throw 'The VBE bridge returned no response.' }
            $response = ConvertFrom-Json -InputObject $line
            if (-not $response.Ok) { throw [string]$response.Error }
            return $response.Data
        }
        finally { $reader.Dispose(); $writer.Dispose() }
    }
    finally { $pipe.Dispose() }
}

function Get-WindowInfo {
    param([IntPtr] $Handle)
    $title = [System.Text.StringBuilder]::new(512)
    $className = [System.Text.StringBuilder]::new(256)
    [void][VbeWindowApi]::GetWindowText($Handle, $title, $title.Capacity)
    [void][VbeWindowApi]::GetClassName($Handle, $className, $className.Capacity)
    return [pscustomobject]@{
        Handle = $Handle.ToInt64()
        Class = $className.ToString()
        Title = $title.ToString()
        Visible = [VbeWindowApi]::IsWindowVisible($Handle)
    }
}

function Get-VbeWindows {
    $script:rootHandle = [IntPtr]::Zero
    $callback = [VbeWindowApi+EnumWindowsProc]{
        param($handle, $parameter)
        [uint32]$ownerId = 0
        [void][VbeWindowApi]::GetWindowThreadProcessId($handle, [ref]$ownerId)
        if ($ownerId -ne $HostProcessId) { return $true }
        $info = Get-WindowInfo $handle
        if ($info.Class -eq 'wndclass_desked_gsk') {
            $script:rootHandle = $handle
            return $false
        }
        return $true
    }
    [void][VbeWindowApi]::EnumWindows($callback, [IntPtr]::Zero)
    if ($script:rootHandle -eq [IntPtr]::Zero) { throw "VBE window not found for host PID $HostProcessId." }
    $children = [System.Collections.Generic.List[object]]::new()
    $childCallback = [VbeWindowApi+EnumWindowsProc]{
        param($handle, $parameter)
        $info = Get-WindowInfo $handle
        if ($info.Class -eq 'VbaWindow') { $children.Add($info) }
        return $true
    }
    [void][VbeWindowApi]::EnumChildWindows($script:rootHandle, $childCallback, [IntPtr]::Zero)
    return [pscustomobject]@{
        Root = Get-WindowInfo $script:rootHandle
        Foreground = ([VbeWindowApi]::GetForegroundWindow() -eq $script:rootHandle)
        Windows = @($children.ToArray())
    }
}

function Get-VbeAutomationRoot {
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    $window = Get-VbeWindows
    return [Windows.Automation.AutomationElement]::FromHandle([IntPtr]$window.Root.Handle)
}

function Find-VbeNamedElements {
    param([string[]] $Names)
    $root = Get-VbeAutomationRoot
    $result = [System.Collections.Generic.List[object]]::new()
    foreach ($name in $Names) {
        $condition = [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::NameProperty, $name)
        $found = $root.FindAll([Windows.Automation.TreeScope]::Descendants, $condition)
        for ($index = 0; $index -lt $found.Count; $index++) { $result.Add($found.Item($index)) }
    }
    return @($result.ToArray())
}

function Get-VbeLocals {
    $lists = @(Find-VbeNamedElements @('Variables locales', 'Locals') |
        Where-Object { $_.Current.ControlType -eq [Windows.Automation.ControlType]::List })
    if ($lists.Count -ne 1) { throw "Expected one VBE Locals list; found $($lists.Count)." }
    $condition = [Windows.Automation.PropertyCondition]::new(
        [Windows.Automation.AutomationElement]::ControlTypeProperty,
        [Windows.Automation.ControlType]::ListItem)
    $elements = $lists[0].FindAll([Windows.Automation.TreeScope]::Descendants, $condition)
    $items = [System.Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt $elements.Count; $index++) {
        $element = $elements.Item($index)
        $pattern = $null
        $value = $null
        if ($element.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) {
            $value = ([Windows.Automation.ValuePattern]$pattern).Current.Value
        }
        $items.Add([pscustomobject]@{ Name = $element.Current.Name; Value = $value })
    }
    return [pscustomobject]@{ Items = @($items.ToArray()) }
}

function Get-VbeImmediate {
    $window = Get-VbeWindows
    $panes = @($window.Windows | Where-Object { $_.Title -like 'Ex*cution' -or $_.Title -eq 'Immediate' })
    if ($panes.Count -ne 1) { throw "Expected one VBE Immediate window; found $($panes.Count)." }
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    $root = [Windows.Automation.AutomationElement]::FromHandle([IntPtr]$panes[0].Handle)
    $condition = [Windows.Automation.PropertyCondition]::new(
        [Windows.Automation.AutomationElement]::ControlTypeProperty,
        [Windows.Automation.ControlType]::Document)
    $found = $root.FindAll([Windows.Automation.TreeScope]::Descendants, $condition)
    $documents = @()
    for ($index = 0; $index -lt $found.Count; $index++) { $documents += $found.Item($index) }
    if ($documents.Count -ne 1) { throw "Expected one VBE Immediate document; found $($documents.Count)." }
    $pattern = $null
    if (-not $documents[0].TryGetCurrentPattern([Windows.Automation.TextPattern]::Pattern, [ref]$pattern)) {
        throw 'The VBE Immediate document does not expose UI Automation TextPattern.'
    }
    return [pscustomobject]@{ Text = ([Windows.Automation.TextPattern]$pattern).DocumentRange.GetText(-1) }
}

function Get-ProjectModules {
    param([string] $Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { throw 'Project is required.' }
    $projects = @(Invoke-Bridge @{ Command = 'list_projects' })
    $matches = @($projects | Where-Object { $_.Name -eq $Name })
    if ($matches.Count -ne 1) { throw "Project name is absent or ambiguous: $Name" }
    $modules = @(Invoke-Bridge @{ Command = 'list_modules'; Project = $Name })
    return $modules
}

try {
    if (-not (Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)) {
        throw "Host PID $HostProcessId was not found."
    }
    switch ($Command) {
        'windows' {
            $data = Get-VbeWindows
        }
        'state' {
            $data = [pscustomobject]@{
                Bridge = Invoke-Bridge @{ Command = 'status' }
                Projects = @(Invoke-Bridge @{ Command = 'list_projects' })
                Vbe = Get-VbeWindows
            }
        }
        'find_text' {
            if ([string]::IsNullOrEmpty($Query)) { throw 'Query is required.' }
            $comparison = if ($CaseSensitive) { [StringComparison]::Ordinal } else { [StringComparison]::OrdinalIgnoreCase }
            $results = [System.Collections.Generic.List[object]]::new()
            foreach ($module in (Get-ProjectModules $Project)) {
                $source = Invoke-Bridge @{ Command = 'read_module'; Project = $Project; Module = $module.Name }
                $lines = [regex]::Split([string]$source.Code, '\r\n|\n|\r')
                for ($i = 0; $i -lt $lines.Length; $i++) {
                    $offset = 0
                    while ($offset -lt $lines[$i].Length) {
                        $column = $lines[$i].IndexOf($Query, $offset, $comparison)
                        if ($column -lt 0) { break }
                        $results.Add([pscustomobject]@{ Project = $Project; Module = $module.Name; Line = $i + 1; Column = $column + 1; Text = $lines[$i] })
                        $offset = $column + [Math]::Max(1, $Query.Length)
                    }
                }
            }
            $data = @($results.ToArray())
        }
        'list_symbols' {
            $results = [System.Collections.Generic.List[object]]::new()
            $pattern = '^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?:(?:Declare\s+(?:PtrSafe\s+)?)?(?:Sub|Function)|Property\s+(?:Get|Let|Set)|Type|Enum)\s+([A-Za-z_][A-Za-z0-9_]*)\b'
            foreach ($module in (Get-ProjectModules $Project)) {
                $source = Invoke-Bridge @{ Command = 'read_module'; Project = $Project; Module = $module.Name }
                $lines = [regex]::Split([string]$source.Code, '\r\n|\n|\r')
                for ($i = 0; $i -lt $lines.Length; $i++) {
                    $match = [regex]::Match($lines[$i], $pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    if ($match.Success) {
                        $results.Add([pscustomobject]@{ Project = $Project; Module = $module.Name; Line = $i + 1; Name = $match.Groups[1].Value; Declaration = $lines[$i].Trim() })
                    }
                }
            }
            $data = @($results.ToArray())
        }
        'locals' {
            $data = Get-VbeLocals
        }
        'immediate' {
            $data = Get-VbeImmediate
        }
    }
    [pscustomobject]@{ Ok = $true; Error = $null; Data = $data } | ConvertTo-Json -Depth 10 -Compress
}
catch {
    [pscustomobject]@{ Ok = $false; Error = $_.Exception.Message; Data = $null } | ConvertTo-Json -Depth 5 -Compress
    exit 1
}

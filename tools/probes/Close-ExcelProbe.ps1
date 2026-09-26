param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$hosts = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($hosts.Count -ne 1 -or $hosts[0].Id -ne $HostProcessId) {
    throw 'The test Excel process is no longer the sole Excel instance.'
}

Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class VbeVisibility {
    public delegate bool EnumWindowsProc(IntPtr handle, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr data);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr handle, System.Text.StringBuilder text, int capacity);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr handle, int command);
}
'@

$excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
if ($excel.ActiveWorkbook.Name -ne 'Classeur1') { throw 'The active workbook is not the disposable Classeur1.' }
$script:vbeWindow = [IntPtr]::Zero
$callback = [VbeVisibility+EnumWindowsProc]{
    param($handle, $data)
    [uint32]$owner = 0
    [VbeVisibility]::GetWindowThreadProcessId($handle, [ref]$owner) | Out-Null
    if ($owner -eq $HostProcessId) {
        $class = [Text.StringBuilder]::new(100)
        [VbeVisibility]::GetClassName($handle, $class, $class.Capacity) | Out-Null
        if ($class.ToString() -eq 'wndclass_desked_gsk') { $script:vbeWindow = $handle }
    }
    return $true
}
[VbeVisibility]::EnumWindows($callback, [IntPtr]::Zero) | Out-Null
if ($script:vbeWindow -ne [IntPtr]::Zero) { [VbeVisibility]::ShowWindow($script:vbeWindow, 0) | Out-Null }
$excel.ActiveWorkbook.Close($false)
$excel.Quit()
Write-Output "Closed disposable Excel PID=$HostProcessId without saving."

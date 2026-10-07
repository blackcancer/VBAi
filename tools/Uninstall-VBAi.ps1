param([switch] $Direct, [string]$AssemblyDirectory, [switch]$PreserveMachineMapping)

$ErrorActionPreference = 'Stop'

if ($env:CODEX_SHELL -eq '1' -and -not $Direct -and -not $AssemblyDirectory) {
    & (Join-Path $PSScriptRoot 'Invoke-VBAi-OutsideSandbox.ps1') -Action Uninstall
    return
}

if (-not [Environment]::Is64BitProcess) {
    throw 'Run this script from 64-bit PowerShell.'
}

$classId = '{8E854243-087F-4D6C-9E0E-8622B0E50883}'
$progId = 'VBAi.AddIn'
$registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
    [Microsoft.Win32.RegistryHive]::CurrentUser,
    [Microsoft.Win32.RegistryView]::Registry64)

try {
    $progKey = $registry.OpenSubKey("Software\Classes\$progId\CLSID")
    if ($progKey) {
        try {
            if ($progKey.GetValue('') -ne $classId) { throw "The ProgID $progId belongs to another COM class." }
        }
        finally { $progKey.Dispose() }
    }
    $classKey = $registry.OpenSubKey("Software\Classes\CLSID\$classId\InprocServer32")
    if ($classKey) {
        try {
            $assembly = [string] $classKey.GetValue('Assembly')
            if ($assembly -and -not $assembly.StartsWith('VBAi,', [StringComparison]::Ordinal)) {
                throw "The CLSID $classId belongs to another assembly."
            }
        }
        finally { $classKey.Dispose() }
    }

    & (Join-Path $PSScriptRoot 'Register-VbaTestRuntime.ps1') -Unregister
    $registry.DeleteSubKeyTree("Software\Microsoft\VBA\VBE\6.0\Addins64\$progId", $false)
    if (-not $AssemblyDirectory) { $AssemblyDirectory = Join-Path (Split-Path -Parent $PSScriptRoot) 'bin\Debug\net48' }
    $typeLibPath = Join-Path $AssemblyDirectory 'VBAi.tlb'
    if (Test-Path -LiteralPath $typeLibPath) {
        & (Join-Path $PSScriptRoot 'Register-VBAiTypeLib.ps1') -Unregister -TypeLibPath $typeLibPath
    }
    & (Join-Path $PSScriptRoot 'Register-ChatToolWindow.ps1') -Unregister -PreserveMachineMapping:$PreserveMachineMapping
    $registry.DeleteSubKeyTree("Software\Classes\$progId", $false)
    $registry.DeleteSubKeyTree("Software\Classes\CLSID\$classId", $false)
    Write-Output "Unregistered $progId for the current user."
    Write-Output 'Restart each VBA host to unload the add-in.'
}
finally { $registry.Dispose() }

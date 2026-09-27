param([switch]$Unregister)
$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit PowerShell.' }
$chatClassId = '{0F4D723B-97D8-42E5-9B31-70646B97C8D2}'
$chatProgId = 'CodexVBE.ChatToolWindow'
$chatAssemblyPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'bin\Debug\net48\CodexVBE.dll'
$registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Registry64)
try {
    $existing = $registry.OpenSubKey("Software\Classes\$chatProgId\CLSID")
    if ($existing) {
        try { if ($existing.GetValue('') -ne $chatClassId) { throw 'Chat ProgID belongs to another class.' } }
        finally { $existing.Dispose() }
    }
    $existing = $registry.OpenSubKey("Software\Classes\CLSID\$chatClassId\InprocServer32")
    if ($existing) {
        try { if (-not ([string]$existing.GetValue('Assembly')).StartsWith('CodexVBE,', [StringComparison]::Ordinal)) { throw 'Chat CLSID belongs to another assembly.' } }
        finally { $existing.Dispose() }
    }
    if ($Unregister) {
        $registry.DeleteSubKeyTree("Software\Classes\$chatProgId", $false)
        $registry.DeleteSubKeyTree("Software\Classes\CLSID\$chatClassId", $false)
        return
    }
    $chatAssemblyPath = (Resolve-Path -LiteralPath $chatAssemblyPath).Path
    $chatAssemblyName = [Reflection.AssemblyName]::GetAssemblyName($chatAssemblyPath).FullName
    if (-not $chatAssemblyName.StartsWith('CodexVBE, Version=0.1.0.0,', [StringComparison]::Ordinal)) { throw 'Unexpected assembly identity.' }
    $entries = @{
        "Software\Classes\$chatProgId\CLSID" = @{ '' = $chatClassId }
        "Software\Classes\CLSID\$chatClassId" = @{ '' = $chatProgId }
        "Software\Classes\CLSID\$chatClassId\ProgId" = @{ '' = $chatProgId }
        "Software\Classes\CLSID\$chatClassId\InprocServer32" = @{
            '' = 'mscoree.dll'; ThreadingModel = 'Both'; Class = $chatProgId
            Assembly = $chatAssemblyName; RuntimeVersion = 'v4.0.30319'
            CodeBase = 'file:///' + $chatAssemblyPath.Replace([char]92, [char]47)
        }
        "Software\Classes\CLSID\$chatClassId\Control" = @{ '' = '' }
        "Software\Classes\CLSID\$chatClassId\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}" = @{ '' = '' }
    }
    foreach ($path in $entries.Keys) {
        $key = $registry.CreateSubKey($path)
        try { foreach ($name in $entries[$path].Keys) { $key.SetValue($name, $entries[$path][$name], [Microsoft.Win32.RegistryValueKind]::String) } }
        finally { $key.Dispose() }
    }
    Write-Output 'Registered CodexVBE native chat control for the current user.'
}
finally { $registry.Dispose() }

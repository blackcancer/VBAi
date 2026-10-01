#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$AssemblyPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'bin\Debug\net48\VBAi.dll'),
    [switch]$Unregister,
    [switch]$Preview,
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit PowerShell.' }
$runtimeId = '{5AF2F40B-939B-4CC6-A06C-F0C79841C031}'
$runtimeProgId = 'VBAi.TestRuntime'
$runtimeClass = 'VBAi.VbaTestRuntime'
$progRoot = "Software\Classes\$runtimeProgId"
$classRoot = "Software\Classes\CLSID\$runtimeId"
$registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Registry64)

function Read-Value([string]$Path, [string]$Name) {
    $key = $registry.OpenSubKey($Path)
    if ($null -eq $key) { return $null }
    try { return $key.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) }
    finally { $key.Dispose() }
}

function Read-Tree([string]$Path) {
    $key = $registry.OpenSubKey($Path)
    if ($null -eq $key) { return [pscustomobject]@{ Path = $Path; Exists = $false; Values = @(); Children = @() } }
    try {
        $values = @($key.GetValueNames() | ForEach-Object {
            [pscustomobject]@{ Name = $_; Kind = $key.GetValueKind($_).ToString(); Value = $key.GetValue($_, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) }
        })
        $children = @($key.GetSubKeyNames() | ForEach-Object { Read-Tree ($Path + '\' + $_) })
        return [pscustomobject]@{ Path = $Path; Exists = $true; Values = $values; Children = $children }
    } finally { $key.Dispose() }
}

try {
    $before = @((Read-Tree $progRoot), (Read-Tree $classRoot))
    # Existing incomplete or foreign keys are not evidence of ownership.
    if ($before[0].Exists -and (Read-Value ($progRoot + '\CLSID') '') -ine $runtimeId) { throw 'Runtime ProgID ownership collision; no registry change was made.' }
    if ($before[1].Exists) {
        $owner = Read-Value ($classRoot + '\InprocServer32') 'Class'
        $identity = [string](Read-Value ($classRoot + '\InprocServer32') 'Assembly')
        if ($owner -cne $runtimeClass -or -not $identity.StartsWith('VBAi,', [StringComparison]::Ordinal)) { throw 'Runtime CLSID ownership collision; no registry change was made.' }
        $declaredProgId = Read-Value ($classRoot + '\ProgId') ''
        if ($null -ne $declaredProgId -and $declaredProgId -cne $runtimeProgId) { throw 'Runtime CLSID ProgID ownership collision; no registry change was made.' }
    }
    $entries = @{}
    if (-not $Unregister) {
        $resolvedAssembly = (Resolve-Path -LiteralPath $AssemblyPath).Path
        $assemblyName = [Reflection.AssemblyName]::GetAssemblyName($resolvedAssembly)
        if ($assemblyName.Name -cne 'VBAi') { throw 'Expected the VBAi assembly; no registry change was made.' }
        $codeBase = 'file:///' + $resolvedAssembly.Replace([char]92, [char]47)
        $server = @{ '' = 'mscoree.dll'; Class = $runtimeClass; Assembly = $assemblyName.FullName; RuntimeVersion = 'v4.0.30319'; ThreadingModel = 'Both'; CodeBase = $codeBase }
        $entries[$progRoot] = @{ '' = $runtimeProgId }
        $entries[$progRoot + '\CLSID'] = @{ '' = $runtimeId }
        $entries[$classRoot] = @{ '' = $runtimeProgId }
        $entries[$classRoot + '\ProgId'] = @{ '' = $runtimeProgId }
        $entries[$classRoot + '\InprocServer32'] = $server
        $entries[$classRoot + '\InprocServer32\' + $assemblyName.Version.ToString()] = $server
    }
    $operation = if ($Unregister) { 'Unregister' } else { 'Register' }
    $plan = [pscustomobject]@{ Hive = 'HKCU'; View = 'Registry64'; Operation = $operation; ProgId = $runtimeProgId; Class = $runtimeClass; Clsid = $runtimeId; Before = $before; Proposed = $entries }
    if ($Preview) { $plan; return } # No backup folder, file, registry write or COM activation.
    if ($Unregister -and -not $before[0].Exists -and -not $before[1].Exists) { Write-Output 'VBAi.TestRuntime is already absent from HKCU x64.'; return }
    if ($ReportPath) { $backup = [IO.Path]::GetFullPath($ReportPath) }
    else { $backup = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) ('VBAi\RegistrationBackups\' + [Guid]::NewGuid().ToString('N') + '\test-runtime-before.clixml') }
    if (Test-Path -LiteralPath $backup) { throw 'The registry backup path already exists; it will not be overwritten.' }
    [IO.Directory]::CreateDirectory((Split-Path -Parent $backup)) | Out-Null
    # CLIXML preserves default names, kinds, unexpanded strings, binary and multi-string data.
    $plan | Export-Clixml -LiteralPath $backup -Depth 100
    if (-not (Test-Path -LiteralPath $backup)) { throw 'Registry backup was not created; mutation refused.' }
    try {
        if ($Unregister) {
            $registry.DeleteSubKeyTree($progRoot, $false)
            $registry.DeleteSubKeyTree($classRoot, $false)
        } else {
            foreach ($path in $entries.Keys) {
                $key = $registry.CreateSubKey($path)
                try { foreach ($name in $entries[$path].Keys) { $key.SetValue($name, $entries[$path][$name], [Microsoft.Win32.RegistryValueKind]::String) } }
                finally { $key.Dispose() }
            }
        }
    } catch { throw "Runtime registration may be partially changed. Exact prior values preserved at $backup. No automatic retry or restoration: $($_.Exception.Message)" }
    [pscustomobject]@{ Operation = $operation; Hive = 'HKCU'; View = 'Registry64'; ProgId = $runtimeProgId; Backup = $backup; ComActivation = 'NOT_TESTED' }
} finally { $registry.Dispose() }

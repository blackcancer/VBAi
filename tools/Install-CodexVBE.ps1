param([switch] $Direct)

$ErrorActionPreference = 'Stop'

if ($env:CODEX_SHELL -eq '1' -and -not $Direct) {
    $expectedAssembly = Join-Path (Split-Path -Parent $PSScriptRoot) 'bin\Debug\net48\CodexVBE.dll'
    & (Join-Path $PSScriptRoot 'Invoke-CodexVBE-OutsideSandbox.ps1') -Action Install -ExpectedAssemblyPath $expectedAssembly
    return
}

if (-not [Environment]::Is64BitProcess) {
    throw 'Run this script from 64-bit PowerShell.'
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$assemblyPath = Join-Path $projectRoot 'bin\Debug\net48\CodexVBE.dll'
$typeLibPath = Join-Path $projectRoot 'bin\Debug\net48\CodexVBE.tlb'
if (-not (Test-Path -LiteralPath $assemblyPath)) {
    throw "Build src/CodexVBE/CodexVBE.csproj first. Missing: $assemblyPath"
}
foreach ($dependency in @('Markdig.dll', 'System.Resources.Extensions.dll', 'System.Memory.dll', 'System.Buffers.dll', 'System.Numerics.Vectors.dll', 'System.Runtime.CompilerServices.Unsafe.dll')) {
    if (-not (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $assemblyPath) $dependency))) {
        throw "Missing chat rendering dependency: $dependency. Rebuild and keep the complete output directory."
    }
}
if (-not (Test-Path -LiteralPath $typeLibPath)) {
    throw "Generate the COM type library first. Missing: $typeLibPath"
}
if ((Get-Item -LiteralPath $typeLibPath).LastWriteTimeUtc -lt (Get-Item -LiteralPath $assemblyPath).LastWriteTimeUtc) {
    throw 'The COM type library is older than CodexVBE.dll. Regenerate it with TlbExp.exe before installing.'
}
$assemblyPath = (Resolve-Path -LiteralPath $assemblyPath).Path
$assemblyName = [Reflection.AssemblyName]::GetAssemblyName($assemblyPath).FullName
if (-not $assemblyName.StartsWith('CodexVBE, Version=0.1.0.0,', [StringComparison]::Ordinal)) {
    throw "Unexpected assembly identity: $assemblyName"
}

$classId = '{8E854243-087F-4D6C-9E0E-8622B0E50883}'
$typeLibId = '{AF3C2AF7-155F-4DDB-AC8F-D02CD58DEDC9}'
$progId = 'CodexVBE.AddIn'
$codeBase = 'file:///' + $assemblyPath.Replace([char]92, [char]47)
$registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
    [Microsoft.Win32.RegistryHive]::CurrentUser,
    [Microsoft.Win32.RegistryView]::Registry64)

function Set-StringValue([string] $path, [string] $name, [string] $value) {
    $key = $registry.CreateSubKey($path)
    try { $key.SetValue($name, $value, [Microsoft.Win32.RegistryValueKind]::String) }
    finally { $key.Dispose() }
}

try {
    $existing = $registry.OpenSubKey("Software\Classes\$progId\CLSID")
    if ($existing) {
        try {
            if ($existing.GetValue('') -ne $classId) {
                throw "The ProgID $progId already points to another COM class."
            }
        }
        finally { $existing.Dispose() }
    }

    $existingClass = $registry.OpenSubKey("Software\Classes\CLSID\$classId\InprocServer32")
    if ($existingClass) {
        try {
            $existingAssembly = [string] $existingClass.GetValue('Assembly')
            if ($existingAssembly -and -not $existingAssembly.StartsWith('CodexVBE,', [StringComparison]::Ordinal)) {
                throw "The CLSID $classId belongs to another assembly."
            }
        }
        finally { $existingClass.Dispose() }
    }

    Set-StringValue "Software\Classes\$progId" '' $progId
    Set-StringValue "Software\Classes\$progId\CLSID" '' $classId
    Set-StringValue "Software\Classes\CLSID\$classId" '' $progId
    $serverPath = "Software\Classes\CLSID\$classId\InprocServer32"
    Set-StringValue $serverPath '' 'mscoree.dll'
    Set-StringValue $serverPath 'ThreadingModel' 'Both'
    Set-StringValue $serverPath 'Class' 'CodexVBE.AddIn'
    Set-StringValue $serverPath 'Assembly' $assemblyName
    Set-StringValue $serverPath 'RuntimeVersion' 'v4.0.30319'
    Set-StringValue $serverPath 'CodeBase' $codeBase
    $versionPath = "$serverPath\0.1.0.0"
    Set-StringValue $versionPath 'Class' 'CodexVBE.AddIn'
    Set-StringValue $versionPath 'Assembly' $assemblyName
    Set-StringValue $versionPath 'RuntimeVersion' 'v4.0.30319'
    Set-StringValue $versionPath 'CodeBase' $codeBase
    Set-StringValue "Software\Classes\CLSID\$classId\ProgId" '' $progId
    Set-StringValue "Software\Classes\CLSID\$classId\TypeLib" '' $typeLibId
    Set-StringValue "Software\Classes\CLSID\$classId\Version" '' '0.1'
    $category = $registry.CreateSubKey("Software\Classes\CLSID\$classId\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}")
    $category.Dispose()

    $addinPath = "Software\Microsoft\VBA\VBE\6.0\Addins64\$progId"
    Set-StringValue $addinPath 'FriendlyName' 'CodexVBE'
    Set-StringValue $addinPath 'Description' 'Codex bridge for the VBA editor'
    $addin = $registry.OpenSubKey($addinPath, $true)
    try { $addin.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord) }
    finally { $addin.Dispose() }

    & (Join-Path $PSScriptRoot 'Register-CodexVBETypeLib.ps1')
    & (Join-Path $PSScriptRoot 'Register-ChatToolWindow.ps1')

    Write-Output "Registered $progId for the current user."
    Write-Output "Assembly: $assemblyPath"
    Write-Output 'Restart the VBA host and open its VBA editor to load the add-in.'
}
finally { $registry.Dispose() }

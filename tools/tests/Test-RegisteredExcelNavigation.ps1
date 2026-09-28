param(
    [Parameter(Mandatory=$true)][string]$AssemblyPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [ValidateSet('Navigation','FormLifecycle','FunctionalExtensions','DebuggerTree')][string]$Scenario = 'Navigation',
    [switch]$AllowTemporaryRegistration,
    [switch]$AllowTemporaryVbaAccess
)
$ErrorActionPreference = 'Stop'
if (-not $AllowTemporaryRegistration) { throw 'Explicit temporary-registration opt-in is required.' }
if (@(Get-Process EXCEL,SLDWORKS,WINWORD,MSACCESS,POWERPNT,OUTLOOK -ErrorAction SilentlyContinue).Count) { throw 'Close VBA hosts before this isolated registration test.' }
$assembly = (Resolve-Path -LiteralPath $AssemblyPath).Path
if ([Reflection.AssemblyName]::GetAssemblyName($assembly).Name -ne 'CodexVBE') { throw 'Unexpected test assembly.' }
$paths = @(
 'HKCU:\Software\Classes\CLSID\{8E854243-087F-4D6C-9E0E-8622B0E50883}\InprocServer32',
 'HKCU:\Software\Classes\CLSID\{8E854243-087F-4D6C-9E0E-8622B0E50883}\InprocServer32\0.1.0.0',
 'HKCU:\Software\Classes\CLSID\{0F4D723B-97D8-42E5-9B31-70646B97C8D2}\InprocServer32'
)
$typeLib = Join-Path (Split-Path -Parent $assembly) 'CodexVBE.tlb'
if (-not (Test-Path -LiteralPath $typeLib)) { throw 'Export the current type library before testing.' }
$typeLibKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Classes\TypeLib\{AF3C2AF7-155F-4DDB-AC8F-D02CD58DEDC9}\0.1\0\win64', $true)
$originalTypeLib = $typeLibKey.GetValue('')
if (-not $originalTypeLib) { throw 'Existing type library registration is required.' }
$backup = @()
foreach ($path in $paths) {
    $key = Get-Item -LiteralPath $path
    if (-not ([string]$key.GetValue('Assembly')).StartsWith('CodexVBE,', [StringComparison]::Ordinal)) { throw "Unexpected COM owner: $path" }
    $value = $key.GetValue('CodeBase')
    if (-not $value -or $key.GetValueKind('CodeBase') -ne [Microsoft.Win32.RegistryValueKind]::String) { throw "Missing string CodeBase: $path" }
    $backup += [pscustomobject]@{ Path = $path; CodeBase = $value }
}
$directory = [IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($directory) | Out-Null
$backup | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'temporary-registration-backup.json') -Encoding UTF8
$originalTypeLib | Set-Content -LiteralPath (Join-Path $directory 'temporary-typelib-backup.txt') -Encoding UTF8
try {
    $typeLibKey.SetValue('', $typeLib, [Microsoft.Win32.RegistryValueKind]::String)
    $codeBase = 'file:///' + $assembly.Replace([char]92,[char]47)
    foreach ($entry in $backup) { Set-ItemProperty -LiteralPath $entry.Path -Name CodeBase -Value $codeBase }
    Write-Output ('Registration process bitness: ' + ([IntPtr]::Size * 8))
    foreach ($entry in $backup) { Write-Output ((Get-ItemProperty -LiteralPath $entry.Path).CodeBase) }
    $probeName = if ($Scenario -eq 'DebuggerTree') { '../probes/Test-RegisteredDebuggerTree.ps1' } elseif ($Scenario -eq 'FunctionalExtensions') { '../probes/Test-RegisteredFunctionalExtensions.ps1' } elseif ($Scenario -eq 'FormLifecycle') { '../probes/Test-FormLifecycle.ps1' } else { '../probes/Test-NativeDefinition.ps1' }
    $probeArguments = @('-NoProfile','-STA','-File',(Join-Path $PSScriptRoot $probeName),'-AssemblyPath',$assembly,'-OutputDirectory',$directory,'-UseBridge')
    if ($AllowTemporaryVbaAccess) { $probeArguments += '-AllowTemporaryVbaAccess' }
    & powershell.exe @probeArguments
    if ($LASTEXITCODE -ne 0) { throw "Registered navigation probe exited with $LASTEXITCODE." }
} finally {
    $typeLibKey.SetValue('', $originalTypeLib, [Microsoft.Win32.RegistryValueKind]::String)
    if ($typeLibKey.GetValue('') -cne $originalTypeLib) { throw 'Type library restoration failed.' }
    $typeLibKey.Dispose()
    foreach ($entry in $backup) { Set-ItemProperty -LiteralPath $entry.Path -Name CodeBase -Value $entry.CodeBase }
    foreach ($entry in $backup) {
        if ((Get-ItemProperty -LiteralPath $entry.Path).CodeBase -cne $entry.CodeBase) { throw "Registration restoration failed: $($entry.Path)" }
    }
}
Write-Output 'Original COM CodeBase values restored and verified.'

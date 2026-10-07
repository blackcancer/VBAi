#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$InstallerPath,
    [Parameter(Mandatory=$true)][string]$PackageDirectory,
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
    [switch]$AllowCurrentUserDeployment
)
$ErrorActionPreference = 'Stop'
if (-not $AllowCurrentUserDeployment) { throw 'Explicit deployment opt-in is required.' }
if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit Windows PowerShell.' }
$installer = (Resolve-Path -LiteralPath $InstallerPath).Path
$package = (Resolve-Path -LiteralPath $PackageDirectory).Path
$evidence = [IO.Path]::GetFullPath($EvidenceDirectory)
[IO.Directory]::CreateDirectory($evidence) | Out-Null
$installation = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Programs\VBAi'
if (Test-Path -LiteralPath $installation) { throw 'Lifecycle qualification needs an absent managed install; an existing deployment is preserved.' }
$rows = New-Object Collections.Generic.List[object]
$canaryRoot = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'VBAi\InstallerQualification'
[IO.Directory]::CreateDirectory($canaryRoot) | Out-Null
$canary = Join-Path $canaryRoot ([Guid]::NewGuid().ToString('N') + '.txt')
[IO.File]::WriteAllText($canary, 'Owned installer qualification canary')
function Record([string]$Name, [scriptblock]$Body) {
    try { & $Body; $rows.Add(@{Scenario=$Name; Status='PASS'}) }
    catch { $rows.Add(@{Scenario=$Name; Status='FAIL'; Error=$_.Exception.Message}); throw }
    finally { $rows | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $evidence 'scenarios.json') -Encoding UTF8 }
}
function Run-Setup([string]$Name, [string]$Executable) {
    $log = Join-Path $evidence ($Name + '.log')
    $arguments = '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /LOG="' + $log + '"'
    $child = Start-Process -FilePath $Executable -ArgumentList $arguments -WindowStyle Hidden -PassThru
    try {
        if (-not $child.WaitForExit(120000)) { throw ('Installer outcome uncertain: PID ' + $child.Id + '. No retry or termination.') }
        if ($child.ExitCode -ne 0) { throw ('Installer returned ' + $child.ExitCode + '; inspect ' + $log) }
    } finally { $child.Dispose() }
    if ($Name -eq 'uninstall') {
        # Inno can finish its self-removal helper just after the initial process exits.
        # Observe cleanup; never dispatch another uninstall or delete its files ourselves.
        $deadline = [DateTime]::UtcNow.AddSeconds(30)
        while ([IO.Directory]::Exists($installation) -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 100 }
        if ([IO.Directory]::Exists($installation)) { throw 'Owned installation cleanup is incomplete; do not reinstall automatically.' }
    }
}
function Read-Marker { Get-Content -LiteralPath (Join-Path $installation 'vbai-installation.json') -Raw | ConvertFrom-Json }
function Assert-Installed {
    $marker = Read-Marker
    if ($marker.Version -cne '1.0.0' -or $marker.Product -cne 'VBAi' -or $marker.UpdateProtocol -ne 1) { throw 'Installed marker mismatch.' }
    foreach ($file in @(Get-ChildItem -LiteralPath $package -Recurse -File)) {
        $relative = $file.FullName.Substring($package.Length).TrimStart('\')
        if ($relative -eq 'vbai-installation.json') { continue }
        $installed = Join-Path $installation $relative
        if (-not (Test-Path -LiteralPath $installed) -or (Get-FileHash -LiteralPath $file.FullName).Hash -cne (Get-FileHash -LiteralPath $installed).Hash) { throw ('Payload readback mismatch: ' + $relative) }
    }
    $uninstaller = Join-Path $installation 'unins000.exe'
    if (-not (Test-Path -LiteralPath $uninstaller)) { throw 'Missing uninstaller.' }
    if ((Get-AuthenticodeSignature -LiteralPath $uninstaller).Status -ne 'NotSigned') { throw 'Unexpected uninstaller signature.' }
    $registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser','Registry64')
    try {
        $arp = $registry.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\Uninstall\VBAi.win-x64_is1')
        if (-not $arp) { throw 'Missing Windows Installed apps entry.' }
        try { if ($arp.GetValue('DisplayVersion') -cne '1.0.0') { throw 'Installed apps version mismatch.' } } finally { $arp.Dispose() }
        foreach ($id in @('{8E854243-087F-4D6C-9E0E-8622B0E50883}','{0F4D723B-97D8-42E5-9B31-70646B97C8D2}','{5AF2F40B-939B-4CC6-A06C-F0C79841C031}')) {
            $key = $registry.OpenSubKey('Software\Classes\CLSID\' + $id + '\InprocServer32')
            if (-not $key) { throw 'Missing COM class.' }
            try { if (-not ([Uri]$key.GetValue('CodeBase')).LocalPath.Equals((Join-Path $installation 'VBAi.dll'),[StringComparison]::OrdinalIgnoreCase)) { throw 'Wrong deployed COM path.' } }
            finally { $key.Dispose() }
        }
    } finally { $registry.Dispose() }
}
try {
    Record 'fresh-install' { Run-Setup 'install' $installer; Assert-Installed }
    $firstId = (Read-Marker).InstallationId
    Record 'repair-preserves-identity' { Run-Setup 'repair' $installer; Assert-Installed; if ((Read-Marker).InstallationId -cne $firstId) { throw 'Repair changed installation identity.' } }
    Record 'uninstall-preserves-user-data' {
        Run-Setup 'uninstall' (Join-Path $installation 'unins000.exe')
        if (Test-Path -LiteralPath (Join-Path $installation 'VBAi.dll')) { throw 'Installed DLL remains.' }
        if (-not (Test-Path -LiteralPath $canary)) { throw 'User-data canary was removed.' }
        $registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser','Registry64')
        try {
            foreach ($path in @('Software\Classes\VBAi.AddIn','Software\Classes\VBAi.ChatToolWindow','Software\Classes\VBAi.TestRuntime','Software\Microsoft\VBA\VBE\6.0\Addins64\VBAi.AddIn','Software\Microsoft\Windows\CurrentVersion\Uninstall\VBAi.win-x64_is1')) {
                $key = $registry.OpenSubKey($path)
                if ($key) { $key.Dispose(); throw ('Registry entry remains: ' + $path) }
            }
        } finally { $registry.Dispose() }
    }
    Record 'reinstall-new-identity' { Run-Setup 'reinstall' $installer; Assert-Installed; if ((Read-Marker).InstallationId -ceq $firstId) { throw 'Reinstall reused removed installation identity.' } }
    @{Status='PASS'; InstallationDirectory=$installation; InstallationId=(Read-Marker).InstallationId; InstallerSha256=(Get-FileHash -LiteralPath $installer).Hash; NativeHostLoading='NOT_RUN'; UpgradeToDifferentVersion='NOT_RUN'; FailureRecovery='NOT_RUN'; ScenarioCount=$rows.Count} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $evidence 'terminal.json') -Encoding UTF8
} catch {
    @{Status='FAIL_OR_UNCERTAIN'; Error=$_.Exception.Message; InstallationDirectory=$installation; NativeHostLoading='NOT_RUN'} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $evidence 'terminal.json') -Encoding UTF8
    throw
} finally { if (Test-Path -LiteralPath $canary) { Remove-Item -LiteralPath $canary } }

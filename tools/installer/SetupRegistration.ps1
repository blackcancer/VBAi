#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidateSet('Preflight','Install','Uninstall')][string]$Action,
    [Parameter(Mandatory=$true)][string]$InstallationDirectory,
    [string]$Version,
    [string]$LogPath,
    [switch]$ForUninstall
)
# Entry point called by Setup and its uninstaller, never by a VBA host.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2
if (-not [Environment]::Is64BitProcess) { throw '64-bit Windows PowerShell is required.' }
$transcribing = $false
if ($LogPath) { Start-Transcript -LiteralPath $LogPath -Append | Out-Null; $transcribing = $true }
$registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser', 'Registry64')
$classes = @(
    @{Id='{8E854243-087F-4D6C-9E0E-8622B0E50883}'; Prog='VBAi.AddIn'; Class='VBAi.AddIn'},
    @{Id='{0F4D723B-97D8-42E5-9B31-70646B97C8D2}'; Prog='VBAi.ChatToolWindow'; Class='VBAi.ChatToolWindow'},
    @{Id='{5AF2F40B-939B-4CC6-A06C-F0C79841C031}'; Prog='VBAi.TestRuntime'; Class='VBAi.VbaTestRuntime'}
)
$typeRoot = 'Software\Classes\TypeLib\{AF3C2AF7-155F-4DDB-AC8F-D02CD58DEDC9}'
function Read-Value([string]$Path, [string]$Name) {
    $key = $registry.OpenSubKey($Path)
    if (-not $key) { return $null }
    try { return $key.GetValue($Name, $null, 'DoNotExpandEnvironmentNames') }
    finally { $key.Dispose() }
}
function Read-Tree([string]$Path) {
    $key = $registry.OpenSubKey($Path)
    if (-not $key) { return @{Path=$Path; Exists=$false; Values=@(); Children=@()} }
    try {
        $values = @($key.GetValueNames() | Sort-Object | ForEach-Object {
            @{Name=$_; Kind=$key.GetValueKind($_); Value=$key.GetValue($_, $null, 'DoNotExpandEnvironmentNames')}
        })
        $children = @($key.GetSubKeyNames() | Sort-Object | ForEach-Object { Read-Tree ($Path + '\' + $_) })
        return @{Path=$Path; Exists=$true; Values=$values; Children=$children}
    } finally { $key.Dispose() }
}
function Restore-Tree($Tree) {
    $registry.DeleteSubKeyTree($Tree.Path, $false)
    if (-not $Tree.Exists) { return }
    $key = $registry.CreateSubKey($Tree.Path)
    try { foreach ($value in $Tree.Values) { $key.SetValue($value.Name, $value.Value, $value.Kind) } }
    finally { $key.Dispose() }
    foreach ($child in $Tree.Children) { Restore-Tree $child }
}
function Assert-HostsClosed {
    $hosts = @(Get-Process -Name EXCEL,WINWORD,POWERPNT,MSACCESS,OUTLOOK,VISIO,WINPROJ,MSPUB,SLDWORKS -ErrorAction SilentlyContinue)
    if ($hosts.Count) { throw ('Save your work and close VBA hosts: ' + (($hosts | ForEach-Object { $_.ProcessName + ' (PID ' + $_.Id + ')' }) -join ', ')) }
}
try {
    # A fixed per-user destination prevents moving an installation away from the updater.
    $directory = [IO.Path]::GetFullPath($InstallationDirectory).TrimEnd('\')
    $expected = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Programs\VBAi'
    if (-not $directory.Equals($expected, [StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected per-user installation directory.' }
    for ($ancestor = $directory; $ancestor; $ancestor = [IO.Path]::GetDirectoryName($ancestor)) {
        if ((Test-Path -LiteralPath $ancestor) -and ((Get-Item -LiteralPath $ancestor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Installation ancestors must not be reparse points.' }
    }
    Assert-HostsClosed
    $assembly = Join-Path $directory 'VBAi.dll'
    foreach ($item in $classes) {
        $prog = Read-Value ('Software\Classes\' + $item.Prog + '\CLSID') ''
        if ($prog -and $prog -ine $item.Id) { throw ('Foreign ProgID: ' + $item.Prog) }
        $server = 'Software\Classes\CLSID\' + $item.Id + '\InprocServer32'
        $owner = [string](Read-Value $server 'Assembly')
        if ($owner -and $owner -notmatch '^(VBAi|CodexVBE),') { throw ('Foreign COM class: ' + $item.Id) }
        $codeBase = Read-Value $server 'CodeBase'
        if (($Action -eq 'Uninstall' -or $ForUninstall) -and $owner) {
            if (-not $codeBase -or -not ([Uri]$codeBase).LocalPath.Equals($assembly, [StringComparison]::OrdinalIgnoreCase)) {
                throw 'Another deployment owns the current COM registration. No files will be removed.'
            }
        }
    }
    $markerPath = Join-Path $directory 'vbai-installation.json'
    $oldMarker = $null
    if (Test-Path -LiteralPath $markerPath) {
        $oldMarker = Get-Content -LiteralPath $markerPath -Raw | ConvertFrom-Json
        $identifier = [Guid]::Empty
        if ($oldMarker.Product -cne 'VBAi' -or $oldMarker.Architecture -cne 'win-x64' -or
            $oldMarker.UpdateProtocol -ne 1 -or -not [Guid]::TryParse($oldMarker.InstallationId, [ref]$identifier) -or $identifier -eq [Guid]::Empty) { throw 'Invalid installation marker; preserve this installation for diagnosis.' }
    } elseif ($Action -eq 'Uninstall' -or $ForUninstall) { throw 'Missing installation marker; automatic removal refused.' }
    if ($Action -eq 'Preflight') { Write-Output 'Preflight completed without registry or file mutation.'; return }
    if ($Action -eq 'Install' -and $Version -notmatch '^\d+\.\d+\.\d+(-[0-9A-Za-z.-]+)?$') { throw 'Missing or invalid product version.' }
    $paths = @($typeRoot)
    foreach ($item in $classes) {
        $paths += 'Software\Classes\' + $item.Prog
        $paths += 'Software\Classes\CLSID\' + $item.Id
    }
    foreach ($location in @('Addins64','Addins')) {
        foreach ($prog in @('VBAi.AddIn','CodexVBE.AddIn')) { $paths += 'Software\Microsoft\VBA\VBE\6.0\' + $location + '\' + $prog }
    }
    $paths += @('Software\Classes\CodexVBE.AddIn','Software\Classes\CodexVBE.ChatToolWindow')
    $before = @($paths | ForEach-Object { Read-Tree $_ })
    $backupDirectory = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) ('VBAi\SetupBackups\' + [Guid]::NewGuid().ToString('N'))
    [IO.Directory]::CreateDirectory($backupDirectory) | Out-Null
    $backup = Join-Path $backupDirectory 'registry-before.clixml'
    $before | Export-Clixml -LiteralPath $backup -Depth 100
    Write-Output ('Registry recovery snapshot: ' + $backup)
    try {
        if ($Action -eq 'Install') {
            & (Join-Path $PSScriptRoot 'Install-VBAi.ps1') -Direct -AssemblyDirectory $directory
            foreach ($item in $classes) {
                $server = 'Software\Classes\CLSID\' + $item.Id + '\InprocServer32'
                $registered = Read-Value $server 'CodeBase'
                if (-not $registered -or -not ([Uri]$registered).LocalPath.Equals($assembly, [StringComparison]::OrdinalIgnoreCase)) { throw ('COM readback failed: ' + $item.Prog) }
                if ((Read-Value ('Software\Classes\' + $item.Prog + '\CLSID') '') -ine $item.Id) { throw 'ProgID readback failed.' }
            }
            if ((Read-Value 'Software\Microsoft\VBA\VBE\6.0\Addins64\VBAi.AddIn' 'LoadBehavior') -ne 3) { throw 'VBE discovery readback failed.' }
            $registeredTlb = Read-Value ($typeRoot + '\0.1\0\win64') ''
            if (-not $registeredTlb -or -not $registeredTlb.Equals((Join-Path $directory 'VBAi.tlb'), [StringComparison]::OrdinalIgnoreCase)) { throw 'Type library readback failed.' }
            $installationId = if ($oldMarker) { $oldMarker.InstallationId } else { [Guid]::NewGuid().ToString() }
            $marker = @{Product='VBAi'; Architecture='win-x64'; UpdateProtocol=1; InstallationId=$installationId; Version=$Version}
            $temporaryMarker = Join-Path $directory ('marker-' + [Guid]::NewGuid().ToString('N') + '.tmp')
            [IO.File]::WriteAllText($temporaryMarker, ($marker | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))
            if (Test-Path -LiteralPath $markerPath) { [IO.File]::Replace($temporaryMarker, $markerPath, (Join-Path $backupDirectory 'marker-before.json')) }
            else { [IO.File]::Move($temporaryMarker, $markerPath) }
        } else {
            # The shared machine ProgID contains no path and may serve another user's installation.
            & (Join-Path $PSScriptRoot 'Uninstall-VBAi.ps1') -Direct -AssemblyDirectory $directory -PreserveMachineMapping
            foreach ($item in $classes) {
                if (Read-Value ('Software\Classes\CLSID\' + $item.Id + '\InprocServer32') 'CodeBase') { throw 'COM removal readback failed.' }
            }
            Remove-Item -LiteralPath $markerPath
        }
    } catch {
        $failure = $_
        try {
            foreach ($tree in $before) { Restore-Tree $tree }
            $restored = @($paths | ForEach-Object { Read-Tree $_ })
            if (([Management.Automation.PSSerializer]::Serialize($before, 100)) -cne ([Management.Automation.PSSerializer]::Serialize($restored, 100))) { throw 'Registry recovery readback differs.' }
            Write-Output 'The exact prior HKCU registry trees were restored.'
        } catch { throw ('Partial recovery. Keep snapshot ' + $backup + '. ' + $_.Exception.Message + ' Original failure: ' + $failure.Exception.Message) }
        throw $failure
    }
    Write-Output ($Action + ' registration completed. Native host loading was not tested.')
} finally {
    $registry.Dispose()
    if ($transcribing) { Stop-Transcript | Out-Null }
}

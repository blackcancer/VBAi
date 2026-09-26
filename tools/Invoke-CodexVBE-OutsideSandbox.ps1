param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Install', 'Uninstall', 'Verify')]
    [string] $Action,
    [switch] $TaskHost,
    [switch] $Worker,
    [string] $StatusPath,
    [string] $ExpectedSid
)

$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit PowerShell.' }
if ($ExpectedSid -and [Security.Principal.WindowsIdentity]::GetCurrent().User.Value -ne $ExpectedSid) {
    $message = 'Windows elevation changed the user account; CodexVBE was not installed for another user.'
    if ($StatusPath) { Set-Content -LiteralPath $StatusPath -Value "ERROR $message" -Encoding UTF8 }
    throw $message
}

$classId = '{8E854243-087F-4D6C-9E0E-8622B0E50883}'
$typeLibId = '{AF3C2AF7-155F-4DDB-AC8F-D02CD58DEDC9}'
$progId = 'CodexVBE.AddIn'
$projectRoot = Split-Path -Parent $PSScriptRoot
$dll = Join-Path $projectRoot 'bin\Debug\net48\CodexVBE.dll'
$typeLib = Join-Path $projectRoot 'bin\Debug\net48\CodexVBE.tlb'
$expectedCodeBase = 'file:///' + ([IO.Path]::GetFullPath($dll)).Replace([char]92, [char]47)

function Assert-Registration([bool] $shouldExist) {
    $registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser', 'Registry64')
    try {
        $addin = $registry.OpenSubKey("Software\Microsoft\VBA\VBE\6.0\Addins64\$progId")
        $progid = $registry.OpenSubKey("Software\Classes\$progId\CLSID")
        $server = $registry.OpenSubKey("Software\Classes\CLSID\$classId\InprocServer32")
        $registeredTypeLib = $registry.OpenSubKey("Software\Classes\TypeLib\$typeLibId\0.1\0\win64")
        try {
            if ($shouldExist) {
                if (-not $addin -or -not $progid -or -not $server -or -not $registeredTypeLib) { throw 'A VBE, COM or type library registration key is missing.' }
                if ($addin.GetValue('LoadBehavior') -ne 3) { throw 'LoadBehavior is not 3.' }
                if ($progid.GetValue('') -ne $classId) { throw 'The ProgID points to another CLSID.' }
                if ($server.GetValue('CodeBase') -cne $expectedCodeBase) { throw 'The COM CodeBase points to another DLL.' }
                if ($registeredTypeLib.GetValue('') -cne ([IO.Path]::GetFullPath($typeLib))) { throw 'The type library points to another file.' }
            }
            elseif ($addin -or $progid -or $server -or $registeredTypeLib) {
                throw 'A CodexVBE registration key remains after uninstall.'
            }
        }
        finally {
            if ($addin) { $addin.Dispose() }
            if ($progid) { $progid.Dispose() }
            if ($server) { $server.Dispose() }
            if ($registeredTypeLib) { $registeredTypeLib.Dispose() }
        }
    }
    finally { $registry.Dispose() }
}

if ($Worker) {
    try {
        if ($Action -eq 'Install') { & (Join-Path $PSScriptRoot 'Install-CodexVBE.ps1') -Direct | Out-Null }
        if ($Action -eq 'Uninstall') { & (Join-Path $PSScriptRoot 'Uninstall-CodexVBE.ps1') -Direct | Out-Null }
        Assert-Registration ($Action -ne 'Uninstall')
        Set-Content -LiteralPath $StatusPath -Value "SUCCESS $Action" -Encoding ASCII
    }
    catch {
        Set-Content -LiteralPath $StatusPath -Value "ERROR $($_.Exception.Message)" -Encoding UTF8
        throw
    }
    return
}

if (-not $TaskHost) {
    $status = Join-Path $PSScriptRoot ('.codexvbe-' + [guid]::NewGuid().ToString('N') + '.status')
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action $Action -TaskHost -StatusPath `"$status`" -ExpectedSid $sid"
    try {
        $process = Start-Process -FilePath "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList $arguments -Verb RunAs -WindowStyle Hidden -Wait -PassThru
        if (-not (Test-Path -LiteralPath $status)) { throw "Independent installation did not return a result (task host exit code $($process.ExitCode))." }
        $result = Get-Content -LiteralPath $status -Raw
        if (-not $result.StartsWith("SUCCESS $Action", [StringComparison]::Ordinal)) { throw $result.Trim() }
        if ($process.ExitCode -ne 0) { throw "Elevated task host failed with exit code $($process.ExitCode)." }
        Write-Output "CodexVBE $Action verified outside the calling process."
    }
    finally { Remove-Item -LiteralPath $status -ErrorAction SilentlyContinue }
    return
}

$taskName = 'CodexVBE-Install-' + [guid]::NewGuid().ToString('N')
$account = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action $Action -Worker -StatusPath `"$StatusPath`" -ExpectedSid $ExpectedSid"
$taskAction = New-ScheduledTaskAction -Execute "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument $arguments
$principal = New-ScheduledTaskPrincipal -UserId $account -LogonType Interactive -RunLevel Highest
try {
    Register-ScheduledTask -TaskName $taskName -Action $taskAction -Principal $principal -Force | Out-Null
    Start-ScheduledTask -TaskName $taskName
    for ($i = 0; $i -lt 45; $i++) {
        if (Test-Path -LiteralPath $StatusPath) { break }
        Start-Sleep -Seconds 1
    }
    if (-not (Test-Path -LiteralPath $StatusPath)) { throw 'The independent task timed out.' }
    $result = Get-Content -LiteralPath $StatusPath -Raw
    if (-not $result.StartsWith("SUCCESS $Action", [StringComparison]::Ordinal)) { throw $result.Trim() }
}
finally { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue }

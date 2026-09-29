param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Install', 'Uninstall', 'Verify')]
    [string] $Action,
    [switch] $TaskHost,
    [switch] $Worker,
    [string] $StatusPath,
    [string] $ExpectedSid,
    [string] $ExpectedAssemblyPath
)

$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit PowerShell.' }
if ($ExpectedSid -and [Security.Principal.WindowsIdentity]::GetCurrent().User.Value -ne $ExpectedSid) {
    $message = 'Windows elevation changed the user account; VBAi was not installed for another user.'
    if ($StatusPath) { Set-Content -LiteralPath $StatusPath -Value "ERROR $message" -Encoding UTF8 }
    throw $message
}

$classId = '{8E854243-087F-4D6C-9E0E-8622B0E50883}'
$typeLibId = '{AF3C2AF7-155F-4DDB-AC8F-D02CD58DEDC9}'
$progId = 'VBAi.AddIn'
$chatProgId = 'VBAi.ChatToolWindow'
$chatClassId = '{0F4D723B-97D8-42E5-9B31-70646B97C8D2}'
$projectRoot = Split-Path -Parent $PSScriptRoot
$dll = if ($ExpectedAssemblyPath) { [IO.Path]::GetFullPath($ExpectedAssemblyPath) }
    else { Join-Path $projectRoot 'bin\Debug\net48\VBAi.dll' }
$typeLib = Join-Path (Split-Path -Parent $dll) 'VBAi.tlb'
$expectedCodeBase = 'file:///' + ([IO.Path]::GetFullPath($dll)).Replace([char]92, [char]47)

function Assert-Registration([bool] $shouldExist) {
    $registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser', 'Registry64')
    $machine = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64')
    try {
        $addin = $registry.OpenSubKey("Software\Microsoft\VBA\VBE\6.0\Addins64\$progId")
        $progid = $registry.OpenSubKey("Software\Classes\$progId\CLSID")
        $server = $registry.OpenSubKey("Software\Classes\CLSID\$classId\InprocServer32")
        $chatProgIdKey = $registry.OpenSubKey("Software\Classes\$chatProgId\CLSID")
        $machineChatProgId = $machine.OpenSubKey("Software\Classes\$chatProgId\CLSID")
        $chatServer = $registry.OpenSubKey("Software\Classes\CLSID\$chatClassId\InprocServer32")
        $registeredTypeLib = $registry.OpenSubKey("Software\Classes\TypeLib\$typeLibId\0.1\0\win64")
        try {
            if ($shouldExist) {
                if (-not $addin) { throw 'VBE Addins64 registration is missing.' }
                if (-not $progid) { throw 'Add-in ProgID registration is missing.' }
                if (-not $server) { throw 'Add-in COM server registration is missing.' }
                if (-not $registeredTypeLib) { throw 'Type library registration is missing.' }
                if (-not $chatProgIdKey) { throw 'Chat control ProgID registration is missing.' }
                if (-not $machineChatProgId) { throw 'Machine chat ProgID registration required by VBE is missing.' }
                if (-not $chatServer) { throw 'Chat control COM server registration is missing.' }
                if ($addin.GetValue('LoadBehavior') -ne 3) { throw 'LoadBehavior is not 3.' }
                if ($progid.GetValue('') -ne $classId) { throw 'The ProgID points to another CLSID.' }
                if ($server.GetValue('CodeBase') -cne $expectedCodeBase) { throw 'The COM CodeBase points to another DLL.' }
                if ($chatProgIdKey.GetValue('') -ne $chatClassId) { throw 'The chat control ProgID points to another CLSID.' }
                if ($machineChatProgId.GetValue('') -ne $chatClassId) { throw 'The machine chat ProgID points to another CLSID.' }
                if ($chatServer.GetValue('CodeBase') -cne $expectedCodeBase) { throw 'The chat control CodeBase points to another DLL.' }
                if ($registeredTypeLib.GetValue('') -cne ([IO.Path]::GetFullPath($typeLib))) { throw 'The type library points to another file.' }
            }
            elseif ($addin -or $progid -or $server -or $registeredTypeLib -or $chatProgIdKey -or $machineChatProgId -or $chatServer) {
                throw 'A VBAi registration key remains after uninstall.'
            }
        }
        finally {
            if ($addin) { $addin.Dispose() }
            if ($progid) { $progid.Dispose() }
            if ($server) { $server.Dispose() }
            if ($chatProgIdKey) { $chatProgIdKey.Dispose() }
            if ($machineChatProgId) { $machineChatProgId.Dispose() }
            if ($chatServer) { $chatServer.Dispose() }
            if ($registeredTypeLib) { $registeredTypeLib.Dispose() }
        }
    }
    finally { $machine.Dispose(); $registry.Dispose() }
}

if ($Worker) {
    try {
        if ($Action -eq 'Install') { & (Join-Path $PSScriptRoot 'Install-VBAi.ps1') -Direct | Out-Null }
        if ($Action -eq 'Uninstall') { & (Join-Path $PSScriptRoot 'Uninstall-VBAi.ps1') -Direct | Out-Null }
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
    $status = Join-Path $PSScriptRoot ('.VBAi-' + [guid]::NewGuid().ToString('N') + '.status')
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action $Action -TaskHost -StatusPath `"$status`" -ExpectedSid $sid"
    if ($ExpectedAssemblyPath) { $arguments += " -ExpectedAssemblyPath `"$dll`"" }
    try {
        $process = Start-Process -FilePath "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList $arguments -Verb RunAs -WindowStyle Hidden -Wait -PassThru
        if (-not (Test-Path -LiteralPath $status)) { throw "Independent installation did not return a result (task host exit code $($process.ExitCode))." }
        $result = Get-Content -LiteralPath $status -Raw
        if (-not $result.StartsWith("SUCCESS $Action", [StringComparison]::Ordinal)) { throw $result.Trim() }
        if ($process.ExitCode -ne 0) { throw "Elevated task host failed with exit code $($process.ExitCode)." }
        Write-Output "VBAi $Action verified outside the calling process."
    }
    finally { Remove-Item -LiteralPath $status -ErrorAction SilentlyContinue }
    return
}

$taskName = 'VBAi-Install-' + [guid]::NewGuid().ToString('N')
$account = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action $Action -Worker -StatusPath `"$StatusPath`" -ExpectedSid $ExpectedSid"
if ($ExpectedAssemblyPath) { $arguments += " -ExpectedAssemblyPath `"$dll`"" }
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

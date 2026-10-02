param(
    [Parameter(Mandatory=$true)][string]$ScriptPath,
    [Parameter(Mandatory=$true)][string]$ExpectedDesktop,
    [Parameter(Mandatory=$true)][string]$ProofPath,
    [Parameter(Mandatory=$true)][string]$HelperAssembly
)
$ErrorActionPreference='Stop'
$env:PSModulePath=(Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules')+';'+$env:PSModulePath
Import-Module Microsoft.PowerShell.Utility
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
if($PSVersionTable.PSEdition -ne 'Desktop' -or -not [Environment]::Is64BitProcess){throw 'Desktop PowerShell x64 is required'}
if(-not [IO.Path]::IsPathRooted($ScriptPath) -or -not(Test-Path -LiteralPath $ScriptPath -PathType Leaf) -or
    -not [IO.Path]::IsPathRooted($ProofPath) -or (Test-Path -LiteralPath $ProofPath)){throw 'Exact new worker paths required'}
[Reflection.Assembly]::LoadFrom($HelperAssembly) | Out-Null
# The helper is internal in the shared test assembly; reflection calls only its fixed desktop guard/readers.
$desktopType=[Reflection.Assembly]::LoadFrom($HelperAssembly).GetType('VBAi.Tests.Integration.IsolatedTestDesktop',$true)
$flags=[Reflection.BindingFlags]'Static,NonPublic'
$desktopType.GetMethod('RequireCurrent',$flags).Invoke($null,@($ExpectedDesktop)) | Out-Null
$thread=$desktopType.GetMethod('GetCurrentThreadId',$flags).Invoke($null,@())
$actual=$desktopType.GetMethod('DesktopName',$flags).Invoke($null,@([uint32]$thread))
$input=$desktopType.GetMethod('InputDesktopName',$flags).Invoke($null,@())
@{ExpectedDesktop=$ExpectedDesktop;ActualDesktop=$actual;InputDesktop=$input;WorkerPid=$PID;ThreadId=$thread;DesktopSwitches=0;Utc=[DateTime]::UtcNow.ToString('o')} |
    ConvertTo-Json | Set-Content -LiteralPath $ProofPath -Encoding UTF8
$env:VBAi_TEST_DESKTOP_NAME=$ExpectedDesktop
$env:VBAi_QUALIFICATION_DESKTOP=$ExpectedDesktop
$evidence=Join-Path (Split-Path $ProofPath) 'testhosts'
[IO.Directory]::CreateDirectory($evidence) | Out-Null
$env:VBAi_QUALIFICATION_DESKTOP_EVIDENCE=$evidence
& $ScriptPath
exit $LASTEXITCODE

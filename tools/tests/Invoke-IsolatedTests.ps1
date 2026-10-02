#requires -Version 5.1
param(
    [Parameter(Mandatory=$true)][string]$ScriptPath,
    [Parameter(Mandatory=$true)][string]$HelperAssembly,
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory
)
$ErrorActionPreference='Stop'
$env:PSModulePath=(Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules')+';'+$env:PSModulePath
Import-Module Microsoft.PowerShell.Utility
Import-Module ScheduledTasks
foreach($path in @($ScriptPath,$HelperAssembly)){
    if(-not [IO.Path]::IsPathRooted($path) -or -not(Test-Path -LiteralPath $path -PathType Leaf)){throw 'Existing absolute candidate files required'}
}
if(-not [IO.Path]::IsPathRooted($EvidenceDirectory) -or (Test-Path -LiteralPath $EvidenceDirectory)){throw 'A fresh absolute evidence directory is required'}
$worker=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot 'Invoke-IsolatedDesktopWorker.ps1')).Path
[IO.Directory]::CreateDirectory($EvidenceDirectory) | Out-Null
$launch=Join-Path $EvidenceDirectory 'launcher.ps1'
$terminal=Join-Path $EvidenceDirectory 'limited-terminal.json'
$taskName='VBAi-IsolatedTests-'+[Guid]::NewGuid().ToString('N')
# Values are data: use a JSON plan and fixed script, never interpolate shell command text.
@{Script=$ScriptPath;Helper=$HelperAssembly;Worker=$worker;Output=(Join-Path $EvidenceDirectory 'desktop');Terminal=$terminal;
    HelperSha256=(Get-FileHash -LiteralPath $HelperAssembly).Hash;ScriptSha256=(Get-FileHash -LiteralPath $ScriptPath).Hash;
    TaskName=$taskName;User=[Security.Principal.WindowsIdentity]::GetCurrent().Name;Utc=[DateTime]::UtcNow.ToString('o')} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $EvidenceDirectory 'intent.json') -Encoding UTF8
$body=@'
$ErrorActionPreference='Stop'
$env:PSModulePath=(Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules')+';'+$env:PSModulePath
Import-Module Microsoft.PowerShell.Utility
$plan=Get-Content (Join-Path $PSScriptRoot 'intent.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if((Test-Path -LiteralPath $plan.Output) -or (Test-Path -LiteralPath $plan.Terminal)){throw 'One-shot launch already claimed'}
if((Get-FileHash -LiteralPath $plan.Helper).Hash -cne $plan.HelperSha256 -or
    (Get-FileHash -LiteralPath $plan.Script).Hash -cne $plan.ScriptSha256){throw 'Frozen launch files changed'}
$code=1
try {
    & $plan.Helper --run $plan.Script $plan.Output $plan.Worker
    $code=$LASTEXITCODE
} finally {
    @{State='CHILD_TERMINAL';ExitCode=$code;User=[Security.Principal.WindowsIdentity]::GetCurrent().Name;Utc=[DateTime]::UtcNow.ToString('o')} |
        ConvertTo-Json | Set-Content -LiteralPath $plan.Terminal -Encoding UTF8
}
exit $code
'@
[IO.File]::WriteAllText($launch,$body,[Text.UTF8Encoding]::new($false))
$ps=Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
if($launch.Contains('"')){throw 'Invalid launcher path'}
$action=New-ScheduledTaskAction -Execute $ps -Argument ('-NoProfile -NonInteractive -WindowStyle Hidden -File "'+$launch+'"')
$principal=New-ScheduledTaskPrincipal -UserId ([Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
# No recurring trigger and no execution time limit that could stop an uncertain native host.
$settings=New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $settings | Out-Null
Start-ScheduledTask -TaskName $taskName
@{State='STARTED_ONCE';TaskName=$taskName;Terminal=$terminal;DesktopOutput=(Join-Path $EvidenceDirectory 'desktop');
    NoDesktopSwitch=$true;Utc=[DateTime]::UtcNow.ToString('o')} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $EvidenceDirectory 'wrapper.json') -Encoding UTF8
Get-Content -LiteralPath (Join-Path $EvidenceDirectory 'wrapper.json') -Raw -Encoding UTF8

#requires -Version 5.1
param(
    [Parameter(Mandatory=$true)][string]$ScriptPath,
    [Parameter(Mandatory=$true)][string]$HelperAssembly,
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
    [switch]$DirectGuiLauncher
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
$native=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot 'IsolatedHelperProcessNative.cs')).Path
$taskName='VBAi-IsolatedTests-'+[Guid]::NewGuid().ToString('N')
# Values are data: use a JSON plan and fixed script, never interpolate shell command text.
@{Script=$ScriptPath;Helper=$HelperAssembly;Worker=$worker;Output=(Join-Path $EvidenceDirectory 'desktop');Terminal=$terminal;
    HelperSha256=(Get-FileHash -LiteralPath $HelperAssembly).Hash;ScriptSha256=(Get-FileHash -LiteralPath $ScriptPath).Hash;WorkerSha256=(Get-FileHash -LiteralPath $worker).Hash;
    NativeSource=$native;NativeSourceSha256=(Get-FileHash -LiteralPath $native).Hash;UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value;TaskName=$taskName;User=[Security.Principal.WindowsIdentity]::GetCurrent().Name;Utc=[DateTime]::UtcNow.ToString('o')} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $EvidenceDirectory 'intent.json') -Encoding UTF8
$body=@'
$ErrorActionPreference='Stop'
$env:PSModulePath=(Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules')+';'+$env:PSModulePath
Import-Module Microsoft.PowerShell.Utility
$plan=Get-Content (Join-Path $PSScriptRoot 'intent.json') -Raw -Encoding UTF8|ConvertFrom-Json
if((Test-Path $plan.Output) -or (Test-Path $plan.Terminal)){throw 'One-shot launch already claimed'}
foreach($f in @(@{Path=$plan.Helper;Hash=$plan.HelperSha256},@{Path=$plan.Script;Hash=$plan.ScriptSha256},@{Path=$plan.Worker;Hash=$plan.WorkerSha256},@{Path=$plan.NativeSource;Hash=$plan.NativeSourceSha256})){if((Get-FileHash -LiteralPath $f.Path).Hash -cne $f.Hash){throw 'Frozen launch files changed'}}
if([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -cne $plan.UserSid){throw 'Exact original task actor required'}
foreach($path in @($plan.Helper,$plan.Script,$plan.Output,$plan.Worker)){if(-not [IO.Path]::IsPathRooted($path) -or $path.Contains('"') -or $path.IndexOf([char]0) -ge 0){throw 'Absolute quote-free native argument paths required'}}
function Write-Durable([string]$Path,$Value){$bytes=[Text.UTF8Encoding]::new($false).GetBytes((ConvertTo-Json -InputObject $Value -Depth 12));$file=[IO.File]::Open($Path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None);try{$file.Write($bytes,0,$bytes.Length);$file.Flush($true)}finally{$file.Dispose()}}
Add-Type -Path $plan.NativeSource
$process=[Diagnostics.Process]::new();$start=[Diagnostics.ProcessStartInfo]::new()
$start.FileName=$plan.Helper;$start.Arguments='--run "'+$plan.Script+'" "'+$plan.Output+'" "'+$plan.Worker+'"'
$start.WorkingDirectory=(Get-Location).ProviderPath
$start.UseShellExecute=$false;$start.CreateNoWindow=$true;$start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true
$process.StartInfo=$start
$code=1;$failure=$null;$started=$false;$startEntered=$false;$startReturned=$false;$observedExit=$false;$originalHandle=[IntPtr]::Zero
try{
    Write-Durable (Join-Path $PSScriptRoot 'helper-start-claim.json') @{State='START_CLAIMED_ONCE';ActorPid=$PID;UserSid=$plan.UserSid;Executable=$plan.Helper;Arguments=$start.Arguments;WorkingDirectory=$start.WorkingDirectory;UseShellExecute=$false;CreateNoWindow=$true;InvocationLimit=1;Utc=[DateTime]::UtcNow.ToString('o')}
    $startEntered=$true;$startResult=$process.Start();$startReturned=$true
    if(-not $startResult){throw 'Original helper creation returned false'};$started=$true
    $originalHandle=$process.Handle
    $stdoutTask=$process.StandardOutput.ReadToEndAsync();$stderrTask=$process.StandardError.ReadToEndAsync()
    $identity=[IsolatedHelperProcessNative]::Describe($originalHandle,[uint32]$process.Id,$plan.Helper)
    Write-Durable (Join-Path $PSScriptRoot 'helper-original-started.json') @{State='ORIGINAL_CREATED_HELPER_IDENTIFIED';Original=$identity;ActorPid=$PID;UseShellExecute=$false;CreateNoWindow=$true;BothStreamsReadAsync=$true;Utc=[DateTime]::UtcNow.ToString('o')}
    $process.WaitForExit();$nativeCode=[IsolatedHelperProcessNative]::ExitCode($originalHandle);$observedExit=$true
    $code=$process.ExitCode;$unsignedCode=[BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$code),0)
    Write-Durable (Join-Path $PSScriptRoot 'helper-original-exit.json') @{State='ORIGINAL_HELD_HELPER_EXIT_OBSERVED';Original=$identity;NativeExitCode=$nativeCode;ProcessExitCode=$code;ExitCodesEquivalent=($nativeCode -eq $unsignedCode);ForcedTermination=$false;Utc=[DateTime]::UtcNow.ToString('o')}
    if($nativeCode -ne $unsignedCode){throw 'Original native/Process exit codes disagree'}
    [IO.File]::WriteAllText((Join-Path $PSScriptRoot 'helper-stdout.log'),$stdoutTask.GetAwaiter().GetResult(),[Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $PSScriptRoot 'helper-stderr.log'),$stderrTask.GetAwaiter().GetResult(),[Text.UTF8Encoding]::new($false))
}catch{
    $failure=$_.Exception.ToString();$code=1
    if(($started -or ($startEntered -and -not $startReturned)) -and -not $observedExit){
        $retainedPid=$null;try{if($originalHandle -eq [IntPtr]::Zero){$originalHandle=$process.Handle};$retainedPid=$process.Id}catch{}
        try { Write-Durable (Join-Path $PSScriptRoot 'launcher-retained.json') @{State='ORIGINAL_HELPER_LIFETIME_UNCERTAIN_RETAINED';Failure=$failure;OriginalProcessHandle=$originalHandle.ToInt64();HelperProcessId=$retainedPid;StartEntered=$startEntered;StartReturned=$startReturned;AutomaticKill=$false;ReplayAllowed=$false;Utc=[DateTime]::UtcNow.ToString('o')} } catch { $failure += " | Retention receipt: " + $_.Exception.ToString() }
        for(;;){if($originalHandle -ne [IntPtr]::Zero -and [IsolatedHelperProcessNative]::WaitForSingleObject($originalHandle,250) -eq 0){$observedExit=$true;break};[Threading.Thread]::Sleep(250)}
    }
}finally{
    if((-not $startEntered) -or ($startReturned -and -not $started) -or $observedExit){
        $process.Dispose()
        Write-Durable $plan.Terminal @{State='CHILD_TERMINAL';ExitCode=$code;User=[Security.Principal.WindowsIdentity]::GetCurrent().Name;Failure=$failure;OriginalHelperStarted=$started;OriginalHelperExitObserved=$observedExit;UseShellExecute=$false;CreateNoWindow=$true;Utc=[DateTime]::UtcNow.ToString('o')}
    }
}
exit $code

'@
[IO.File]::WriteAllText($launch,$body,[Text.UTF8Encoding]::new($false))
$ps=Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
if($launch.Contains('"')){throw 'Invalid launcher path'}
$action=New-ScheduledTaskAction -Execute $ps -Argument ('-NoProfile -NonInteractive -WindowStyle Hidden -File "'+$launch+'"')
if($DirectGuiLauncher){
    $planPath=Join-Path $EvidenceDirectory 'intent.json'
    if($planPath.Contains('"')){throw 'Invalid launcher plan path'}
    $action=New-ScheduledTaskAction -Execute $HelperAssembly -Argument ('--run-plan "'+$planPath+'"')
}
$principal=New-ScheduledTaskPrincipal -UserId ([Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
# No recurring trigger and no execution time limit that could stop an uncertain native host.
$settings=New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew -DisallowHardTerminate -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $settings | Out-Null
Start-ScheduledTask -TaskName $taskName
@{State='STARTED_ONCE';TaskName=$taskName;Terminal=$terminal;DesktopOutput=(Join-Path $EvidenceDirectory 'desktop');
    NoDesktopSwitch=$true;DirectGuiLauncher=[bool]$DirectGuiLauncher;Utc=[DateTime]::UtcNow.ToString('o')} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $EvidenceDirectory 'wrapper.json') -Encoding UTF8
Get-Content -LiteralPath (Join-Path $EvidenceDirectory 'wrapper.json') -Raw -Encoding UTF8

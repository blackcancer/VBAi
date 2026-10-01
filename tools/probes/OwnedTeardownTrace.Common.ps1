$ErrorActionPreference = 'Stop'
if (-not ('VBAiOwnedTeardownNative' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class VBAiOwnedTeardownNative {
 [DllImport("kernel32.dll", SetLastError=true)] public static extern bool CheckRemoteDebuggerPresent(IntPtr handle, out bool present);
 [DllImport("kernel32.dll", SetLastError=true)] public static extern bool DebugBreakProcess(IntPtr handle);
 [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] public static extern bool QueryFullProcessImageName(IntPtr handle,uint flags,StringBuilder text,ref int length);
 [DllImport("kernel32.dll")] public static extern uint GetACP();
}
'@
}
function Get-OwnedTeardownImage($process) {
    $text = New-Object Text.StringBuilder 32768; $length = 32768
    if (-not [VBAiOwnedTeardownNative]::QueryFullProcessImageName($process.Handle, 0, $text, [ref]$length)) {
        throw ('Exact owned image query failed: ' + [Runtime.InteropServices.Marshal]::GetLastWin32Error())
    }
    return $text.ToString()
}
function Write-TeardownJson([string]$path, $value) {
    [IO.File]::WriteAllText($path, ($value | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding $false))
}
function Write-TeardownMarker([string]$path, $value) {
    if([IO.File]::Exists($path)){throw 'Diagnostic marker already exists; no replay.'}
    $temporary=$path+'.'+[Guid]::NewGuid().ToString('N')+'.tmp'
    Write-TeardownJson $temporary $value
    [IO.File]::Move($temporary,$path)
}
function Get-TeardownCommands([string]$log, [int]$targetProcessId, [string]$nonce) {
    if ($log.Contains('"') -or $nonce -notmatch '^[a-f0-9]{32}$') { throw 'Unsupported debugger identity/path.' }
    $safeLog = $log.Replace('\','/')
    # -c is first-chance only; failfast can arrive directly at second chance.
    # https://learn.microsoft.com/en-us/windows-hardware/drivers/debuggercmds/sx--sxd--sxe--sxi--sxn--sxr--sx---set-exceptions-
    # Continue fatal exceptions unhandled; qd would mark the outstanding event handled.
    # https://learn.microsoft.com/en-us/windows-hardware/drivers/debuggercmds/gn--gn--go-with-exception-not-handled-
    # .ecxr is documented for minidumps, not this live exception event.
    # https://learn.microsoft.com/en-us/windows-hardware/drivers/debuggercmds/r--registers-
    $capture = '.echo VBAI_TEARDOWN_EXCEPTION_BEGIN; .lastevent; .exr -1; .echo VBAI_TEARDOWN_REGISTER_MODE LiveEventThread; r; kv; .echo VBAI_TEARDOWN_EXCEPTION_END; gn'
    $forward='.echo VBAI_FIRSTCHANCE_AV_BEGIN '+$targetProcessId+' '+$nonce+'; .lastevent; .echo VBAI_FIRSTCHANCE_AV_END; gn'
    return ".logopen /u `"$safeLog`"`n" +
        "sxe -c `"$capture`" -c2 `"$capture`" 0xc0000409`n" +
        "sxd -c `"$forward`" -c2 `"$capture`" 0xc0000005`n" +
        "sxn -c `"qd`" epr`n" +
        ".echo VBAI_TEARDOWN_READY $targetProcessId $nonce`ng`n"
}
function Get-TeardownCapturedCode([string]$text, [int]$targetProcessId, [string]$expectedCode='') {
    if($targetProcessId -le 0){return $null}
    foreach($match in [regex]::Matches($text,'(?ms)^VBAI_TEARDOWN_EXCEPTION_BEGIN\r?\n(.*?)^VBAI_TEARDOWN_EXCEPTION_END\r?$')){
        $body=$match.Groups[1].Value
        $event=[regex]::Match($body,'(?im)^\s*Last event:\s*([0-9a-f]+)\.([0-9a-f]+):[^\r\n]*\bcode\s+(c0000409|c0000005)\b[^\r\n]*')
        if(-not $event.Success -or [Convert]::ToInt64($event.Groups[1].Value,16) -ne $targetProcessId){continue}
        $code=$event.Groups[3].Value.ToLowerInvariant()
        if($expectedCode -ne '' -and $code -ne $expectedCode.ToLowerInvariant()){continue}
        if($code -eq 'c0000005' -and $event.Value -notmatch '(?i)second chance'){continue}
        $rip=[regex]::Match($body,'(?im)^\s*rip=([0-9a-f]+)')
        if($body -notmatch ('(?im)^\s*ExceptionCode:\s*'+$code+'\b') -or
            $body -notmatch '(?m)^VBAI_TEARDOWN_REGISTER_MODE LiveEventThread\r?$' -or
            -not $rip.Success -or [Convert]::ToUInt64($rip.Groups[1].Value,16) -eq 0 -or
            $body -notmatch 'Child-SP\s+RetAddr' -or $body -notmatch '(?im)^\s*[0-9a-f`]{8,}\s+[0-9a-f`]{8,}\s+\S'){continue}
        return $code
    }
    return $null
}
function Test-TeardownExceptionCapture([string]$text, [int]$targetProcessId, [string]$expectedCode='') {
    return $null -ne (Get-TeardownCapturedCode $text $targetProcessId $expectedCode)
}
function Test-OwnedTeardownFatalExit([int]$debuggerExit, [bool]$targetExited, [Nullable[int]]$targetExit, [string]$text, [int]$targetProcessId) {
    # A nonzero CDB exit is accepted only when it propagates this exact terminal target's fatal code.
    $code=if($debuggerExit -eq -1073740791){'c0000409'}elseif($debuggerExit -eq -1073741819){'c0000005'}else{''}
    return $targetExited -and $null -ne $targetExit -and $code -ne '' -and
        $targetExit -eq $debuggerExit -and (Test-TeardownExceptionCapture $text $targetProcessId $code)
}
function Test-OwnedFirstChanceAv([string]$text, [int]$targetProcessId, [string]$nonce) {
    $marker='VBAI_FIRSTCHANCE_AV_BEGIN '+$targetProcessId+' '+$nonce
    $match=[regex]::Match($text,'(?ms)^'+[regex]::Escape($marker)+'\r?\n(.*?)^VBAI_FIRSTCHANCE_AV_END\r?$')
    if(-not $match.Success){return $false}
    $event=[regex]::Match($match.Groups[1].Value,'(?im)^\s*Last event:\s*([0-9a-f]+)\.([0-9a-f]+):[^\r\n]*\bcode\s+c0000005\b[^\r\n]*first chance')
    return $event.Success -and [Convert]::ToInt64($event.Groups[1].Value,16) -eq $targetProcessId
}
function Test-OwnedStopBreakpoint([string]$text, [int]$targetProcessId, [string]$nonce) {
    $marker='VBAI_STOP_PROBE_BEGIN '+$nonce
    $match=[regex]::Match($text,'(?ms)^'+[regex]::Escape($marker)+'\r?\n(.*?)^VBAI_STOP_PROBE_END\r?$')
    if(-not $match.Success){return $false}
    $event=[regex]::Match($match.Groups[1].Value,'(?im)^\s*Last event:\s*([0-9a-f]+)\.([0-9a-f]+):[^\r\n]*\bcode\s+80000003\b')
    return $event.Success -and [Convert]::ToInt64($event.Groups[1].Value,16) -eq $targetProcessId
}
function Test-OwnedTeardownReady([string]$text, [int]$targetProcessId, [string]$nonce) {
    if($targetProcessId -le 0 -or $nonce -notmatch '^[a-f0-9]{32}$'){return $false}
    $line='VBAI_TEARDOWN_READY '+$targetProcessId+' '+$nonce
    return [regex]::IsMatch($text,'(?m)^'+[regex]::Escape($line)+'\r?$')
}
function Start-OwnedTeardownDebugger([string]$cdb, $target, [string]$directory, [string]$nonce) {
    $commands = Join-Path $directory 'teardown.commands.txt'
    $log = Join-Path $directory 'teardown.cdb.log'
    if ([IO.File]::Exists($commands) -or [IO.File]::Exists($log)) { throw 'Diagnostic outputs already exist; no attachment replay.' }
    $acp = [int][VBAiOwnedTeardownNative]::GetACP()
    $encoding = [Text.Encoding]::GetEncoding($acp, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
    [IO.File]::WriteAllBytes($commands, $encoding.GetBytes((Get-TeardownCommands $log $target.Id $nonce)))
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName=$cdb; $info.Arguments='-pd -p '+$target.Id+' -netsyms:no -cf "'+$commands+'"'
    $info.UseShellExecute=$false; $info.CreateNoWindow=$true
    $info.RedirectStandardInput=$true; $info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
    $debugger=New-Object Diagnostics.Process; $debugger.StartInfo=$info
    [void]$debugger.Start()
    $stdout=$debugger.StandardOutput.ReadToEndAsync(); $stderr=$debugger.StandardError.ReadToEndAsync()
    return @{ Process=$debugger; Stdout=$stdout; Stderr=$stderr; Log=$log; Commands=$commands }
}
function Wait-OwnedTeardownArmed($session,$target,[string]$nonce) {
    $watch=[Diagnostics.Stopwatch]::StartNew()
    while ($watch.Elapsed.TotalSeconds -lt 10 -and -not $session.Process.HasExited -and -not $target.HasExited) {
        $present=$false
        if ([IO.File]::Exists($session.Log) -and
            (Test-OwnedTeardownReady (Get-Content -LiteralPath $session.Log -Raw -Encoding Unicode) $target.Id $nonce) -and
            [VBAiOwnedTeardownNative]::CheckRemoteDebuggerPresent($target.Handle,[ref]$present) -and $present) { return }
        Start-Sleep -Milliseconds 100
    }
    throw 'Owned debugger failed to arm; no cleanup permission is emitted.'
}
function Stop-OwnedTeardownDebugger($session,$target) {
    if (-not $session.Process.HasExited) {
        if($session.StopAttempted){throw 'Bounded stop already attempted; retain without another break or detach.'}
        $session.StopAttempted=$true
        $nonce=[Guid]::NewGuid().ToString('N')
        $session.Process.StandardInput.WriteLine('.echo VBAI_STOP_PROBE_BEGIN '+$nonce+'; .lastevent; .echo VBAI_STOP_PROBE_END'); $session.Process.StandardInput.Flush()
        if (-not $target.HasExited -and -not [VBAiOwnedTeardownNative]::DebugBreakProcess($target.Handle)) {
            throw 'Exact debugger stop break failed; host/debugger are retained, no native cleanup replay.'
        }
        $watch=[Diagnostics.Stopwatch]::StartNew();$safeStop=$false
        while($watch.Elapsed.TotalSeconds -lt 5 -and -not $session.Process.HasExited){
            $text=if([IO.File]::Exists($session.Log)){Get-Content -LiteralPath $session.Log -Raw -Encoding Unicode}else{''}
            if(Test-OwnedStopBreakpoint $text $target.Id $nonce){$safeStop=$true;break}
            Start-Sleep -Milliseconds 100
        }
        if(-not $session.Process.HasExited){
            if(-not $safeStop){throw 'Stop event is not the exact owned breakpoint; retain pending fault without qd.'}
            $session.StopBreakpointVerified=$true
            $session.Process.StandardInput.WriteLine('qd'); $session.Process.StandardInput.Flush()
        }
        if (-not $session.Process.WaitForExit(5000)) { throw 'Owned debugger stop deadline; retained debugger/host, no termination.' }
    }
    if (-not $target.HasExited) {
        $present=$true
        if (-not [VBAiOwnedTeardownNative]::CheckRemoteDebuggerPresent($target.Handle,[ref]$present) -or $present) {
            throw 'Debugger detachment is not proven; preserve host without further actions.'
        }
    }
    $session.DebuggerExitCode=$session.Process.ExitCode
    $session.TargetExitObserved=$target.HasExited
    $session.TargetExitCode=if($target.HasExited){$target.ExitCode}else{$null}
    $session.FatalExitMatchedTarget=$false
    if ($session.DebuggerExitCode -ne 0) {
        $text=if([IO.File]::Exists($session.Log)){Get-Content -LiteralPath $session.Log -Raw -Encoding Unicode}else{''}
        $session.FatalExitMatchedTarget=Test-OwnedTeardownFatalExit $session.DebuggerExitCode $session.TargetExitObserved $session.TargetExitCode $text $target.Id
        if(-not $session.FatalExitMatchedTarget){throw ('Debugger returned unverified nonzero: '+$session.DebuggerExitCode)}
    }
}

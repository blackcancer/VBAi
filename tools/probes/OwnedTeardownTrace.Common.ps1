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
    return ".logopen /u `"$safeLog`"`n" +
        "sxe -c `".echo VBAI_TEARDOWN_EXCEPTION_BEGIN; .lastevent; .exr -1; .ecxr; kv; .echo VBAI_TEARDOWN_EXCEPTION_END; qd`" 0xc0000409`n" +
        ".echo VBAI_TEARDOWN_READY $targetProcessId $nonce`ng`n"
}
function Test-TeardownExceptionCapture([string]$text) {
    # Script command echoes are not executed exception-handler evidence.
    $match=[regex]::Match($text,'(?ms)^VBAI_TEARDOWN_EXCEPTION_BEGIN\r?\n(.*?)^VBAI_TEARDOWN_EXCEPTION_END\r?$')
    return $match.Success -and $match.Groups[1].Value -match '(?im)^\s*ExceptionCode:\s*c0000409\b' -and
        $match.Groups[1].Value -match '(?im)^\s*rip=[0-9a-f]+' -and
        $match.Groups[1].Value -match 'Child-SP\s+RetAddr'
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
            (Get-Content -LiteralPath $session.Log -Raw -Encoding Unicode).Contains("VBAI_TEARDOWN_READY $($target.Id) $nonce") -and
            [VBAiOwnedTeardownNative]::CheckRemoteDebuggerPresent($target.Handle,[ref]$present) -and $present) { return }
        Start-Sleep -Milliseconds 100
    }
    throw 'Owned debugger failed to arm; no cleanup permission is emitted.'
}
function Stop-OwnedTeardownDebugger($session,$target) {
    if (-not $session.Process.HasExited) {
        $session.Process.StandardInput.WriteLine('qd'); $session.Process.StandardInput.Flush()
        if (-not $target.HasExited -and -not [VBAiOwnedTeardownNative]::DebugBreakProcess($target.Handle)) {
            throw 'Exact debugger stop break failed; host/debugger are retained, no native cleanup replay.'
        }
        if (-not $session.Process.WaitForExit(5000)) { throw 'Owned debugger stop deadline; retained debugger/host, no termination.' }
    }
    if (-not $target.HasExited) {
        $present=$true
        if (-not [VBAiOwnedTeardownNative]::CheckRemoteDebuggerPresent($target.Handle,[ref]$present) -or $present) {
            throw 'Debugger detachment is not proven; preserve host without further actions.'
        }
    }
    if ($session.Process.ExitCode -ne 0) { throw ('Debugger returned nonzero: '+$session.Process.ExitCode) }
}

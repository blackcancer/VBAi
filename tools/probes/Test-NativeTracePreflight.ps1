[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $HelperPath,
    [Parameter(Mandatory)] [string] $OutputRoot,
    [string] $CdbPath = 'D:\Windows Kits\10\Debuggers\x64\cdb.exe'
)
$ErrorActionPreference = 'Stop'
if (-not [IO.Path]::IsPathRooted($HelperPath) -or -not [IO.Path]::IsPathRooted($OutputRoot)) { throw 'Helper/output paths must be absolute.' }
if ([IO.Path]::GetFileName($HelperPath) -ne 'VBAi.NativeTrace.Helper.exe') { throw 'Only the disposable non-Office helper can be launched.' }
if (-not [IO.File]::Exists($HelperPath) -or -not [IO.File]::Exists($CdbPath)) { throw 'Helper/CDB unavailable; no installation is performed.' }
$trial = Join-Path $OutputRoot ([Guid]::NewGuid().ToString('N'))
$child = Join-Path $trial ([Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($child)
$reportPath = Join-Path $trial 'preflight.json'
$jsProvider = Join-Path ([IO.Path]::GetDirectoryName($CdbPath)) 'winext/JsProvider.dll'
if (-not [IO.File]::Exists($jsProvider)) { throw 'Installed CDB JavaScript provider is unavailable; nothing will be installed.' }
$report = [ordered]@{
    State = 'STARTED'; Mode = 'INVASIVE_P_PD'; CdbPath = $CdbPath;
    CdbSha256 = (Get-FileHash -LiteralPath $CdbPath -Algorithm SHA256).Hash;
    CdbVersion = (Get-Item -LiteralPath $CdbPath).VersionInfo.FileVersion;
    JsProviderPath = $jsProvider; JsProviderSha256 = (Get-FileHash -LiteralPath $jsProvider -Algorithm SHA256).Hash;
    HelperPath = $HelperPath; HelperSha256 = (Get-FileHash -LiteralPath $HelperPath -Algorithm SHA256).Hash;
    Child = $child; ChildAttributesBefore = [IO.File]::GetAttributes($child).ToString();
    AttachmentObserved = $false; DebuggerDetachedVerified = $false; TargetStayedAlive = $false;
    TargetNormalShutdownVerified = $false; DetachOnExitAccepted = $false; SyntheticOpenRequests = 0;
    Scope = 'Only a newly created owned non-Office helper; no existing process attach or ACL/EFS/token/installation changes.'
}
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class VBAiPreflightLifecycle {
 [DllImport("kernel32.dll")] public static extern uint GetACP();
 [DllImport("kernel32.dll", SetLastError=true)] public static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
 [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr process);
 [DllImport("kernel32.dll", SetLastError=true)] public static extern bool DebugBreakProcess(IntPtr process);
 [DllImport("kernel32.dll", SetLastError=true)] public static extern bool CheckRemoteDebuggerPresent(IntPtr process, out bool present);
}
'@
function Start-OwnedProcess([string] $path, [string] $arguments) {
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $path; $info.Arguments = $arguments; $info.UseShellExecute = $false; $info.CreateNoWindow = $true
    $info.RedirectStandardInput = $true; $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process; $process.StartInfo = $info
    [void]$process.Start(); return $process
}
function Read-BoundedLine($process) {
    $task = $process.StandardOutput.ReadLineAsync()
    if (-not $task.Wait(10000)) { throw 'Owned helper response timeout; no process termination is attempted.' }
    return $task.Result
}
$helper = $null; $debugger = $null; $query = [IntPtr]::Zero; $failure = $null; $cleanup = New-Object Collections.Generic.List[string]
try {
    $helper = Start-OwnedProcess $HelperPath ('"' + $child + '"')
    $ready = Read-BoundedLine $helper
    if ($ready -notmatch '^READY (\d+) (.+)$' -or [int]$Matches[1] -ne $helper.Id) { throw 'Owned helper identity handshake failed.' }
    $report.TargetProcessId = $helper.Id; $report.TargetStartedUtc = $helper.StartTime.ToUniversalTime().ToString('o')
    if ([DateTime]::Parse($Matches[2]).ToUniversalTime() -ne $helper.StartTime.ToUniversalTime()) { throw 'Owned helper start-time mismatch.' }
    $query = [VBAiPreflightLifecycle]::OpenProcess(0x400, $false, $helper.Id)
    $present = $false
    if ($query -eq [IntPtr]::Zero -or -not [VBAiPreflightLifecycle]::CheckRemoteDebuggerPresent($query, [ref]$present) -or $present) { throw 'Owned helper debugger-state precondition failed.' }
    $log = Join-Path $trial 'trace.cdb.log'; $commandPath = Join-Path $trial 'trace.commands.txt'
    $scriptPath = Join-Path $PSScriptRoot 'NativeExportTrace.js'
    $report.TraceScriptSha256 = (Get-FileHash -LiteralPath $scriptPath -Algorithm SHA256).Hash
    $rootLiteral = ConvertTo-Json -InputObject $child -Compress
    # Quoted debugger meta-command paths interpret backslash escapes (\x64
    # becomes a hex character). Forward slashes preserve the native filename.
    $logCommand = $log.Replace('\', '/'); $providerCommand = $jsProvider.Replace('\', '/'); $scriptCommand = $scriptPath.Replace('\', '/')
    $content = ".logopen /u `"$logCommand`"`n.load `"$providerCommand`"`n.scriptproviders`n.scriptload `"$scriptCommand`"`n.scriptlist`ndx @`$scriptContents.configure($($helper.Id), $rootLiteral)`ng`n"
    $acp = [int][VBAiPreflightLifecycle]::GetACP()
    $encoding = [Text.Encoding]::GetEncoding($acp, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
    [IO.File]::WriteAllBytes($commandPath, $encoding.GetBytes($content))
    $report.CommandFileEncoding = "WindowsACP-$acp-noBOM"
    $arguments = '-pd -p ' + $helper.Id + ' -netsyms:no -cf "' + $commandPath + '"'
    $report.CdbArguments = $arguments
    $debugger = Start-OwnedProcess $CdbPath $arguments
    $report.DebuggerProcessId = $debugger.Id
    $stdout = $debugger.StandardOutput.ReadToEndAsync(); $stderr = $debugger.StandardError.ReadToEndAsync()
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($watch.Elapsed.TotalSeconds -lt 10 -and -not $debugger.HasExited) {
        $present = $false
        if ([VBAiPreflightLifecycle]::CheckRemoteDebuggerPresent($query, [ref]$present) -and $present) {
            $report.AttachmentObserved = $true; $report.DetachOnExitAccepted = $true
        }
        if ([IO.File]::Exists($log) -and (Get-Content -LiteralPath $log -Raw -Encoding Unicode) -match "VBAI_TRACE_READY pid=$($helper.Id)") { break }
        Start-Sleep -Milliseconds 100
    }
    if ($debugger.HasExited -or -not [IO.File]::Exists($log) -or (Get-Content -LiteralPath $log -Raw -Encoding Unicode) -notmatch "VBAI_TRACE_READY pid=$($helper.Id)") { throw 'CDB/JS trace did not arm; no synthetic file open will be issued.' }
    $present = $false
    if (-not [VBAiPreflightLifecycle]::CheckRemoteDebuggerPresent($query, [ref]$present) -or -not $present) { throw 'Debugger attachment was not observed.' }
    $report.AttachmentObserved = $true; $report.DetachOnExitAccepted = $true
    $report.SyntheticOpenRequests = 1
    $helper.StandardInput.WriteLine('OPEN'); $helper.StandardInput.Flush()
    $report.SyntheticResponse = Read-BoundedLine $helper
    if ($report.SyntheticResponse -ne 'OPEN_VERIFIED') { throw 'Synthetic one-open readback failed.' }
} catch { $failure = $_; $report.Failure = $_.Exception.ToString() }
finally {
    if ($debugger -ne $null -and -not $debugger.HasExited) {
        try {
            $debugger.StandardInput.WriteLine('bc *'); $debugger.StandardInput.WriteLine('dx @$scriptContents.summary()')
            $debugger.StandardInput.WriteLine('.logclose'); $debugger.StandardInput.WriteLine('qd'); $debugger.StandardInput.Flush()
            $break = [VBAiPreflightLifecycle]::OpenProcess(0x1FFFFF, $false, $helper.Id)
            try {
                if ($break -eq [IntPtr]::Zero -or -not [VBAiPreflightLifecycle]::DebugBreakProcess($break)) { throw 'Owned diagnostic stop break denied.' }
            } finally { if ($break -ne [IntPtr]::Zero) { [void][VBAiPreflightLifecycle]::CloseHandle($break) } }
        } catch { $cleanup.Add($_.Exception.ToString()) }
        if (-not $debugger.WaitForExit(5000)) {
            # Only this newly created debugger, never the helper/Office/other process.
            try { $report.DebuggerStopForced = $true; $debugger.Kill(); [void]$debugger.WaitForExit(5000) }
            catch { $cleanup.Add($_.Exception.ToString()) }
        }
    }
    if ($debugger -ne $null -and $debugger.HasExited) {
        $report.DebuggerExitCode = $debugger.ExitCode
        $text = $stdout.Result + $stderr.Result
        [IO.File]::WriteAllText((Join-Path $trial 'owned-helper-debugger-output.txt'), $text, [Text.Encoding]::UTF8)
        $report.DetachOnExitUnsupportedObserved = $text.IndexOf('does not support detach on exit', [StringComparison]::OrdinalIgnoreCase) -ge 0
    }
    if ($helper -ne $null -and -not $helper.HasExited -and $query -ne [IntPtr]::Zero) {
        $present = $false
        $report.DebuggerDetachedVerified = [VBAiPreflightLifecycle]::CheckRemoteDebuggerPresent($query, [ref]$present) -and -not $present
        $report.TargetStayedAlive = -not $helper.HasExited
        if ($report.DebuggerDetachedVerified) {
            try {
                $helper.StandardInput.WriteLine('STOP'); $helper.StandardInput.Flush()
                $report.ShutdownResponse = Read-BoundedLine $helper
                $report.TargetNormalShutdownVerified = $helper.WaitForExit(5000) -and $helper.ExitCode -eq 0 -and $report.ShutdownResponse -eq 'STOPPED'
                if ($helper.HasExited) { $report.TargetExitCode = $helper.ExitCode }
            } catch { $cleanup.Add($_.Exception.ToString()) }
        } else { $cleanup.Add('Detachment not proven: helper left untouched, no STOP or forced target termination.') }
    }
    if ($query -ne [IntPtr]::Zero) { [void][VBAiPreflightLifecycle]::CloseHandle($query) }
    $report.ScriptTraceVerified = $false
    try {
      $report.ChildAttributesAfter = [IO.File]::GetAttributes($child).ToString()
      if ([IO.File]::Exists((Join-Path $trial 'trace.cdb.log'))) {
        $records = @(Get-Content -LiteralPath (Join-Path $trial 'trace.cdb.log') -Encoding Unicode | Where-Object { $_.StartsWith('VBAI_FS ') } |
            ForEach-Object { $_.Substring('VBAI_FS '.Length) | ConvertFrom-Json })
        $entries = @($records | Where-Object { $_.Stage -eq 'ENTRY' -and $_.Call.Api -eq 'NtCreateFile' })
        $returns = @($records | Where-Object { $_.Stage -eq 'RETURN' -and $_.Api -eq 'NtCreateFile' })
        $summary = @($records | Where-Object { $_.Stage -eq 'SUMMARY' })
        $report.NativeTraceRecords = $records
        $report.ScriptTraceVerified = $entries.Count -eq 1 -and $returns.Count -eq 1 -and $summary.Count -eq 1 -and
            $entries[0].Call.Id -eq $returns[0].Id -and $entries[0].Call.ThreadId -eq $returns[0].ThreadId -and
            $entries[0].Call.ProcessId -eq $helper.Id -and $returns[0].ProcessId -eq $helper.Id -and
            $entries[0].Call.Path -eq ('\??\' + (Join-Path $child 'synthetic-owned.txt')) -and
            $returns[0].Path -eq $entries[0].Call.Path -and $returns[0].NtStatus -eq '0x00000000' -and
            $summary[0].Completed -eq 1 -and $summary[0].Pending -eq 0 -and $summary[0].ReadOrBreakpointErrors -eq 0
      }
    } catch { $report.EvidenceFailure = $_.Exception.ToString(); $cleanup.Add($_.Exception.ToString()) }
    $report.CleanupFailures = @($cleanup.ToArray())
    $report.State = if ($failure -eq $null -and $cleanup.Count -eq 0 -and $report.AttachmentObserved -and $report.DebuggerDetachedVerified -and
        $report.TargetStayedAlive -and $report.TargetNormalShutdownVerified -and $report.ScriptTraceVerified -and
        $report.DebuggerExitCode -eq 0 -and $report.DebuggerStopForced -ne $true -and $report.DetachOnExitUnsupportedObserved -ne $true -and
        $report.ChildAttributesBefore -eq $report.ChildAttributesAfter) { 'PASS' } else { 'FAIL' }
    $report.CompletedUtc = [DateTime]::UtcNow.ToString('o')
    try { [IO.File]::WriteAllText($reportPath, ($report | ConvertTo-Json -Depth 10), [Text.Encoding]::UTF8) }
    finally {
        if ($debugger -ne $null) { $debugger.Dispose() }; if ($helper -ne $null) { $helper.Dispose() }
    }
    Write-Output $reportPath
}
if ($report.State -ne 'PASS') { throw 'CDB owned-helper preflight failed; retain artifacts and do not attach to Office.' }

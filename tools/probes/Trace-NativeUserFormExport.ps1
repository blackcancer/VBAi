[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $PendingReport,
    [Parameter(Mandatory)] [Guid] $ExpectedMvid,
    [Parameter(Mandatory)] [ValidatePattern('^[A-Fa-f0-9]{64}$')] [string] $ExpectedAssemblySha256,
    [string] $CdbPath = 'D:\Windows Kits\10\Debuggers\x64\cdb.exe',
    [string] $DebuggerPreflightReport,
    [ValidateRange(10, 45)] [int] $MaxSeconds = 30,
    [switch] $Execute
)
$ErrorActionPreference = 'Stop'
if (-not [IO.Path]::IsPathRooted($PendingReport)) { throw 'Pending report must be absolute.' }
$pending = Get-Content -LiteralPath $PendingReport -Raw | ConvertFrom-Json
$directory = [IO.Path]::GetDirectoryName($PendingReport)
$targetPid = [int] $pending.ProcessId
$targetStart = [DateTime]::Parse($pending.ProcessStartedUtc).ToUniversalTime()
$targetPath = [IO.Path]::GetDirectoryName($pending.Destination).TrimEnd('\')
if ([IO.Path]::GetFileName($targetPath) -notmatch '^[a-fA-F0-9]{32}$') { throw 'Expected fresh GUID export directory.' }
if (-not [IO.Directory]::Exists($targetPath)) { throw 'Export child must already exist; no directory will be created.' }
if ([Guid]$pending.AssemblyMvid -ne $ExpectedMvid) { throw 'Pending candidate MVID mismatch.' }
if ((Get-FileHash -LiteralPath $pending.AssemblyPath -Algorithm SHA256).Hash -ne $ExpectedAssemblySha256) { throw 'Pending candidate assembly hash mismatch.' }
if (-not [IO.File]::Exists($CdbPath)) { throw 'CDB is unavailable; no tool will be installed.' }
$plan = [ordered]@{
    Mode = 'PREPARE_ONLY'; ProcessId = $targetPid; ProcessStartedUtc = $targetStart.ToString('o');
    Destination = $pending.Destination; ExpectedMvid = $ExpectedMvid.ToString(); ExpectedAssemblySha256 = $ExpectedAssemblySha256;
    DebuggerDetachOnExit = $true; MaxSeconds = $MaxSeconds;
    Scope = 'Exact owned Excel PID and absolute GUID child; native create/open/query-attributes arguments and paired NTSTATUS only. No exports, host launch/termination, installation, ACL/EFS/token/parent or trust mutation.'
}
if (-not $Execute) { $plan | ConvertTo-Json -Depth 8; return }

# A noninvasive -pvr/-pd refusal is not invasive-mode acceptance. Require root's
# measured preflight on a disposable non-Office process with this exact binary.
if ([string]::IsNullOrWhiteSpace($DebuggerPreflightReport) -or -not [IO.Path]::IsPathRooted($DebuggerPreflightReport)) {
    throw 'Execution requires an absolute measured CDB invasive -p/-pd preflight report; no attachment is attempted.'
}
$preflight = Get-Content -LiteralPath $DebuggerPreflightReport -Raw | ConvertFrom-Json
if ($preflight.CdbSha256 -ne (Get-FileHash -LiteralPath $CdbPath -Algorithm SHA256).Hash -or
    $preflight.Mode -ne 'INVASIVE_P_PD' -or $preflight.AttachmentObserved -ne $true -or
    $preflight.DebuggerDetachedVerified -ne $true -or $preflight.TargetStayedAlive -ne $true -or
    $preflight.TargetNormalShutdownVerified -ne $true -or $preflight.DetachOnExitAccepted -ne $true) {
    throw 'CDB detach-on-exit preflight is absent/incomplete/different; no host attachment is attempted.'
}

# Explicit execution only. Default invocation is a static plan and never attaches.
$target = [Diagnostics.Process]::GetProcessById($targetPid)
if ($target.ProcessName -ne 'EXCEL' -or $target.StartTime.ToUniversalTime() -ne $targetStart) { throw 'Exact Excel PID/start-time mismatch.' }
foreach ($name in 'trace.cdb.log', 'trace.lifecycle.json', 'trace.commands.txt', 'trace.armed') {
    if ([IO.File]::Exists((Join-Path $directory $name))) { throw 'Trace output already exists; use a new trial, never retry a native export.' }
}
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class VBAiNativeTraceLifecycle {
    [DllImport("kernel32.dll", SetLastError=true)] public static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
    [DllImport("kernel32.dll", SetLastError=true)] public static extern bool CloseHandle(IntPtr process);
    [DllImport("kernel32.dll", SetLastError=true)] public static extern bool DebugBreakProcess(IntPtr process);
    [DllImport("kernel32.dll", SetLastError=true)] public static extern bool CheckRemoteDebuggerPresent(IntPtr process, out bool present);
}
'@
$query = [VBAiNativeTraceLifecycle]::OpenProcess(0x400, $false, $targetPid)
if ($query -eq [IntPtr]::Zero) { throw 'Exact host debugger-state query denied; no elevation is attempted.' }
$present = $false
if (-not [VBAiNativeTraceLifecycle]::CheckRemoteDebuggerPresent($query, [ref]$present) -or $present) {
    [void][VBAiNativeTraceLifecycle]::CloseHandle($query)
    throw 'Host debugger state unavailable or another debugger is already attached.'
}
$log = Join-Path $directory 'trace.cdb.log'
$commands = Join-Path $directory 'trace.commands.txt'
$script = Join-Path $PSScriptRoot 'NativeExportTrace.js'
if ($script.Contains('"') -or $targetPath.Contains('"') -or $log.Contains('"')) { throw 'Unsupported quote in trace paths.' }
$rootLiteral = ConvertTo-Json -InputObject $targetPath -Compress
$content = @"
.logopen "$log"
.scriptload "$script"
dx @`$scriptContents.configure($targetPid, $rootLiteral)
g
"@
[IO.File]::WriteAllText($commands, $content, [Text.Encoding]::Unicode)
$info = New-Object Diagnostics.ProcessStartInfo
$info.FileName = $CdbPath
$info.Arguments = '-pd -p ' + $targetPid + ' -netsym:no -cf "' + $commands + '"'
$info.UseShellExecute = $false; $info.CreateNoWindow = $true
$info.RedirectStandardInput = $true; $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true
$debugger = New-Object Diagnostics.Process
$debugger.StartInfo = $info
$lifecycle = $plan
$lifecycle.Mode = 'EXECUTE'; $lifecycle.State = 'ATTACHING'; $lifecycle.DebuggerStopForced = $false
$lifecycle.AttachmentObserved = $false; $lifecycle.InvasiveDetachPreflightReport = $DebuggerPreflightReport
$started = $false; $primaryFailure = $null; $stopFailures = New-Object Collections.Generic.List[string]
try {
    [void]$debugger.Start(); $started = $true
    $lifecycle.DebuggerProcessId = $debugger.Id
    # Drain output without preserving unrelated startup module/register text.
    $stdout = $debugger.StandardOutput.ReadToEndAsync(); $stderr = $debugger.StandardError.ReadToEndAsync()
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($watch.Elapsed.TotalSeconds -lt 10 -and -not $debugger.HasExited) {
        if ([IO.File]::Exists($log) -and (Get-Content -LiteralPath $log -Raw) -match "VBAI_TRACE_READY pid=$targetPid") { break }
        Start-Sleep -Milliseconds 100
    }
    if ($debugger.HasExited -or -not [IO.File]::Exists($log) -or (Get-Content -LiteralPath $log -Raw) -notmatch "VBAI_TRACE_READY pid=$targetPid") {
        throw 'Trace was not armed; no export permission marker will be created.'
    }
    $present = $false
    if (-not [VBAiNativeTraceLifecycle]::CheckRemoteDebuggerPresent($query, [ref]$present) -or -not $present) {
        throw 'Debugger attachment was not independently observed; no export permission marker will be created.'
    }
    $lifecycle.AttachmentObserved = $true
    $lifecycle.State = 'ARMED'; $lifecycle.ArmedUtc = [DateTime]::UtcNow.ToString('o')
    [IO.File]::WriteAllText((Join-Path $directory 'trace.armed'), "$targetPid $($ExpectedMvid.ToString())", [Text.Encoding]::ASCII)
    while ($watch.Elapsed.TotalSeconds -lt $MaxSeconds -and -not $debugger.HasExited -and -not $target.HasExited -and
        -not [IO.File]::Exists((Join-Path $directory 'trace.export-complete'))) { Start-Sleep -Milliseconds 100 }
    $lifecycle.ExportCompletionObserved = [IO.File]::Exists((Join-Path $directory 'trace.export-complete'))
    $lifecycle.CaptureTimedOut = -not $lifecycle.ExportCompletionObserved
    if ($lifecycle.CaptureTimedOut) { throw 'Bounded diagnostic capture ended without export completion; this is not the original native export error.' }
} catch { $primaryFailure = $_; $lifecycle.Failure = $_.Exception.ToString() }
finally {
    $lifecycle.State = 'DETACHING'
    if ($started -and -not $debugger.HasExited) {
        try {
            # Queue cleanup first, then send a programmatic debugger stop to this
            # exact retained host only. This stop is explicitly diagnostic-induced.
            $debugger.StandardInput.WriteLine('bc *')
            $debugger.StandardInput.WriteLine('dx @$scriptContents.summary()')
            $debugger.StandardInput.WriteLine('.logclose')
            $debugger.StandardInput.WriteLine('qd')
            $debugger.StandardInput.Flush()
            if (-not $target.HasExited -and $target.StartTime.ToUniversalTime() -eq $targetStart) {
                $breakHandle = [VBAiNativeTraceLifecycle]::OpenProcess(0x1FFFFF, $false, $targetPid)
                try {
                    $lifecycle.DiagnosticStopBreakRequested = $true
                    if ($breakHandle -eq [IntPtr]::Zero -or -not [VBAiNativeTraceLifecycle]::DebugBreakProcess($breakHandle)) {
                        throw 'Diagnostic stop break denied; no elevation or host termination is allowed.'
                    }
                } finally { if ($breakHandle -ne [IntPtr]::Zero) { [void][VBAiNativeTraceLifecycle]::CloseHandle($breakHandle) } }
            }
        } catch { $stopFailures.Add($_.Exception.ToString()) }
        if (-not $debugger.WaitForExit(5000)) {
            # Only our newly created CDB is terminated. -pd provides detach-on-exit;
            # target liveness/debugger state below must still be verified independently.
            try { $lifecycle.DebuggerStopForced = $true; $debugger.Kill(); [void]$debugger.WaitForExit(5000) }
            catch { $stopFailures.Add($_.Exception.ToString()) }
        }
    }
    try {
        $present = $false
        $lifecycle.TargetExitedDuringTrace = $target.HasExited
        $lifecycle.DebuggerDetachedVerified = -not $target.HasExited -and
            [VBAiNativeTraceLifecycle]::CheckRemoteDebuggerPresent($query, [ref]$present) -and -not $present
        if (-not $lifecycle.DebuggerDetachedVerified) { $stopFailures.Add('Exact host liveness/detachment could not be verified; leave it untouched and do not retry export.') }
        if ($started -and $debugger.HasExited) {
            $lifecycle.DebuggerExitCode = $debugger.ExitCode
            $startupText = $stdout.Result + $stderr.Result
            $lifecycle.DetachOnExitUnsupportedObserved = $startupText.IndexOf('does not support detach on exit', [StringComparison]::OrdinalIgnoreCase) -ge 0
            # Only classify the known refusal; do not persist unrelated stdout,
            # startup registers or module information from the owned process.
        }
        $lifecycle.State = if ($stopFailures.Count -eq 0) { 'DETACHED' } else { 'DETACH_UNVERIFIED' }
        $lifecycle.StopFailures = @($stopFailures.ToArray())
        $lifecycle.CompletedUtc = [DateTime]::UtcNow.ToString('o')
        [IO.File]::WriteAllText((Join-Path $directory 'trace.lifecycle.json'), ($lifecycle | ConvertTo-Json -Depth 10), [Text.Encoding]::UTF8)
        if ($lifecycle.DebuggerDetachedVerified -and $stopFailures.Count -eq 0) {
            [IO.File]::WriteAllText((Join-Path $directory 'trace.detached'), "$targetPid $($ExpectedMvid.ToString())", [Text.Encoding]::ASCII)
        }
    } finally {
        [void][VBAiNativeTraceLifecycle]::CloseHandle($query); $debugger.Dispose(); $target.Dispose()
    }
}
if ($primaryFailure -or $stopFailures.Count -gt 0) { throw 'Trace failed or detach was not proven. See trace.lifecycle.json; no native failure attribution or automatic retry is allowed.' }

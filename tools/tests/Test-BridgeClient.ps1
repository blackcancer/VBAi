$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
public sealed class BridgeClientFixture : IDisposable {
    private readonly NamedPipeServerStream pipe;
    private readonly ManualResetEvent stop = new ManualResetEvent(false);
    private readonly Thread worker;
    public BridgeClientFixture(int id, string mode) {
        pipe = new NamedPipeServerStream("VBAi." + id, PipeDirection.InOut, 1, PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
        worker = new Thread(() => {
            try {
                pipe.WaitForConnection();
                var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, true);
                reader.ReadLine();
                if (mode == "reply") {
                    var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true);
                    writer.AutoFlush = true; writer.WriteLine("{\"Ok\":true}");
                    stop.WaitOne();
                } else if (mode == "silent") stop.WaitOne();
                else pipe.Dispose();
            } catch (IOException) {} catch (ObjectDisposedException) {}
        });
        worker.IsBackground = true; worker.Start();
    }
    public void Dispose() { stop.Set(); pipe.Dispose(); worker.Join(2000); stop.Dispose(); }
}
"@
$client = Join-Path $PSScriptRoot '../Invoke-VBAi.ps1'
$testId = $PID
foreach ($mode in @('reply','silent','close')) {
    $server = New-Object BridgeClientFixture($testId, $mode)
    try {
        $timer = [Diagnostics.Stopwatch]::StartNew()
        $message = $null; $reply = $null
        try { $reply = & $client -HostProcessId $testId -Command status -ResponseTimeoutSeconds 1 }
        catch { $message = $_.Exception.Message }
        if ($mode -eq 'reply' -and (($reply | ConvertFrom-Json).Ok -ne $true -or $message)) { throw 'Valid bridge reply failed.' }
        if ($mode -eq 'silent' -and (-not $message -or -not $message.Contains('may still be running') -or $timer.Elapsed.TotalSeconds -gt 5)) { throw "Unbounded or misleading timeout: $message" }
        if ($mode -eq 'close' -and (-not $message -or -not $message.Contains('without replying'))) { throw "Unexpected closed-pipe result: $message" }
        Write-Output "$mode PASS ($($timer.ElapsedMilliseconds) ms)"
    } finally { $server.Dispose() }
}

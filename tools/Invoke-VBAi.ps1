param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [ValidateRange(1, 300)] [int] $ResponseTimeoutSeconds = 30,
    [Parameter(Mandatory = $true, ParameterSetName = 'Json')] [string] $RequestJson,
    [Parameter(Mandatory = $true, ParameterSetName = 'Command')] [ValidateSet('status', 'list_projects', 'list_modules', 'read_module', 'debug_state', 'list_commands')] [string] $Command,
    [Parameter(ParameterSetName = 'Command')] [string] $Project,
    [Parameter(ParameterSetName = 'Command')] [string] $Module,
    [Parameter(ParameterSetName = 'Command')] [string] $Query
)

if ($PSCmdlet.ParameterSetName -eq 'Command') {
    $request = @{ Command = $Command }
    if ($Project) { $request.Project = $Project }
    if ($Module) { $request.Module = $Module }
    if ($Query) { $request.Query = $Query }
    $RequestJson = ConvertTo-Json -InputObject $request -Compress
}

if (-not ('VBAiPipePeer' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class VBAiPipePeer {
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool GetNamedPipeServerProcessId(IntPtr pipe, out uint processId);
}
"@
}

$pipe = [System.IO.Pipes.NamedPipeClientStream]::new('.', "VBAi.$HostProcessId", [System.IO.Pipes.PipeDirection]::InOut)
try {
    $pipe.Connect(3000)
    [uint32]$serverProcessId = 0
    if (-not [VBAiPipePeer]::GetNamedPipeServerProcessId($pipe.SafePipeHandle.DangerousGetHandle(), [ref]$serverProcessId)) {
        throw (New-Object ComponentModel.Win32Exception([Runtime.InteropServices.Marshal]::GetLastWin32Error()))
    }
    if ($serverProcessId -ne $HostProcessId) {
        throw "Pipe owner mismatch: requested host $HostProcessId, actual server $serverProcessId. No request was sent."
    }
    $writer = [System.IO.StreamWriter]::new($pipe, [System.Text.UTF8Encoding]::new($false), 4096, $true)
    $reader = [System.IO.StreamReader]::new($pipe, [System.Text.UTF8Encoding]::new($false), $false, 4096, $true)
    try {
        $writer.AutoFlush = $true
        $writer.WriteLine($RequestJson)
        $pendingRead = $reader.ReadLineAsync()
        if (-not $pendingRead.Wait([TimeSpan]::FromSeconds($ResponseTimeoutSeconds))) {
            throw "VBAi host $HostProcessId did not reply within $ResponseTimeoutSeconds seconds. The request may still be running; inspect the host before retrying."
        }
        $reply = $pendingRead.GetAwaiter().GetResult()
        if ($null -eq $reply) { throw "VBAi host $HostProcessId closed the pipe without replying." }
        $reply
    }
    finally {
        # A broken pipe during cleanup must not hide the original timeout/EOF.
        try { $reader.Dispose() } catch { }
        try { $writer.Dispose() } catch { }
    }
}
finally {
    $pipe.Dispose()
}

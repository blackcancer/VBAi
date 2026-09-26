param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
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

$pipe = [System.IO.Pipes.NamedPipeClientStream]::new('.', "CodexVBE.$HostProcessId", [System.IO.Pipes.PipeDirection]::InOut)
try {
    $pipe.Connect(3000)
    $writer = [System.IO.StreamWriter]::new($pipe, [System.Text.UTF8Encoding]::new($false), 4096, $true)
    $reader = [System.IO.StreamReader]::new($pipe, [System.Text.UTF8Encoding]::new($false), $false, 4096, $true)
    try {
        $writer.AutoFlush = $true
        $writer.WriteLine($RequestJson)
        $reader.ReadLine()
    }
    finally {
        $reader.Dispose()
        $writer.Dispose()
    }
}
finally {
    $pipe.Dispose()
}

# Loaded by the disposable native-theme probe. All execution goes through the
# add-in's native VBE Run Sub command; the host Application.Run API is not used.
function Invoke-RenderTraceProcedure($Context, [string]$Procedure, [string]$Cell) {
    $bridge = Join-Path $PSScriptRoot '../../Invoke-CodexVBE.ps1'
    $fields = @{ Command = 'read_module'; Project = $Context.Project; Module = $Context.Module }
    $read = & $bridge -HostProcessId $Context.HostProcessId -RequestJson ($fields | ConvertTo-Json -Compress) | ConvertFrom-Json
    if (-not $read.Ok) { throw ('Trace fixture read failed: ' + $read.Error) }
    $fields.Command = 'run_sub'
    $fields.Procedure = $Procedure
    $fields.ExpectedMode = 2
    $fields.ExpectedSha256 = $read.Data.Sha256
    $run = & $bridge -HostProcessId $Context.HostProcessId -RequestJson ($fields | ConvertTo-Json -Compress) | ConvertFrom-Json
    if (-not $run.Ok) { throw ('Trace fixture execution failed: ' + $run.Error) }
    $deadline = [Diagnostics.Stopwatch]::StartNew()
    do {
        $value = $Context.Workbook.Worksheets.Item(1).Range($Cell).Value2
        if ($null -ne $value) { return [int]$value }
        Start-Sleep -Milliseconds 100
    } while ($deadline.ElapsedMilliseconds -lt 5000)
    throw ('Native trace did not acknowledge ' + $Procedure)
}

function Start-VbeRenderTrace {
    param($Excel, $Workbook, [IntPtr]$Editor, [int]$HostProcessId,
          [string]$LibraryPath, [string]$OutputDirectory, [IntPtr]$SourceWindow = [IntPtr]::Zero, [switch]$PatternPilot)
    $library = (Resolve-Path -LiteralPath $LibraryPath -ErrorAction Stop).Path
    $tracePath = [IO.Path]::GetFullPath((Join-Path $OutputDirectory 'native-render.jsonl'))
    if ($library.Contains('"') -or $tracePath.Contains('"')) { throw 'Invalid quote in trace path.' }
    $component = $Workbook.VBProject.VBComponents.Add(1)
    $component.Name = 'NativeRenderTraceFixture'
    $startDeclaration = 'Private Declare PtrSafe Function TraceStart Lib "' + $library + '" Alias "CodexVbeTraceStart" (ByVal editor As LongPtr, ByVal path As LongPtr, ByVal limit As Long) As Long'
    $startArguments = 'CLngPtr(' + $Editor.ToInt64() + '), StrPtr(outputPath), 100000'
    if ($SourceWindow -ne [IntPtr]::Zero) {
        $startDeclaration = 'Private Declare PtrSafe Function TraceStart Lib "' + $library + '" Alias "CodexVbeTraceStartForWindow" (ByVal editor As LongPtr, ByVal source As LongPtr, ByVal path As LongPtr, ByVal limit As Long) As Long'
        $startArguments = 'CLngPtr(' + $Editor.ToInt64() + '), CLngPtr(' + $SourceWindow.ToInt64() + '), StrPtr(outputPath), 100000'
    }
    if ($PatternPilot) {
        if ($SourceWindow -eq [IntPtr]::Zero) { throw 'The pattern pilot requires one explicit toolbar window.' }
        $startDeclaration = $startDeclaration.Replace('CodexVbeTraceStartForWindow', 'CodexVbeToolbarPatternStart')
    }
    $component.CodeModule.AddFromString(@"
Option Explicit
$startDeclaration
Private Declare PtrSafe Function TraceStop Lib "$library" Alias "CodexVbeTraceStop" () As Long
Public Sub StartNativeRenderTrace()
    Dim outputPath As String
    outputPath = "$tracePath"
    ThisWorkbook.Worksheets(1).Range("Z1").Value2 = TraceStart($startArguments)
End Sub
Public Sub StopNativeRenderTrace()
    ThisWorkbook.Worksheets(1).Range("Z2").Value2 = TraceStop()
End Sub
"@)
    $context = [pscustomobject]@{
        HostProcessId = $HostProcessId; Project = [string]$Workbook.VBProject.Name
        Module = [string]$component.Name; Workbook = $Workbook
        OutputDirectory = $OutputDirectory; TracePath = $tracePath; Stopped = $false
    }
    try {
        $status = Invoke-RenderTraceProcedure $context 'StartNativeRenderTrace' 'Z1'
        if ($status -ne 0) { throw "Native trace start failed: Win32 status $status" }
        @{ Started = $true; Library = $library; LibrarySha256 = (Get-FileHash -LiteralPath $library).Hash;
           ProcessId = $HostProcessId; Editor = $Editor.ToInt64(); SourceWindow = $SourceWindow.ToInt64(); PatternPilot = [bool]$PatternPilot; EventLimit = 100000 } |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutputDirectory 'native-render-start.json') -Encoding UTF8
        return $context
    } catch {
        # Start failure must not leave an unknown active trace until teardown.
        try { Stop-VbeRenderTrace -Context $context } catch { Write-Warning $_.Exception.Message }
        throw
    } finally { $component = $null }
}

function Stop-VbeRenderTrace {
    param($Context)
    if ($Context.Stopped) { return }
    $status = Invoke-RenderTraceProcedure $Context 'StopNativeRenderTrace' 'Z2'
    if ($status -ne 0) { throw "Native trace stop failed: Win32 status $status" }
    $Context.Stopped = $true
    if (-not (Test-Path -LiteralPath $Context.TracePath)) { throw 'The stopped trace produced no log.' }
    @{ Stopped = $true; Status = $status; TracePath = $Context.TracePath } |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $Context.OutputDirectory 'native-render-stop.json') -Encoding UTF8
    $Context.Workbook = $null
}

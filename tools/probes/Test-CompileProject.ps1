param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') { throw 'An isolated Excel process is required.' }

function Invoke-Vbe([hashtable] $Request) {
    $payload = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $response = & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $payload |
        ConvertFrom-Json
    if (-not $response.Ok) { throw "$($Request.Command): $($response.Error)" }
    return $response.Data
}

$project = 'VBAProject'
$module = 'ThisWorkbook'
$initial = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
if ($initial.Code.Length -ne 0) { throw 'Use a fresh disposable workbook with an empty ThisWorkbook module.' }

$invalid = "Option Explicit`r`nPublic Sub CodexCompileProbe()`r`n    missingValue = 1`r`nEnd Sub"
Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
    ExpectedSha256 = $initial.Sha256; StartLine = 1; Count = 0; Text = $invalid } | Out-Null
$failure = Invoke-Vbe @{ Command = 'compile_project'; Project = $project; ExpectedMode = 2 }
$location = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
if ($failure.Compiled -or $failure.Verification -ne 'NativeDiagnosticCaptured' -or
    $failure.Diagnostic -notmatch 'Variable non définie|Variable not defined' -or
    $location.ActiveModule -ne $module -or $location.Selection.StartLine -ne 3) {
    throw 'The expected native compile diagnostic or source selection was not captured.'
}

$current = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
$corrected = "Option Explicit`r`nPublic Sub CodexCompileProbe()`r`n    Dim definedValue As Long`r`n    definedValue = 1`r`nEnd Sub"
$lineCount = @($current.Code -split '\r?\n').Count
Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
    ExpectedSha256 = $current.Sha256; StartLine = 1; Count = $lineCount; Text = $corrected } | Out-Null
$success = Invoke-Vbe @{ Command = 'compile_project'; Project = $project; ExpectedMode = 2 }
if (-not $success.Compiled -or $success.Diagnostic -or
    $success.Verification -ne 'NoNativeDiagnosticObserved') {
    throw 'The corrected VBA code did not compile without a native diagnostic.'
}

[pscustomobject]@{
    HostProcessId = $HostProcessId
    ErrorDiagnostic = $failure.Diagnostic
    ErrorLine = $location.Selection.StartLine
    CorrectedCompiled = $success.Compiled
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

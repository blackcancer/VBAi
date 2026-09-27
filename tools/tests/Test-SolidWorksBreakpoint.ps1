param([Parameter(Mandatory = $true)] [int] $HostProcessId,
    [switch] $InspectNativePanes)

$ErrorActionPreference = 'Stop'
if ((Get-Process -Id $HostProcessId -ErrorAction Stop).ProcessName -ne 'SLDWORKS') {
    throw 'An already-open SOLIDWORKS process is required.'
}
function Invoke-Vbe([hashtable] $Request) {
    $payload = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $response = & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $payload |
        ConvertFrom-Json
    if (-not $response.Ok) { throw "$($Request.Command): $($response.Error)" }
    return $response.Data
}
function Get-CommandById([int] $Id) {
    return @(Invoke-Vbe @{ Command = 'list_commands'; Query = '' }) |
        Where-Object { $_.Id -eq $Id -and $_.Enabled } | Select-Object -First 1
}

$project = 'test'
$projectFile = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\test.swp'))
$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if (@($projects | Where-Object { $_.Name -eq $project -and $_.FileName -eq $projectFile -and $_.Mode -eq 2 }).Count -ne 1) {
    throw 'The disposable SOLIDWORKS VBA project is not uniquely available in design mode.'
}
$module = 'CodexSwBreakpointProbe'
$baseline = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
if (@($baseline | Where-Object { $_.Name -eq $module }).Count -ne 0) {
    throw 'The disposable module name already exists.'
}
$original = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
$source = $null
$outputPath = Join-Path $env:TEMP ("CodexVBE-sw-breakpoint-{0}.txt" -f $HostProcessId)
if (Test-Path -LiteralPath $outputPath) { throw 'The disposable output path already exists.' }
$breakpointSet = $false
$breakpointRemovalInvoked = $false
$breakLine = 5
$stopped = $false
$stepped = $false
$continued = $false
$localsProbe = $null
try {
    $created = Invoke-Vbe @{ Command = 'create_module'; Project = $project; Module = $module; ExpectedMode = 2 }
    $code = "Option Explicit`r`nPublic Sub CodexSwBreakpointSmoke()`r`n    Dim probeValue As Long`r`n    probeValue = 1`r`n    probeValue = probeValue + 1`r`n    Open `"$outputPath`" For Output As #1`r`n    Print #1, CStr(probeValue)`r`n    Close #1`r`nEnd Sub"
    Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
        ExpectedSha256 = $created.Sha256; StartLine = 1; Count = $created.Lines; Text = $code } | Out-Null
    $settled = $false
    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        $source = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
        Start-Sleep -Milliseconds 100
        $stable = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
        if ($stable.Sha256 -eq $source.Sha256) { $settled = $true; break }
        $source = $stable
    }
    if (-not $settled) { throw 'The temporary module did not settle after editing.' }
    Invoke-Vbe @{ Command = 'select_code'; Project = $project; Module = $module;
        ExpectedSha256 = $source.Sha256; StartLine = $breakLine } | Out-Null
    $breakCommand = Get-CommandById 51
    if (-not $breakCommand) { throw 'The native Toggle Breakpoint command is unavailable.' }
    Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
        ExpectedSha256 = $source.Sha256; StartLine = $breakLine; ExpectedMode = 2;
        Action = 'toggle_breakpoint'; ControlId = 51; ControlCaption = $breakCommand.Caption } | Out-Null
    $breakpointSet = $true
    $source = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
    Invoke-Vbe @{ Command = 'select_code'; Project = $project; Module = $module;
        ExpectedSha256 = $source.Sha256; StartLine = 2 } | Out-Null
    $runCommand = Get-CommandById 186
    if (-not $runCommand) { throw 'The native Run Sub command is unavailable.' }
    Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
        ExpectedSha256 = $source.Sha256; StartLine = 2; ExpectedMode = 2;
        Action = 'run'; ControlId = 186; ControlCaption = $runCommand.Caption } | Out-Null
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
        if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq $breakLine) { $stopped = $true; break }
        Start-Sleep -Milliseconds 100
    }
    if (-not $stopped) { throw "The VBA project did not stop at line $breakLine (mode=$($state.Mode))." }
    Invoke-Vbe @{ Command = 'open_debug_pane'; Action = 'locals' } | Out-Null
    if ($InspectNativePanes) {
        Invoke-Vbe @{ Command = 'open_debug_pane'; Action = 'immediate' } | Out-Null
        & (Join-Path $PSScriptRoot '..\probes\Inspect-VbePaneNative.ps1') -HostProcessId $HostProcessId
        & (Join-Path $PSScriptRoot '..\probes\Inspect-VbeMsaaChildren.ps1') -HostProcessId $HostProcessId
    }
    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        $windows = Invoke-Vbe @{ Command = 'debug_windows' }
        $localsProbe = $windows.Locals
        if (@($localsProbe.Items).Count -gt 0) { break }
        Start-Sleep -Milliseconds 100
    }
    $stepCommand = Get-CommandById 194
    if (-not $stepCommand) { throw 'The native Step Over command is unavailable in break mode.' }
    Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
        ExpectedSha256 = $source.Sha256; StartLine = $breakLine; ExpectedMode = 1;
        Action = 'step_over'; ControlId = 194; ControlCaption = $stepCommand.Caption } | Out-Null
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
        if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq ($breakLine + 1)) { $stepped = $true; break }
        Start-Sleep -Milliseconds 100
    }
    if (-not $stepped) { throw "Step Over did not reach line $($breakLine + 1) (mode=$($state.Mode))." }
    $continueCommand = Get-CommandById 186
    if (-not $continueCommand) { throw 'The native Continue command is unavailable in break mode.' }
    Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
        ExpectedSha256 = $source.Sha256; StartLine = ($breakLine + 1); ExpectedMode = 1;
        Action = 'continue'; ControlId = 186; ControlCaption = $continueCommand.Caption } | Out-Null
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
        if ($state.Mode -eq 2 -and (Test-Path -LiteralPath $outputPath) -and
            (Get-Content -LiteralPath $outputPath -Raw).Trim() -eq '2') {
            $continued = $true
            break
        }
        Start-Sleep -Milliseconds 100
    }
    if (-not $continued) { throw "Continue did not finish with expected value 2 (mode=$($state.Mode))." }
}
finally {
    $cleanupError = $null
    $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
    if ($state.Mode -eq 1) {
        Invoke-Vbe @{ Command = 'debug_global'; Project = $project; ExpectedMode = 1; Action = 'reset' } | Out-Null
        for ($attempt = 0; $attempt -lt 20; $attempt++) {
            $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
            if ($state.Mode -eq 2) { break }
            Start-Sleep -Milliseconds 100
        }
    }
    if ($state.Mode -ne 2) { throw 'The SOLIDWORKS VBE did not return to design mode.' }
    if ($breakpointSet -and $source) {
        try {
            $source = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
            Invoke-Vbe @{ Command = 'select_code'; Project = $project; Module = $module;
                ExpectedSha256 = $source.Sha256; StartLine = $breakLine } | Out-Null
            $breakCommand = Get-CommandById 51
            if (-not $breakCommand) { throw 'Cannot remove the disposable breakpoint.' }
            Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
                ExpectedSha256 = $source.Sha256; StartLine = $breakLine; ExpectedMode = 2;
                Action = 'toggle_breakpoint'; ControlId = 51; ControlCaption = $breakCommand.Caption } | Out-Null
            $breakpointRemovalInvoked = $true
        }
        catch { $cleanupError = "Breakpoint removal was not verified: $_" }
    }
    $current = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
    if (@($current | Where-Object { $_.Name -eq $module }).Count -eq 1) {
        $projectState = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
        $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $module }
        Invoke-Vbe @{ Command = 'remove_component'; Project = $project; Module = $module;
            ExpectedProjectVersion = $projectState.Version; ExpectedComponentVersion = $component.Version } | Out-Null
    }
    $last = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
    $after = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
    if ($last.Count -ne $baseline.Count -or $after.Sha256 -ne $original.Sha256) {
        throw 'The SOLIDWORKS project did not return to its initial code/component inventory.'
    }
    if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }
    if ($cleanupError) { throw $cleanupError }
}
[pscustomobject]@{ HostProcessId = $HostProcessId; BreakpointHit = $stopped;
    StepOverVerified = $stepped; ContinueVerified = $continued;
    BreakpointRemovalInvoked = $breakpointRemovalInvoked;
    LocalsAvailable = $localsProbe.Available; LocalsError = $localsProbe.Error;
    LocalsRows = @($localsProbe.Items).Count; LocalsSnapshot = ($localsProbe | ConvertTo-Json -Compress -Depth 3);
    OriginalModuleShaPreserved = $true; ModuleInventoryRestored = $true } | Format-List

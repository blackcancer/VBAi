param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
if ((Get-Process -Id $HostProcessId -ErrorAction Stop).ProcessName -ne 'SLDWORKS') {
    throw 'An already-open SOLIDWORKS process is required.'
}

function Invoke-Vbe([hashtable] $Request) {
    $payload = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $response = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $payload |
        ConvertFrom-Json
    if (-not $response.Ok) { throw "$($Request.Command): $($response.Error)" }
    return $response.Data
}

$project = 'test'
$projectFile = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\test.swp'))
$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if (@($projects | Where-Object { $_.Name -eq $project -and $_.FileName -eq $projectFile -and $_.Mode -eq 2 }).Count -ne 1) {
    throw 'The expected disposable SOLIDWORKS project is not uniquely available in design mode.'
}
$module = 'CodexSwDebugProbe'
$baseline = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
if (@($baseline | Where-Object { $_.Name -eq $module }).Count -ne 0) {
    throw 'The disposable module name already exists.'
}
$original = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
$outputPath = Join-Path $env:TEMP ("VBAi-sw-debug-{0}.txt" -f $HostProcessId)
if (Test-Path -LiteralPath $outputPath) { throw 'The disposable output path already exists.' }
$verified = $false
try {
    $created = Invoke-Vbe @{ Command = 'create_module'; Project = $project; Module = $module; ExpectedMode = 2 }
    $code = "Option Explicit`r`nPublic Sub CodexSwDebugSmoke()`r`n    Open `"$outputPath`" For Output As #1`r`n    Print #1, `"CodexSwDebugSmoke:42`"`r`n    Close #1`r`nEnd Sub"
    Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
        ExpectedSha256 = $created.Sha256; StartLine = 1; Count = $created.Lines; Text = $code } | Out-Null
    $source = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
    $compiled = Invoke-Vbe @{ Command = 'compile_project'; Project = $project; ExpectedMode = 2 }
    $diagnostic = Invoke-Vbe @{ Command = 'debug_dialog' }
    if ($diagnostic.Visible) { throw 'Compilation raised a native diagnostic dialog.' }
    $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
    if ($state.Mode -ne 2) { throw 'Compilation did not leave the project in design mode.' }
    Invoke-Vbe @{ Command = 'open_debug_pane'; Action = 'immediate' } | Out-Null
    Invoke-Vbe @{ Command = 'select_code'; Project = $project; Module = $module;
        ExpectedSha256 = $source.Sha256; StartLine = 2 } | Out-Null
    $runCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = '' }) |
        Where-Object { $_.Id -eq 186 -and $_.Enabled } | Select-Object -First 1
    if (-not $runCommand) { throw 'The native Run Sub command is unavailable.' }
    $run = Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
        ExpectedSha256 = $source.Sha256; StartLine = 2; ExpectedMode = 2;
        Action = 'run'; ControlId = 186; ControlCaption = $runCommand.Caption }
    $observed = $false
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
        if ($state.Mode -eq 2 -and (Test-Path -LiteralPath $outputPath) -and
            (Get-Content -LiteralPath $outputPath -Raw) -match 'CodexSwDebugSmoke:42') {
            $observed = $true
            break
        }
        Start-Sleep -Milliseconds 100
    }
    if (-not $observed) { throw "The native run did not produce the expected output in design mode (mode=$($state.Mode))." }
    $verified = $true
}
finally {
    try {
        $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
        if ($state.Mode -eq 1) {
            Invoke-Vbe @{ Command = 'debug_global'; Project = $project; ExpectedMode = 1; Action = 'reset' } | Out-Null
        }
    }
    catch { Write-Warning "Could not verify/reset the debugger: $_" }
    $current = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
    if (@($current | Where-Object { $_.Name -eq $module }).Count -eq 1) {
        $projectState = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
        $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $module }
        Invoke-Vbe @{ Command = 'remove_component'; Project = $project; Module = $module;
            ExpectedProjectVersion = $projectState.Version; ExpectedComponentVersion = $component.Version } | Out-Null
    }
    $last = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
    $after = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
    if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }
    if ($last.Count -ne $baseline.Count -or $after.Sha256 -ne $original.Sha256) {
        throw 'The SOLIDWORKS test project did not return to its original component/code inventory.'
    }
}
[pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project;
    CompileAndRunVerified = $verified; OriginalModuleShaPreserved = $true;
    ModuleInventoryRestored = $true } | Format-List

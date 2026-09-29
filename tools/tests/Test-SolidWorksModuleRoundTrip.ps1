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
$target = @($projects | Where-Object { $_.Name -eq $project -and $_.FileName -eq $projectFile -and $_.Mode -eq 2 })
if ($target.Count -ne 1) {
    throw "The expected SOLIDWORKS project matched $($target.Count) entries. Name=$($projects[-1].Name -eq $project) Path=$($projects[-1].FileName -eq $projectFile) Mode=$($projects[-1].Mode -eq 2); expected=$projectFile; actual=$($projects[-1].FileName)."
}
$baseline = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
$module = 'CodexSwRoundTrip'
if (@($baseline | Where-Object { $_.Name -eq $module }).Count -ne 0) {
    throw 'The disposable module name already exists.'
}
$existing = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
$path = Join-Path $env:TEMP ("VBAi-sw-module-{0}.bas" -f $HostProcessId)
if (Test-Path -LiteralPath $path) { throw 'The disposable export path already exists.' }
$roundTrip = $false
$cleanupVerified = $false
try {
    $created = Invoke-Vbe @{ Command = 'create_module'; Project = $project;
        Module = $module; ExpectedMode = 2 }
    $code = "Option Explicit`r`nPublic Sub CodexSwSmoke()`r`n    Debug.Print 42`r`nEnd Sub"
    Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
        ExpectedSha256 = $created.Sha256; StartLine = 1; Count = $created.Lines; Text = $code } | Out-Null
    $before = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
    $procedure = Invoke-Vbe @{ Command = 'list_procedures'; Project = $project; Module = $module }
    if (@($procedure.Procedures | Where-Object { $_.Name -eq 'CodexSwSmoke' }).Count -ne 1) {
        throw 'The SOLIDWORKS VBE did not recognize the test procedure.'
    }
    $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $module }
    $exported = Invoke-Vbe @{ Command = 'export_component'; Project = $project;
        Module = $module; Path = $path; ExpectedComponentVersion = $component.Version }
    if (-not (Test-Path -LiteralPath $path) -or $exported.Bytes -lt 1) {
        throw 'The module export file is missing or empty.'
    }
    $projectState = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
    $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $module }
    $removed = Invoke-Vbe @{ Command = 'remove_component'; Project = $project; Module = $module;
        ExpectedProjectVersion = $projectState.Version; ExpectedComponentVersion = $component.Version }
    $imported = Invoke-Vbe @{ Command = 'import_component'; Project = $project;
        Path = $path; ExpectedProjectVersion = $removed.Version }
    if (-not $imported.Applied -or $imported.ImportedName -ne $module) {
        throw 'The module import was not confirmed.'
    }
    $after = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
    if ($after.Sha256 -ne $before.Sha256) { throw 'The imported VBA code SHA differs.' }
    $roundTrip = $true
}
finally {
    $current = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
    if (@($current | Where-Object { $_.Name -eq $module }).Count -eq 1) {
        $projectState = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
        $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $module }
        Invoke-Vbe @{ Command = 'remove_component'; Project = $project; Module = $module;
            ExpectedProjectVersion = $projectState.Version; ExpectedComponentVersion = $component.Version } | Out-Null
    }
    $last = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
    $original = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
    $cleanupVerified = @($last | Where-Object { $_.Name -eq $module }).Count -eq 0 -and
        $original.Sha256 -eq $existing.Sha256 -and $last.Count -eq $baseline.Count
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    if (-not $cleanupVerified) { throw 'The SOLIDWORKS test project did not return to its original component/code inventory.' }
}
[pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project;
    RoundTripVerified = $roundTrip; OriginalModuleShaPreserved = $cleanupVerified;
    ModuleInventoryRestored = $cleanupVerified } | Format-List

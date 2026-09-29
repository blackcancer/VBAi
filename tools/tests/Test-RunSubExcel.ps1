param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$hosts = @(Get-Process EXCEL -ErrorAction Stop)
if ($hosts.Count -ne 1 -or $hosts[0].Id -ne $HostProcessId) {
    throw 'Exactly one isolated Excel process is required.'
}
function Invoke-Vbe([hashtable] $Request) {
    $payload = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $response = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $payload |
        ConvertFrom-Json
    if (-not $response.Ok) { throw "$($Request.Command): $($response.Error)" }
    return $response.Data
}
$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2 -or $projects[0].FileName) {
    throw 'A single unsaved disposable Excel VBA project is required.'
}
$project = [string]$projects[0].Name
$module = 'CodexRunSubProbe'
$outputPath = Join-Path $env:TEMP ("VBAi-run-sub-{0}.txt" -f $HostProcessId)
if (Test-Path -LiteralPath $outputPath) { throw 'The disposable output path already exists.' }
try {
    $created = Invoke-Vbe @{ Command = 'create_module'; Project = $project; Module = $module; ExpectedMode = 2 }
    $code = "Public Sub CodexRunSubSmoke()`r`n    Open `"$outputPath`" For Output As #1`r`n    Print #1, `"CodexRunSubSmoke:42`"`r`n    Close #1`r`nEnd Sub"
    Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
        ExpectedSha256 = $created.Sha256; StartLine = 1; Count = $created.Lines; Text = $code } | Out-Null
    $source = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
    $run = Invoke-Vbe @{ Command = 'run_sub'; Project = $project; Module = $module;
        Procedure = 'CodexRunSubSmoke'; ExpectedSha256 = $source.Sha256; ExpectedMode = 2 }
    $verified = $false
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
        if ($state.Mode -eq 2 -and (Test-Path -LiteralPath $outputPath) -and
            (Get-Content -LiteralPath $outputPath -Raw) -match 'CodexRunSubSmoke:42') {
            $verified = $true
            break
        }
        Start-Sleep -Milliseconds 100
    }
    if (-not $verified) { throw 'run_sub did not produce its expected side effect in design mode.' }
    [pscustomobject]@{ HostProcessId = $HostProcessId; RunSubVerified = $true;
        Procedure = 'CodexRunSubSmoke'; Control = $run.Control } | Format-List
}
finally {
    $current = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
    if (@($current | Where-Object { $_.Name -eq $module }).Count -eq 1) {
        $projectState = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
        $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $module }
        Invoke-Vbe @{ Command = 'remove_component'; Project = $project; Module = $module;
            ExpectedProjectVersion = $projectState.Version; ExpectedComponentVersion = $component.Version } | Out-Null
    }
    if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }
}

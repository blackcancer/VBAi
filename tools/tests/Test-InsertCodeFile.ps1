param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
if ((Get-Process -Id $HostProcessId -ErrorAction Stop).ProcessName -ne 'EXCEL') {
    throw 'An isolated Excel process is required.'
}

function Invoke-VbeRaw([hashtable] $Request) {
    $payload = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    return (& (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $payload |
        ConvertFrom-Json)
}

function Invoke-Vbe([hashtable] $Request) {
    $response = Invoke-VbeRaw $Request
    if (-not $response.Ok) { throw "$($Request.Command): $($response.Error)" }
    return $response.Data
}

$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2 -or $projects[0].FileName) {
    throw 'Use one unsaved disposable design-mode VBA project.'
}
$project = [string]$projects[0].Name
$path = Join-Path $env:TEMP ("VBAi-code-insert-{0}.bas" -f $HostProcessId)
$ansiPath = Join-Path $env:TEMP ("VBAi-code-insert-ansi-{0}.bas" -f $HostProcessId)
if ((Test-Path -LiteralPath $path) -or (Test-Path -LiteralPath $ansiPath)) {
    throw 'A disposable source path already exists.'
}
try {
    $accent = [char]0x00E9
    $source = "' UTF-8 $($accent)preuve`r`nPublic Sub InsertedFromFile()`r`n    Debug.Print 42`r`nEnd Sub"
    [IO.File]::WriteAllText($path, $source, [Text.UTF8Encoding]::new($false))
    $sourceHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
    $inspection = Invoke-Vbe @{ Command = 'inspect_code_file'; Path = $path }
    if (-not $inspection.StrictUtf8Valid -or -not $inspection.ExplicitEncodingRequired -or
        $inspection.Sha256 -ne $sourceHash -or $inspection.ContentIncluded) {
        throw 'UTF-8 source inspection did not report the ambiguous BOM-less encoding.'
    }
    $module = 'CodexFileInsertProbe'
    $created = Invoke-Vbe @{ Command = 'create_module'; Project = $project; Module = $module; ExpectedMode = 2 }
    $before = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
    $ambiguous = Invoke-VbeRaw @{ Command = 'insert_code_file'; Project = $project; Module = $module;
        Path = $path; StartLine = ($created.Lines + 1); ExpectedSha256 = $before.Sha256 }
    if ($ambiguous.Ok -or $ambiguous.Error -notmatch 'ambiguous') {
        throw 'An ambiguous BOM-less UTF-8 file was accepted without SourceEncoding.'
    }
    $result = Invoke-Vbe @{ Command = 'insert_code_file'; Project = $project; Module = $module;
        Path = $path; StartLine = ($created.Lines + 1); ExpectedSha256 = $before.Sha256;
        SourceEncoding = 'utf-8' }
    $after = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
    $procedures = Invoke-Vbe @{ Command = 'list_procedures'; Project = $project; Module = $module }
    if ($result.SourceSha256 -ne $sourceHash -or $result.SourceEncoding -ne 'utf-8' -or
        $result.Sha256 -ne $after.Sha256 -or $result.InsertedLineCount -lt 4 -or
        $after.Code -notmatch 'Sub InsertedFromFile\(' -or
        -not $after.Code.Contains("UTF-8 $($accent)preuve") -or
        @($procedures.Procedures | Where-Object { $_.Name -eq 'InsertedFromFile' }).Count -ne 1) {
        throw 'Inserted source or provenance could not be read back.'
    }
    $stale = Invoke-VbeRaw @{ Command = 'insert_code_file'; Project = $project; Module = $module;
        Path = $path; StartLine = 1; ExpectedSha256 = $before.Sha256; SourceEncoding = 'utf-8' }
    if ($stale.Ok) { throw 'A stale module SHA was accepted.' }
    $unchanged = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
    if ($unchanged.Sha256 -ne $after.Sha256) { throw 'A rejected request changed the module.' }
    $ansiSource = "' ANSI $($accent)preuve`r`nPublic Sub InsertedAnsi()`r`n    Debug.Print 24`r`nEnd Sub"
    [IO.File]::WriteAllBytes($ansiPath, [Text.Encoding]::GetEncoding(1252).GetBytes($ansiSource))
    $ansiInspection = Invoke-Vbe @{ Command = 'inspect_code_file'; Path = $ansiPath }
    if ($ansiInspection.StrictUtf8Valid -or -not $ansiInspection.ExplicitEncodingRequired) {
        throw 'Windows-1252 source inspection was not classified as ambiguous.'
    }
    $ansiModule = 'CodexAnsiInsertProbe'
    $ansiCreated = Invoke-Vbe @{ Command = 'create_module'; Project = $project; Module = $ansiModule; ExpectedMode = 2 }
    $ansiBefore = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $ansiModule }
    $ansiResult = Invoke-Vbe @{ Command = 'insert_code_file'; Project = $project; Module = $ansiModule;
        Path = $ansiPath; StartLine = ($ansiCreated.Lines + 1); ExpectedSha256 = $ansiBefore.Sha256;
        SourceEncoding = 'windows-1252' }
    $ansiAfter = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $ansiModule }
    if ($ansiResult.SourceEncoding -ne 'windows-1252' -or
        -not $ansiAfter.Code.Contains("ANSI $($accent)preuve") -or
        $ansiResult.Sha256 -ne $ansiAfter.Sha256) {
        throw 'Windows-1252 accented text was not preserved in the VBE.'
    }
    [pscustomobject]@{ HostProcessId = $HostProcessId; SourceSha256 = $sourceHash;
        ModuleSha256 = $after.Sha256; InsertedLineCount = $result.InsertedLineCount;
        ProcedureCount = @($procedures.Procedures).Count; StaleShaRejected = -not $stale.Ok;
        AmbiguousEncodingRejected = -not $ambiguous.Ok; Utf8AccentPreserved = $true;
        Windows1252AccentPreserved = $true } | Format-List
}
finally {
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    if (Test-Path -LiteralPath $ansiPath) { Remove-Item -LiteralPath $ansiPath -Force }
}

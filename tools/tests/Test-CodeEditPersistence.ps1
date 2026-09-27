param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$hosts = @(Get-Process EXCEL -ErrorAction Stop)
if ($hosts.Count -ne 1 -or $hosts[0].Id -ne $HostProcessId) {
    throw 'Exactly one isolated Excel process with the requested PID is required.'
}

function Invoke-Vbe([hashtable] $Request) {
    $payload = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $response = & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $payload |
        ConvertFrom-Json
    if (-not $response.Ok) { throw "$($Request.Command): $($response.Error)" }
    return $response.Data
}

$excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
if ($excel.Workbooks.Count -ne 1) { throw 'Exactly one disposable workbook is required.' }
$book = $excel.Workbooks.Item(1)
if ($book.Path) { throw 'The test workbook must be unsaved and disposable.' }
$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) {
    throw 'Exactly one design-mode VBA project is required.'
}
$project = [string]$projects[0].Name
$workbookPath = Join-Path $env:TEMP ("CodexVBE-code-persistence-{0}.xlsm" -f $HostProcessId)
$sourcePath = Join-Path $env:TEMP ("CodexVBE-code-persistence-{0}.bas" -f $HostProcessId)
if ((Test-Path -LiteralPath $workbookPath) -or (Test-Path -LiteralPath $sourcePath)) {
    throw 'A disposable test path already exists.'
}

try {
    $removeName = 'CodexPersistenceRemove'
    $insertName = 'CodexPersistenceInsert'
    $created = Invoke-Vbe @{ Command = 'create_module'; Project = $project;
        Module = $removeName; ExpectedMode = 2 }
    $code = "Option Explicit`r`nPublic Sub Gone()`r`n    Debug.Print 1`r`nEnd Sub`r`nPublic Sub Kept()`r`n    Debug.Print 2`r`nEnd Sub"
    Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $removeName;
        ExpectedSha256 = $created.Sha256; StartLine = 1; Count = $created.Lines; Text = $code } | Out-Null
    $before = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $removeName }
    $removed = Invoke-Vbe @{ Command = 'remove_procedure'; Project = $project; Module = $removeName;
        Procedure = 'Gone'; ProcKind = 0; ExpectedSha256 = $before.Sha256 }
    $insertedModule = Invoke-Vbe @{ Command = 'create_module'; Project = $project;
        Module = $insertName; ExpectedMode = 2 }
    $source = "Public Sub FromFile()`r`n    Debug.Print 3`r`nEnd Sub"
    [IO.File]::WriteAllText($sourcePath, $source, [Text.UTF8Encoding]::new($true))
    $inserted = Invoke-Vbe @{ Command = 'insert_code_file'; Project = $project;
        Module = $insertName; Path = $sourcePath; StartLine = ($insertedModule.Lines + 1);
        ExpectedSha256 = $insertedModule.Sha256 }
    $excel.DisplayAlerts = $false
    $book.SaveAs($workbookPath, 52)
    $savedRemove = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $removeName }
    $savedInsert = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $insertName }
    if ($savedRemove.Sha256 -ne $removed.Sha256 -or $savedInsert.Sha256 -ne $inserted.Sha256) {
        throw 'Code changed during Excel SaveAs.'
    }
    $book.Close($false)
    $book = $excel.Workbooks.Open($workbookPath)
    $reopenedProjects = @(Invoke-Vbe @{ Command = 'list_projects' })
    if ($reopenedProjects.Count -ne 1) { throw 'The reopened workbook has an unexpected project count.' }
    $project = [string]$reopenedProjects[0].Name
    $afterRemove = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $removeName }
    $afterInsert = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $insertName }
    $procedures = Invoke-Vbe @{ Command = 'list_procedures'; Project = $project; Module = $removeName }
    if ($afterRemove.Sha256 -ne $removed.Sha256 -or $afterInsert.Sha256 -ne $inserted.Sha256 -or
        @($procedures.Procedures | Where-Object { $_.Name -eq 'Gone' }).Count -ne 0 -or
        @($procedures.Procedures | Where-Object { $_.Name -eq 'Kept' }).Count -ne 1 -or
        $afterInsert.Code -notmatch 'Sub FromFile\(') {
        throw 'The code edits did not survive workbook reopen.'
    }
    [pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project;
        RemovedProcedurePersisted = $true; InsertedFilePersisted = $true;
        RemoveModuleSha256 = $afterRemove.Sha256; InsertModuleSha256 = $afterInsert.Sha256 } | Format-List
}
finally {
    try { if ($book -ne $null) { $book.Close($false) } } catch { }
    try { $excel.DisplayAlerts = $true } catch { }
    if (Test-Path -LiteralPath $sourcePath) { Remove-Item -LiteralPath $sourcePath -Force }
    if (Test-Path -LiteralPath $workbookPath) { Remove-Item -LiteralPath $workbookPath -Force }
}

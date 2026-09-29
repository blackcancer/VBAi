param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
if ((Get-Process -Id $HostProcessId -ErrorAction Stop).ProcessName -ne 'EXCEL') {
    throw 'An isolated Excel process is required.'
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
    throw 'Use one unsaved disposable design-mode VBA project.'
}
$project = [string]$projects[0].Name
$standardName = 'CodexRemovalProbe'
$className = 'CodexRemovalClass'
$standard = Invoke-Vbe @{ Command = 'create_module'; Project = $project; Module = $standardName; ExpectedMode = 2 }
$standardCode = "Option Explicit`r`n' Keep this comment`r`nPublic Sub First()`r`n    Debug.Print 1`r`nEnd Sub`r`n`r`nPublic Sub Second()`r`n    Debug.Print 2`r`nEnd Sub"
Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $standardName;
    ExpectedSha256 = $standard.Sha256; StartLine = 1; Count = $standard.Lines; Text = $standardCode } | Out-Null
$standardBefore = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $standardName }
$removed = Invoke-Vbe @{ Command = 'remove_procedure'; Project = $project; Module = $standardName;
    Procedure = 'First'; ProcKind = 0; ExpectedSha256 = $standardBefore.Sha256 }
if ($removed.Code -match 'Sub First\(' -or $removed.Code -notmatch "' Keep this comment" -or
    $removed.Code -notmatch 'Sub Second\(' -or $removed.Sha256 -eq $standardBefore.Sha256) {
    throw 'Standard module removal did not preserve its neighboring code and comment.'
}
$standardAfter = Invoke-Vbe @{ Command = 'list_procedures'; Project = $project; Module = $standardName }
if (@($standardAfter.Procedures).Count -ne 1 -or $standardAfter.Procedures[0].Name -ne 'Second') {
    throw 'The standard module procedure inventory is wrong after removal.'
}
$staleRejected = $false
try {
    Invoke-Vbe @{ Command = 'remove_procedure'; Project = $project; Module = $standardName;
        Procedure = 'Second'; ProcKind = 0; ExpectedSha256 = $standardBefore.Sha256 } | Out-Null
}
catch { $staleRejected = $true }
if (-not $staleRejected) { throw 'A stale SHA was accepted.' }

$class = Invoke-Vbe @{ Command = 'create_class'; Project = $project; Module = $className; ExpectedMode = 2 }
$classCode = "Private mValue As Long`r`n' Keep this property comment`r`nPublic Property Get Value() As Long`r`n    Value = mValue`r`nEnd Property`r`n`r`nPublic Property Let Value(ByVal newValue As Long)`r`n    mValue = newValue`r`nEnd Property"
Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $className;
    ExpectedSha256 = $class.Sha256; StartLine = 1; Count = $class.Lines; Text = $classCode } | Out-Null
$classBefore = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $className }
$propertyRemoved = Invoke-Vbe @{ Command = 'remove_procedure'; Project = $project; Module = $className;
    Procedure = 'Value'; ProcKind = 3; ExpectedSha256 = $classBefore.Sha256 }
if ($propertyRemoved.Code -match 'Property Get Value' -or
    $propertyRemoved.Code -notmatch 'Property Let Value' -or
    $propertyRemoved.Code -notmatch "' Keep this property comment") {
    throw 'Property Get removal changed another accessor or the preceding comment.'
}
$classAfter = Invoke-Vbe @{ Command = 'list_procedures'; Project = $project; Module = $className }
if (@($classAfter.Procedures | Where-Object { $_.Name -eq 'Value' -and $_.Kind -eq 3 }).Count -ne 0 -or
    @($classAfter.Procedures | Where-Object { $_.Name -eq 'Value' -and $_.Kind -eq 1 }).Count -ne 1) {
    throw 'The class accessor inventory is wrong after removal.'
}
[pscustomobject]@{ HostProcessId = $HostProcessId; StandardSha = $removed.Sha256;
    ClassSha = $propertyRemoved.Sha256; StandardRemaining = @($standardAfter.Procedures).Count;
    ClassRemaining = @($classAfter.Procedures).Count; StaleShaRejected = $staleRejected } | Format-List

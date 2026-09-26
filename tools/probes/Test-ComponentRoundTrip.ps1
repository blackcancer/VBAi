param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [Parameter(Mandatory = $true)] [ValidateSet('Class', 'Form')] [string] $Kind,
    [string] $Project = 'VBAProject',
    [string] $Module = $(if ($Kind -eq 'Class') { 'CodexClassRoundTrip' } else { 'CodexFormRoundTrip' })
)

$ErrorActionPreference = 'Stop'
function Invoke-Vbe([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $reply = & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $reply.Ok) { throw "$($Request.Command): $($reply.Error)" }
    return $reply.Data
}

$extension = if ($Kind -eq 'Class') { '.cls' } else { '.frm' }
$path = Join-Path $env:TEMP ("$Module-" + [guid]::NewGuid().ToString('N') + $extension)
$frxPath = [IO.Path]::ChangeExtension($path, '.frx')
$completed = $false
try {
    $before = Invoke-Vbe @{ Command = 'component_properties'; Project = $Project; Module = $Module }
    $codeBefore = Invoke-Vbe @{ Command = 'read_module'; Project = $Project; Module = $Module }
    $treeBefore = if ($Kind -eq 'Form') { Invoke-Vbe @{ Command = 'form_tree'; Project = $Project; Form = $Module } } else { $null }
    $export = Invoke-Vbe @{ Command = 'export_component'; Project = $Project; Module = $Module; Path = $path; ExpectedComponentVersion = $before.Version }
    $projectBeforeRemove = Invoke-Vbe @{ Command = 'project_properties'; Project = $Project }
    $componentBeforeRemove = Invoke-Vbe @{ Command = 'component_properties'; Project = $Project; Module = $Module }
    $afterRemove = Invoke-Vbe @{ Command = 'remove_component'; Project = $Project; Module = $Module; ExpectedProjectVersion = $projectBeforeRemove.Version; ExpectedComponentVersion = $componentBeforeRemove.Version }
    $import = Invoke-Vbe @{ Command = 'import_component'; Project = $Project; Path = $path; ExpectedProjectVersion = $afterRemove.Version }
    if (-not $import.Applied) { throw 'Import response did not confirm that the component was added.' }
    $importedName = $import.ImportedName
    if (-not $importedName) { throw 'Import response did not identify the added component.' }
    $codeAfter = Invoke-Vbe @{ Command = 'read_module'; Project = $Project; Module = $importedName }
    $treeAfter = if ($Kind -eq 'Form') { Invoke-Vbe @{ Command = 'form_tree'; Project = $Project; Form = $importedName } } else { $null }
    [pscustomobject]@{
        Kind = $Kind
        ExportBytes = $export.Bytes
        CompanionFrxBytes = if (Test-Path -LiteralPath $frxPath) { (Get-Item -LiteralPath $frxPath).Length } else { 0 }
        Applied = $import.Applied
        Verified = $import.Verified
        ImportError = $import.ImportError
        ComponentReadbackError = $import.ComponentReadbackError
        ProjectReadbackError = $import.ProjectReadbackError
        ImportedName = $importedName
        ImportedType = $import.Imported.Type
        CodeSame = $codeBefore.Code -ceq $codeAfter.Code
        CodeBeforeSha256 = $codeBefore.Sha256
        CodeAfterSha256 = $codeAfter.Sha256
        TreeNodeCountBefore = if ($treeBefore) { $treeBefore.NodeCount } else { $null }
        TreeNodeCountAfter = if ($treeAfter) { $treeAfter.NodeCount } else { $null }
        TreeVersionSame = if ($treeBefore) { $treeBefore.TreeVersion -ceq $treeAfter.TreeVersion } else { $null }
        TreeSame = if ($treeBefore) { ($treeBefore.Nodes | ConvertTo-Json -Depth 6 -Compress) -ceq ($treeAfter.Nodes | ConvertTo-Json -Depth 6 -Compress) } else { $null }
        ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
    } | Format-List
    $completed = $true
}
catch {
    Write-Error "Round trip stopped; exported files retained for diagnosis: $path ; $frxPath. $($_.Exception.Message)"
}
finally {
    if ($completed) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
        if (Test-Path -LiteralPath $frxPath) { Remove-Item -LiteralPath $frxPath -Force }
    }
}

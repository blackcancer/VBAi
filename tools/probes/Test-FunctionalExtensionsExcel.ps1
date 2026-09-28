param(
    [Parameter(Mandatory=$true)][string]$AssemblyPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [switch]$AllowTemporaryVbaAccess
)
$ErrorActionPreference = 'Stop'
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close existing Excel instances before this isolated qualification.' }
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$outputRoot = [IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($outputRoot) | Out-Null
$securityPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$initial = Get-ItemProperty -LiteralPath $securityPath
$hadAccess = $null -ne $initial.PSObject.Properties['AccessVBOM']
$initialAccess = $initial.AccessVBOM
[pscustomobject]@{ HadAccessVBOM=$hadAccess; AccessVBOM=$initialAccess } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $outputRoot 'security-before.json') -Encoding UTF8
if ($initialAccess -ne 1 -and -not $AllowTemporaryVbaAccess) { throw 'Temporary AccessVBOM permission is required.' }
$excel = $null; $book = $null; $otherBook = $null; $session = $null; $createdToolbar = $null
function Invoke-Session([hashtable]$Fields) {
    $request = New-Object CodexVBE.Request
    foreach ($key in $Fields.Keys) { $request.$key = $Fields[$key] }
    $result = $script:session.Execute($request)
    if (-not $result.Ok) { throw $result.Error }
    return $result.Data
}
try {
    if ($initialAccess -ne 1) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null }
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $excel.Visible = $true
    $book = $excel.Workbooks.Add()
    $vbe = $excel.VBE
    $vbe.MainWindow.Visible = $true
    $project = $book.VBProject
    $component = $project.VBComponents.Add(1)
    $component.Name = 'FunctionalProbe'
    $source = "Public Sub RenameProbe()`r`nDim value As Long`r`nvalue = 7`r`nThisWorkbook.Worksheets(1).Cells(1, 1).Value2 = value`r`nEnd Sub"
    $component.CodeModule.AddFromString($source)
    $document = Join-Path $outputRoot ('RenameProbe-' + [Guid]::NewGuid().ToString('N') + '.xlsm')
    $book.SaveAs($document, 52)
    $otherBook = $excel.Workbooks.Add()
    $otherProject = $otherBook.VBProject
    $otherName = $otherProject.Name
    $form = $project.VBComponents.Add(3)
    $form.Name = 'RenameFormProbe'
    $component.CodeModule.CodePane.Show()
    $originalName = $project.Name
    $renameTrials = @()
    foreach ($requestedName in @('QualifiedProbeA','QualifiedProbeB',$originalName,'QualifiedProbeA','QualifiedProbeB',$originalName)) {
        $project.Name = $requestedName
        if ($project.Name -cne $requestedName -or $project.FileName -ine $document -or $otherProject.Name -cne $otherName -or $component.CodeModule.Lines(1,$component.CodeModule.CountOfLines) -cne $source) { throw 'Native project rename changed the wrong scope or altered source.' }
        $renameTrials += [pscustomobject]@{ Name=$project.Name; Path=$project.FileName; OtherProjectName=$otherProject.Name; SourcePreserved=$true; FormName=$form.Name }
    }
    $renameTrials | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $outputRoot 'native-project-rename-trials.json') -Encoding UTF8
    $testAssemblyPath = Join-Path ([IO.FileInfo]$AssemblyPath).Directory.Parent.Parent.Parent.FullName 'CodexVBE.Tests/Debug/net48/CodexVBE.Tests.dll'
    $testAssembly = [Reflection.Assembly]::LoadFrom($testAssemblyPath)
    $scopeProbe = [Activator]::CreateInstance($testAssembly.GetType('CodexVBE.Tests.Infrastructure.ScopedExcelHostProbe', $true))
    $scopeProbe.Application = $excel
    $scopeProbe.ProcessId = [int]$scopeProbe.WindowProcessId([IntPtr][int]$excel.Hwnd)
    $formsService = [Activator]::CreateInstance($assembly.GetType('CodexVBE.VbeForms', $true), [object[]]@($vbe))
    $projectService = [Activator]::CreateInstance($assembly.GetType('CodexVBE.VbeProjectComponents', $true), [Reflection.BindingFlags]'Instance, NonPublic', $null, [object[]]@($vbe,$formsService,$scopeProbe), $null)
    $metadata = $projectService.ProjectProperties($document)
    $nativeRenameRequest = New-Object CodexVBE.Request
    $nativeRenameRequest.Project = $document; $nativeRenameRequest.Property = 'Name'; $nativeRenameRequest.Value = 'QualifiedServiceProbe'; $nativeRenameRequest.ExpectedProjectVersion = $metadata.Version
    $serviceRename = $projectService.SetProjectProperty($nativeRenameRequest)
    if (-not $serviceRename.Verified -or -not $serviceRename.SourcePreserved) { throw 'The production rename service did not verify its mutation.' }
    $metadata = $projectService.ProjectProperties($document)
    $nativeRenameRequest.Value = $originalName; $nativeRenameRequest.ExpectedProjectVersion = $metadata.Version
    $serviceRestore = $projectService.SetProjectProperty($nativeRenameRequest)
    if (-not $serviceRestore.Verified) { throw 'The production rename service did not restore the original project name.' }
    [pscustomobject]@{ Rename=$serviceRename; Restore=$serviceRestore; NativeHostProcessId=$scopeProbe.ProcessId; ExternalComProbe=$true } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $outputRoot 'native-project-rename-service.json') -Encoding UTF8
    $session = [Activator]::CreateInstance($assembly.GetType('CodexVBE.VbeSession', $true), [object[]]@($vbe))
    $symbols = Invoke-Session @{ Command='project_symbols'; Project=$document; Module=$component.Name; Query='value'; WholeWord=$true }
    if ($symbols.Total -ne 1) { throw 'The local declaration was not indexed uniquely.' }
    $local = $symbols.Symbols[0]
    $rename = @{ Command='preview_local_rename'; Project=$document; Module=$component.Name; Procedure='RenameProbe'; ProcKind=0; ExpectedSha256=$local.Sha256; StartLine=$local.Line; StartColumn=$local.Column; Query='value'; NewName='amount' }
    $preview = Invoke-Session $rename
    if (-not $preview.Changed -or $preview.After -notmatch 'Dim amount As Long') { throw 'Local rename preview failed.' }
    $rename.Command='apply_local_rename'; $rename.ExpectedMode=2
    $applied = Invoke-Session $rename
    $read = Invoke-Session @{ Command='read_module'; Project=$document; Module=$component.Name }
    if ($read.Code -notmatch 'Value2 = amount') { throw 'The native module did not retain the renamed references.' }
    $excel.Run("'" + $book.Name.Replace("'","''") + "'!FunctionalProbe.RenameProbe")
    if ($book.Worksheets(1).Cells(1,1).Value2 -ne 7) { throw 'The renamed procedure did not execute its expected result.' }
    $undo = Invoke-Session @{ Command='undo_code_edit'; Project=$document; Module=$component.Name; ExpectedSha256=$read.Sha256 }
    $restored = Invoke-Session @{ Command='read_module'; Project=$document; Module=$component.Name }
    if ($restored.Code -cne $preview.Before) { throw 'Managed undo did not restore the complete module.' }
    $toolbars = Invoke-Session @{ Command='list_toolbars' }
    $custom = Invoke-Session @{ Command='create_toolbar'; ObjectName=('Probe-' + [Guid]::NewGuid().ToString('N')); ExpectedToolbarCollectionVersion=$toolbars.ToolbarCollectionVersion }
    $createdToolbar = $custom.ObjectName
    $state = Invoke-Session @{ Command='toolbar_controls'; ObjectName=$createdToolbar }
    $command = $vbe.CommandBars.FindControl(1, 186)
    if ($null -eq $command) { throw 'Native Run command is unavailable.' }
    $added = Invoke-Session @{ Command='add_toolbar_command'; ObjectName=$createdToolbar; ExpectedToolbarControlsVersion=$state.ToolbarControlsVersion; ControlId=186; ControlCaption=$command.Caption }
    if (-not $added.Verified) { throw ('Native button addition was not verified: ' + ($added | ConvertTo-Json -Depth 8 -Compress)) }
    $state = Invoke-Session @{ Command='toolbar_controls'; ObjectName=$createdToolbar }
    $button = $state.Controls[0]
    $removed = Invoke-Session @{ Command='remove_toolbar_command'; ObjectName=$createdToolbar; ExpectedToolbarControlsVersion=$state.ToolbarControlsVersion; ControlId=$button.Id; ControlCaption=$button.Caption; InsertIndex=$button.Index }
    if (-not $removed.Verified) { throw 'Native button removal failed.' }
    $state = Invoke-Session @{ Command='toolbar_controls'; ObjectName=$createdToolbar }
    $toolbars = Invoke-Session @{ Command='list_toolbars' }
    $deleted = Invoke-Session @{ Command='remove_toolbar'; ObjectName=$createdToolbar; ExpectedToolbarControlsVersion=$state.ToolbarControlsVersion; ExpectedToolbarCollectionVersion=$toolbars.ToolbarCollectionVersion }
    if (-not $deleted.Verified) { throw 'Native toolbar removal failed.' }
    $createdToolbar = $null
    [pscustomobject]@{ Symbol=$local; RenamePreview=$preview; Applied=$applied; RuntimeCell=7; UndoRestored=$true; ToolbarCreated=$custom; ButtonAdded=$added; ButtonRemoved=$removed; ToolbarDeleted=$deleted; ParameterizedImmediateRuntime='NOT_RUN: external probe is outside the add-in process'; AssemblyModuleVersionId=$assembly.ManifestModule.ModuleVersionId.ToString() } | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $outputRoot 'functional-extensions-excel.json') -Encoding UTF8
}
catch {
    $_ | Out-String | Set-Content -LiteralPath (Join-Path $outputRoot 'probe-failure.txt') -Encoding UTF8
    throw
}
finally {
    if ($null -ne $createdToolbar -and $null -ne $excel) { try { $excel.VBE.CommandBars.Item($createdToolbar).Delete() } catch { } }
    if ($null -ne $otherBook) { try { $otherBook.Close($false) } catch { Write-Warning $_ } }
    if ($null -ne $book) { try { $book.Close($false) } catch { Write-Warning $_ } }
    if ($null -ne $excel) { try { $excel.Quit() } catch { Write-Warning $_ } }
    if ($null -ne $book -and [Runtime.InteropServices.Marshal]::IsComObject($book)) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) } catch { Write-Warning $_ } }
    if ($null -ne $excel -and [Runtime.InteropServices.Marshal]::IsComObject($excel)) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch { Write-Warning $_ } }
    if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value $initialAccess -PropertyType DWord -Force | Out-Null }
    else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

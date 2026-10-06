<#
.SYNOPSIS
Runs one explicit VBE event or functional-extension scenario in owned Excel.
.DESCRIPTION
Collection, Reference and ConnectionPoints preserve separate event observation
contracts. FunctionalExtensions uses an in-process session; the registered
scenario requires -UseBridge and verifies the loaded process and assembly.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Collection','Reference','ConnectionPoints','FunctionalExtensions','RegisteredFunctionalExtensions')]
    [string]$Scenario,
    [ValidateNotNullOrEmpty()][string]$AssemblyPath,
    [ValidateNotNullOrEmpty()][string]$OutputDirectory,
    [switch]$UseBridge,
    [switch]$AllowTemporaryVbaAccess
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'VbeProbe.Common.ps1')
$VbeProbeTraceRequests = $false
$required = @('AssemblyPath','OutputDirectory')
$allowed = @('AssemblyPath','OutputDirectory','AllowTemporaryVbaAccess')
if ($Scenario -eq 'RegisteredFunctionalExtensions') { $required += 'UseBridge'; $allowed += 'UseBridge' }
Assert-VbeProbeParameters -Bound $PSBoundParameters -Required $required -Allowed $allowed

switch ($Scenario) {
    'Collection' {
        $ErrorActionPreference = 'Stop'
        . $VbeProbeInitializeAssembly

        try {
            . $VbeProbeOpenAssemblyExcel
            $vbe.MainWindow.Visible = $true
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
            $project = $book.VBProject
            $flags = [Reflection.BindingFlags]'Instance,NonPublic'
            $eventType = $assembly.GetType('VBAi.VbeCollectionEvents', $true)
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -TypeDefinition 'public static class CollectionEventCounter { public static int Components; public static int Projects; public static void ComponentChanged() { Components++; } public static void ProjectChanged() { Projects++; } public static void Observe(object session, object listener, string method, string project) { var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic; var source = session.GetType().GetMethod(method, flags).Invoke(session, method == "ProjectsEventSource" ? new object[0] : new object[] { project }); listener.GetType().GetMethod("Observe", flags).Invoke(listener, new object[] { source }); } }'
            function Pump-Events { for ($i=0; $i -lt 10; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 50 } }
            $componentsCallback = [Delegate]::CreateDelegate([Action], [CollectionEventCounter].GetMethod('ComponentChanged'))
            $projectsCallback = [Delegate]::CreateDelegate([Action], [CollectionEventCounter].GetMethod('ProjectChanged'))
            $componentListener = [Activator]::CreateInstance($eventType, $flags, $null, [object[]]@($componentsCallback,$true,$null,$null), $null)
            $projectListener = [Activator]::CreateInstance($eventType, $flags, $null, [object[]]@($projectsCallback,$false,$null,$null), $null)
            $observe = $eventType.GetMethod('Observe', $flags)
            [CollectionEventCounter]::Observe($session, $projectListener, 'ProjectsEventSource', $null)
            [CollectionEventCounter]::Observe($session, $componentListener, 'ComponentsEventSource', $project.Name)
            $component = $project.VBComponents.Add(1)
            Pump-Events
            $componentAdded = [CollectionEventCounter]::Components
            $component.Name = 'RenamedEventProbe'
            Pump-Events
            $componentRenamed = [CollectionEventCounter]::Components
            $project.VBComponents.Remove($component)
            Pump-Events
            $componentRemoved = [CollectionEventCounter]::Components
            $projectBefore = [CollectionEventCounter]::Projects
            $otherBook = $excel.Workbooks.Add()
            Pump-Events
            $projectAdded = [CollectionEventCounter]::Projects
            $otherBook.VBProject.Name = 'OtherCollectionScope'
            Pump-Events
            $projectRenamed = [CollectionEventCounter]::Projects
            [CollectionEventCounter]::Observe($session, $componentListener, 'ComponentsEventSource', 'OtherCollectionScope')
            $component = $project.VBComponents.Add(1); $project.VBComponents.Remove($component)
            Pump-Events
            $afterOldScope = [CollectionEventCounter]::Components
            $component = $otherBook.VBProject.VBComponents.Add(1)
            Pump-Events
            $afterNewScope = [CollectionEventCounter]::Components
            $otherBook.Close($false); $otherBook = $null
            Pump-Events
            $projectRemoved = [CollectionEventCounter]::Projects
            $componentListener.Dispose(); $projectListener.Dispose()
            $componentBeforeDispose = [CollectionEventCounter]::Components
            $component = $project.VBComponents.Add(1); $project.VBComponents.Remove($component)
            $otherBook = $excel.Workbooks.Add()
            Pump-Events
            $afterDisposeComponents = [CollectionEventCounter]::Components
            $afterDisposeProjects = [CollectionEventCounter]::Projects
            $result = [pscustomobject]@{ ComponentAdded=$componentAdded; ComponentRenamed=$componentRenamed; ComponentRemoved=$componentRemoved; AfterOldScope=$afterOldScope; AfterNewScope=$afterNewScope; ProjectBefore=$projectBefore; ProjectAdded=$projectAdded; ProjectRenamed=$projectRenamed; ProjectRemoved=$projectRemoved; ComponentBeforeDispose=$componentBeforeDispose; AfterDisposeComponents=$afterDisposeComponents; AfterDisposeProjects=$afterDisposeProjects; Mvid=$assembly.ManifestModule.ModuleVersionId.ToString() }
            $result | ConvertTo-Json | Set-Content (Join-Path $directory 'collection-events-native.json') -Encoding UTF8
            $result | ConvertTo-Json
            if ($componentAdded -le 0 -or $componentRenamed -le $componentAdded -or $componentRemoved -le $componentRenamed -or $afterOldScope -ne $componentRemoved -or $afterNewScope -le $afterOldScope -or $projectAdded -le $projectBefore -or $projectRenamed -le $projectAdded -or $projectRemoved -le $projectRenamed -or $afterDisposeComponents -ne $componentBeforeDispose -or $afterDisposeProjects -ne $projectRemoved) { throw 'Native collection notifications or detach failed.' }
        } finally {
            if ($null -ne $componentListener) { $componentListener.Dispose() }
            if ($null -ne $projectListener) { $projectListener.Dispose() }
            if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
            else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
            try { if ($null -ne $otherBook) { $otherBook.Close($false) }; if ($null -ne $book) { $book.Close($false) } }
            finally {
                try { if ($null -ne $excel) { $excel.Quit() } }
                finally { if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) { Stop-Process -Id $probeProcess.Id -Force } }
            }
        }

    }
    'Reference' {
        $ErrorActionPreference = 'Stop'
        . $VbeProbeInitializeAssembly

        try {
            . $VbeProbeOpenAssemblyExcel
            $vbe.MainWindow.Visible = $true
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
            $project = $book.VBProject
            $flags = [Reflection.BindingFlags]'Instance,NonPublic'
            $source = $session.GetType().GetMethod('ReferenceEventSource', $flags).Invoke($session, @($project.Name))
            $eventType = $assembly.GetType('VBAi.VbeReferenceEvents', $true)
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -TypeDefinition 'public static class ReferenceEventCounter { public static int Count; public static void Changed() { Count++; } }'
            [ReferenceEventCounter]::Count = 0
            function Pump-Events { for ($i=0; $i -lt 10; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 50 } }
            $callback = [Delegate]::CreateDelegate([Action], [ReferenceEventCounter].GetMethod('Changed'))
            $listener = [Activator]::CreateInstance($eventType, $flags, $null, [object[]]@($callback,$null,$null), $null)
            $observe = $eventType.GetMethod('Observe', $flags)
            $observe.Invoke($listener, @($source)) | Out-Null
            $reference = $project.References.AddFromGuid('{420B2830-E718-11CF-893D-00A0C9054228}',1,0)
            Pump-Events
            $afterAdd = [ReferenceEventCounter]::Count
            $project.References.Remove($reference)
            Pump-Events
            $afterRemove = [ReferenceEventCounter]::Count
            $otherBook = $excel.Workbooks.Add()
            $otherBook.VBProject.Name = 'OtherEventScope'
            $otherSource = $session.GetType().GetMethod('ReferenceEventSource', $flags).Invoke($session, @('OtherEventScope'))
            $observe.Invoke($listener, @($otherSource)) | Out-Null
            $reference = $project.References.AddFromGuid('{420B2830-E718-11CF-893D-00A0C9054228}',1,0)
            $project.References.Remove($reference)
            Pump-Events
            $afterOldScope = [ReferenceEventCounter]::Count
            $reference = $otherBook.VBProject.References.AddFromGuid('{420B2830-E718-11CF-893D-00A0C9054228}',1,0)
            $otherBook.VBProject.References.Remove($reference)
            Pump-Events
            $afterNewScope = [ReferenceEventCounter]::Count
            $listener.Dispose()
            $reference = $otherBook.VBProject.References.AddFromGuid('{420B2830-E718-11CF-893D-00A0C9054228}',1,0)
            $otherBook.VBProject.References.Remove($reference)
            Pump-Events
            $afterDispose = [ReferenceEventCounter]::Count
            [pscustomobject]@{ AfterAdd=$afterAdd; AfterRemove=$afterRemove; AfterDispose=$afterDispose; AfterOldScope=$afterOldScope; AfterNewScope=$afterNewScope; Assembly=$assembly.Location; Mvid=$assembly.ManifestModule.ModuleVersionId.ToString() } | ConvertTo-Json | Set-Content (Join-Path $directory 'reference-events-native.json') -Encoding UTF8
            if ($afterAdd -ne 1 -or $afterRemove -ne 2 -or $afterDispose -ne 4 -or $afterOldScope -ne 2 -or $afterNewScope -ne 4) { throw 'Native reference notifications or unsubscribe failed.' }
        } finally {
            if ($null -ne $listener) { $listener.Dispose() }
            if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
            else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
            try { if ($null -ne $otherBook) { $otherBook.Close($false) }; if ($null -ne $book) { $book.Close($false) } }
            finally {
                try { if ($null -ne $excel) { $excel.Quit() } }
                finally { if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) { Stop-Process -Id $probeProcess.Id -Force } }
            }
        }

    }
    'ConnectionPoints' {
        $ErrorActionPreference = 'Stop'
        . $VbeProbeInitializeAssembly

        try {
            . $VbeProbeOpenAssemblyExcel
            $vbe.MainWindow.Visible = $true
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
            $project = $book.VBProject
            Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
public static class VbeConnectionProbe {
    public static string[] Read(object target) {
        var result = new List<string>();
        var container = target as IConnectionPointContainer;
        if (container == null) return new[] { "IConnectionPointContainer unavailable" };
        foreach (var text in new[] { "0002E103-0000-0000-C000-000000000046", "0002E116-0000-0000-C000-000000000046", "CDDE3804-2064-11CF-867F-00AA005FF34A" }) {
            Guid iid = new Guid(text);
            try { IConnectionPoint point; container.FindConnectionPoint(ref iid, out point); if (point == null) result.Add(text + ": null"); else { Guid actual; point.GetConnectionInterface(out actual); result.Add(text + ": actual " + actual.ToString()); } }
            catch (Exception error) { result.Add(text + ": " + error.GetType().Name + " " + error.HResult.ToString("X8")); }
        }
        return result.ToArray();
    }
}
"@
            $result = @()
            foreach ($entry in @(@{ Name='VBE'; Source=$vbe }, @{ Name='Events'; Source=$vbe.Events }, @{ Name='VBProjects'; Source=$vbe.VBProjects }, @{ Name='VBProject'; Source=$project }, @{ Name='VBComponents'; Source=$project.VBComponents }, @{ Name='References'; Source=$project.References })) {
                try { $result += [pscustomobject]@{ Name=$entry.Name; Interfaces=[VbeConnectionProbe]::Read($entry.Source); Error=$null } }
                catch { $result += [pscustomobject]@{ Name=$entry.Name; Interfaces=@(); Error=$_.Exception.Message } }
            }
            $result | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $directory 'vbe-event-connection-points.json') -Encoding UTF8
            $result | ConvertTo-Json -Depth 5
        } finally {
            if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
            else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
            try { if ($null -ne $otherBook) { $otherBook.Close($false) }; if ($null -ne $book) { $book.Close($false) } }
            finally {
                try { if ($null -ne $excel) { $excel.Quit() } }
                finally { if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) { Stop-Process -Id $probeProcess.Id -Force } }
            }
        }

    }
    'FunctionalExtensions' {
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
            $testAssemblyPath = Join-Path ([IO.FileInfo]$AssemblyPath).Directory.Parent.Parent.Parent.FullName 'VBAi.Tests/Debug/net48/VBAi.Tests.dll'
            $testAssembly = [Reflection.Assembly]::LoadFrom($testAssemblyPath)
            $scopeProbe = [Activator]::CreateInstance($testAssembly.GetType('VBAi.Tests.Infrastructure.ScopedExcelHostProbe', $true))
            $scopeProbe.Application = $excel
            $scopeProbe.ProcessId = [int]$scopeProbe.WindowProcessId([IntPtr][int]$excel.Hwnd)
            $formsService = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeForms', $true), [object[]]@($vbe))
            $projectService = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeProjectComponents', $true), [Reflection.BindingFlags]'Instance, NonPublic', $null, [object[]]@($vbe,$formsService,$scopeProbe), $null)
            $metadata = $projectService.ProjectProperties($document)
            $nativeRenameRequest = New-Object VBAi.Request
            $nativeRenameRequest.Project = $document; $nativeRenameRequest.Property = 'Name'; $nativeRenameRequest.Value = 'QualifiedServiceProbe'; $nativeRenameRequest.ExpectedProjectVersion = $metadata.Version
            $serviceRename = $projectService.SetProjectProperty($nativeRenameRequest)
            if (-not $serviceRename.Verified -or -not $serviceRename.SourcePreserved) { throw 'The production rename service did not verify its mutation.' }
            $metadata = $projectService.ProjectProperties($document)
            $nativeRenameRequest.Value = $originalName; $nativeRenameRequest.ExpectedProjectVersion = $metadata.Version
            $serviceRestore = $projectService.SetProjectProperty($nativeRenameRequest)
            if (-not $serviceRestore.Verified) { throw 'The production rename service did not restore the original project name.' }
            [pscustomobject]@{ Rename=$serviceRename; Restore=$serviceRestore; NativeHostProcessId=$scopeProbe.ProcessId; ExternalComProbe=$true } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $outputRoot 'native-project-rename-service.json') -Encoding UTF8
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
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

    }
    'RegisteredFunctionalExtensions' {
        $ErrorActionPreference='Stop'
        . $VbeProbeInitializeRegistered
        $excel=$null;$book=$null;$vbe=$null;$probeProcess=$null;$toolbarName=$null;$option=$null;$optionChanged=$false;$report=[ordered]@{};$cleanupErrors=@()



        try{
            if($initialAccess -ne 1){New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null}
            $report.Identity=Open-ProbeExcel
            $book=$excel.Workbooks.Add();$document=Join-Path $outputRoot ('Arguments-'+[Guid]::NewGuid().ToString('N')+'.xlsm');$book.SaveAs($document,52)
            $module=$book.VBProject.VBComponents.Add(1);$module.Name='ArgumentsProbe'
            $module.CodeModule.AddFromString(@"
Option Explicit
Public Sub AcceptArguments(ByVal text As String, ByVal number As Double, ByVal flag As Boolean, ByVal absent As Variant)
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = text
    ThisWorkbook.Worksheets(1).Range("B1").Value2 = number
    ThisWorkbook.Worksheets(1).Range("C1").Value2 = flag
    ThisWorkbook.Worksheets(1).Range("D1").Value2 = IsNull(absent)
End Sub
Public Function ReturnResult(ByVal number As Double) As Double
    ReturnResult = number * 2
    ThisWorkbook.Worksheets(1).Range("E1").Value2 = ReturnResult
End Function
"@)
            $source=Invoke-Bridge @{Command='read_module';Project=$document;Module='ArgumentsProbe'};$runs=@()
            foreach($procedure in @('AcceptArguments','ReturnResult')){
                $arguments=if($procedure -eq 'AcceptArguments'){@('VBAi "quoted" français',12.5,$true,$null)}else{@(21)}
                $queued=Invoke-Bridge @{Command='run_procedure';Project=$document;Module='ArgumentsProbe';Procedure=$procedure;ExpectedMode=2;ExpectedSha256=$source.Sha256;Arguments=@($arguments)}
                for($i=0;$i -lt 30;$i++){Start-Sleep -Milliseconds 200;$state=Invoke-Bridge @{Command='procedure_run_status';Project=$document;Query=$queued.Query};if(-not $state.Pending){break}}
                if($state.State -ne 'Delivered'){throw ('Procedure was not delivered: '+($state|ConvertTo-Json -Depth 8 -Compress))};$runs+=$state
            }
            if($book.Worksheets.Item(1).Range('A1').Value2 -cne 'VBAi "quoted" français' -or $book.Worksheets.Item(1).Range('B1').Value2 -ne 12.5 -or -not $book.Worksheets.Item(1).Range('C1').Value2 -or -not $book.Worksheets.Item(1).Range('D1').Value2 -or $book.Worksheets.Item(1).Range('E1').Value2 -ne 42){throw 'Independent worksheet readback differs from the fixture.'}
            $report.Procedures=@{Runs=$runs;RuntimeReadbackVerified=$true;FunctionResult=42}

            $matrix=@();$catalog=Invoke-Bridge @{Command='list_form_control_types'}
            foreach($type in @('CheckBox','ComboBox','CommandButton','Frame','Image','Label','ListBox','MultiPage','OptionButton','ScrollBar','SpinButton','TabStrip','TextBox','ToggleButton')){
                $form=$book.VBProject.VBComponents.Add(3);$form.Name='Matrix'+$type
                $control=$form.Designer.Controls.Add(('Forms.'+$type+'.1'),'ProbeControl',$true)
                $properties=Invoke-Bridge @{Command='form_control_properties';Project=$document;Form=$form.Name;Control='ProbeControl'}
                foreach($property in @('Left','Enabled','Caption')){
                    $descriptor=@($properties|Where-Object Name -eq $property)
                    if($descriptor.Count -ne 1 -or $descriptor[0].ReadOnly -or $descriptor[0].Error){$matrix+=@{Type=$type;Property=$property;State='NOT_APPLICABLE';Reason='Property absent, read-only or unreadable'};continue}
                    $original=$control.$property
                    if($null -eq $original -and $property -eq 'Caption'){$control.Caption='Baseline caption';$original=$control.Caption}
                    if($null -eq $original){$matrix+=@{Type=$type;Property=$property;State='NOT_APPLICABLE';Reason='Null original value cannot be restored by the scalar setter'};continue}
                    $value=if($property -eq 'Left'){[double]$original+9.5}elseif($property -eq 'Enabled'){-not [bool]$original}else{'VBAi property fixture'}
                    $tree=Invoke-Bridge @{Command='form_tree';Project=$document;Form=$form.Name};$node=@($tree.Controls|Where-Object Name -eq 'ProbeControl')[0]
                    @{Type=$type;Property=$property;Node=$node;TreeVersion=$tree.TreeVersion;Value=$value;Original=$original}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $outputRoot 'current-property-trial.json') -Encoding UTF8
                    Invoke-Bridge @{Command='set_form_node_property';Project=$document;Form=$form.Name;ControlPath=$node.Path;Property=$property;Value=$value;ExpectedTreeVersion=$tree.TreeVersion}|Out-Null
                    if($control.$property -ne $value){throw "Independent property readback failed: $type/$property"}
                    $tree=Invoke-Bridge @{Command='form_tree';Project=$document;Form=$form.Name}
                    Invoke-Bridge @{Command='set_form_node_property';Project=$document;Form=$form.Name;ControlPath=$node.Path;Property=$property;Value=$original;ExpectedTreeVersion=$tree.TreeVersion}|Out-Null
                    if($control.$property -ne $original){throw "Property restoration failed: $type/$property"}
                    $matrix+=@{Type=$type;Property=$property;State='PASS';Before=$original;Changed=$value;Restored=$true}
                }
            }

            $activex=@()
            foreach($candidate in @($catalog|Where-Object { $_.Source -eq 'COM CATID_Control x64' -and $_.ProgId -like 'MSComctlLib.*' })){
                $form=$book.VBProject.VBComponents.Add(3);$form.Name='ActiveXProbe'+$activex.Count
                $state=Invoke-Bridge @{Command='form_state';Project=$document;Form=$form.Name}
                try{
                    $addedControl=Invoke-Bridge @{Command='add_form_control';Project=$document;Form=$form.Name;ControlType=$candidate.ProgId;Control='HostedControl';Left=10;Top=10;Width=150;Height=40;ExpectedFormVersion=$state.Version}
                    if($form.Designer.Controls.Count -ne 1){throw 'Native control count differs from the expected hosting fixture.'}
                    $activex+=@{ProgId=$candidate.ProgId;State='HOSTED';NativeCount=$form.Designer.Controls.Count}
                }catch{
                    $activex+=@{ProgId=$candidate.ProgId;State='HOSTING_REFUSED';Error=$_.Exception.Message;ControlsRemaining=$form.Designer.Controls.Count}
                }
            }
            $report.ActiveXHosting=$activex
            $activex|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $outputRoot 'activex-hosting.json') -Encoding UTF8
            $report.ControlProperties=$matrix;$report.InstalledActiveX=$catalog
            $matrix|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $outputRoot 'control-property-matrix.json') -Encoding UTF8
            $before=Invoke-Bridge @{Command='read_vbe_options'};$before | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $outputRoot 'options-before.json') -Encoding UTF8
            foreach($tab in $before.Tabs){foreach($control in $tab.Controls){
                    if($tab.Tab -match '^(Editor|Éditeur|Editeur)$' -and $control.Type -eq 'ControlType.CheckBox' -and $control.Name.Replace('&','').Trim() -match '^(Auto Syntax Check|Vérification automatique de la syntaxe)$' -and $control.Value -in @('On','Off')){$option=@{Pane=$tab.Tab;Property=$control.Name;Original=($control.Value -eq 'On')}}
            }}
            if($null -eq $option){throw 'The exact supported syntax-check option could not be identified.'}
            $optionChanged=$true
            $changed=Invoke-Bridge @{Command='set_vbe_option';Pane=$option.Pane;Property=$option.Property;Value=(-not $option.Original);ExpectedOptionsVersion=$before.OptionsVersion}
            $after=Invoke-Bridge @{Command='read_vbe_options'}
            $actual=@($after.Tabs|Where-Object Tab -eq $option.Pane|ForEach-Object Controls|Where-Object Name -eq $option.Property);$expected=if($option.Original){'Off'}else{'On'}
            if($actual.Count -ne 1 -or $actual[0].Value -ne $expected){throw 'Preference did not persist after reopening.'}
            Invoke-Bridge @{Command='set_vbe_option';Pane=$option.Pane;Property=$option.Property;Value=$option.Original;ExpectedOptionsVersion=$after.OptionsVersion}|Out-Null
            $restoredRead=Invoke-Bridge @{Command='read_vbe_options'};if($restoredRead.OptionsVersion -ne $before.OptionsVersion){throw 'Full options state was not restored.'};$optionChanged=$false
            $report.Options=@{Changed=$changed;ReopenedValue=$actual[0].Value;OriginalStateRestored=$true}
            $optionMatrix=@()
            $baseline=Invoke-Bridge @{Command='read_vbe_options'}
            foreach($tab in $baseline.Tabs){
                if($tab.Tab -notin @('Éditeur','Editor','Général','General')){continue}
                $originalRadio=@($tab.Controls|Where-Object { $_.Type -eq 'ControlType.RadioButton' -and $_.Value -eq $true })
                foreach($control in $tab.Controls){
                    if($control.Type -eq 'ControlType.CheckBox' -and $control.Name -match '^(Vérification automatique de la syntaxe|Déclaration des variables obligatoire|Complément automatique des instructions|Info express automatique|Info-bulles automatiques|Retrait automatique|Compilation sur demande|Compilation en arrière-plan|Auto Syntax Check|Require Variable Declaration|Auto List Members|Auto Quick Info|Auto Data Tips|Auto Indent|Compile on Demand|Background Compile)$'){
                        $values=@(-not ($control.Value -eq 'On'));$restoreProperty=$control.Name;$restoreValue=($control.Value -eq 'On')
                    }elseif($control.Type -eq 'ControlType.Edit' -and $control.Name -match '^(Largeur de la tabulation|Tab Width)'){
                        $values=@(1,32);$restoreProperty=$control.Name;$restoreValue=[int]$control.Value
                    }elseif($control.Type -eq 'ControlType.RadioButton'){
                        if($originalRadio.Count -ne 1){throw 'Original error-trapping radio is ambiguous.'}
                        $values=@($true);$restoreProperty=$originalRadio[0].Name;$restoreValue=$true
                    }else{continue}
                    foreach($value in $values){
                        $before=Invoke-Bridge @{Command='read_vbe_options'}
                        $option=@{Pane=$tab.Tab;Property=$restoreProperty;Original=$restoreValue};$optionChanged=$true
                        Invoke-Bridge @{Command='set_vbe_option';Pane=$tab.Tab;Property=$control.Name;Value=$value;ExpectedOptionsVersion=$before.OptionsVersion}|Out-Null
                        $after=Invoke-Bridge @{Command='read_vbe_options'}
                        $actual=@($after.Tabs|Where-Object Tab -eq $tab.Tab|ForEach-Object Controls|Where-Object { $_.Name -eq $control.Name -and $_.Type -eq $control.Type })
                        $expected=if($control.Type -eq 'ControlType.CheckBox'){if($value){'On'}else{'Off'}}elseif($control.Type -eq 'ControlType.Edit'){[string]$value}else{$true}
                        if($actual.Count -ne 1 -or $actual[0].Value -ne $expected){throw ('Reopened preference differs: '+$control.Name)}
                        Invoke-Bridge @{Command='set_vbe_option';Pane=$option.Pane;Property=$option.Property;Value=$option.Original;ExpectedOptionsVersion=$after.OptionsVersion}|Out-Null
                        $restoredRead=Invoke-Bridge @{Command='read_vbe_options'}
                        if($restoredRead.OptionsVersion -ne $before.OptionsVersion){throw ('Full preference state not restored: '+$control.Name)}
                        $optionChanged=$false;$optionMatrix+=@{Tab=$tab.Tab;Property=$control.Name;Value=$value;State='PASS';Reopened=$true;FullStateRestored=$true}
                    }
                }
            }
            $report.OptionMatrix=$optionMatrix
            $optionMatrix|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $outputRoot 'option-matrix.json') -Encoding UTF8
            $bars=Invoke-Bridge @{Command='list_toolbars'};$created=Invoke-Bridge @{Command='create_toolbar';ObjectName=('Persistence-'+[Guid]::NewGuid().ToString('N'));Temporary=$false;ExpectedToolbarCollectionVersion=$bars.ToolbarCollectionVersion};$toolbarName=$created.ObjectName
            $commands=Invoke-Bridge @{Command='toolbar_controls';ObjectName=$toolbarName};$native=$vbe.CommandBars.FindControl(1,186)
            $added=Invoke-Bridge @{Command='add_toolbar_command';ObjectName=$toolbarName;ControlId=186;ControlCaption=$native.Caption;Temporary=$false;ExpectedToolbarControlsVersion=$commands.ToolbarControlsVersion};if(-not $added.Verified){throw ('Persistent button creation was not verified: '+($added|ConvertTo-Json -Depth 12 -Compress))}
            $bars=Invoke-Bridge @{Command='list_toolbars'};$layout=@($bars.Toolbars|Where-Object { $_.Properties.Name -eq $toolbarName })
            if($layout.Count -ne 1){throw 'Owned toolbar layout is ambiguous.'}
            $floating=Invoke-Bridge @{Command='set_toolbar_position';ObjectName=$toolbarName;Action='float';ExpectedToolbarLayoutVersion=$layout[0].ToolbarLayoutVersion}
            if(-not $floating.Verified){throw 'Toolbar floating mode was not verified.'}
            $bars=Invoke-Bridge @{Command='list_toolbars'};$layout=@($bars.Toolbars|Where-Object { $_.Properties.Name -eq $toolbarName })
            $placed=Invoke-Bridge @{Command='set_toolbar_placement';ObjectName=$toolbarName;Action='float';ToolbarLeft=320;ToolbarTop=220;ExpectedToolbarLayoutVersion=$layout[0].ToolbarLayoutVersion}
            if(-not $placed.Verified){throw 'Toolbar floating coordinates were not verified.'}
            $report.ToolbarPlacement=$placed
            $report.ToolbarBefore=$added;Close-ProbeExcel;$report.RestartIdentity=Open-ProbeExcel;$report.RestartSourceCaption=$vbe.CommandBars.FindControl(1,186).Caption
            $persisted=Invoke-Bridge @{Command='toolbar_controls';ObjectName=$toolbarName};$report.RestartToolbar=$persisted;$report.RestartToolbarList=Invoke-Bridge @{Command='list_toolbars'};if(@($persisted.Controls).Count -ne 1 -or $persisted.Controls[0].Tag -ne $added.Tag -or $persisted.Controls[0].Id -ne 186){throw 'Toolbar/button ownership did not persist across native restart.'}
            $restartLayout=@($report.RestartToolbarList.Toolbars|Where-Object { $_.Properties.Name -eq $toolbarName })
            if($restartLayout.Count -ne 1 -or $restartLayout[0].Geometry.Position -ne 4 -or $restartLayout[0].Geometry.Left -ne 320 -or $restartLayout[0].Geometry.Top -ne 220){throw 'Toolbar floating geometry did not persist across restart.'}
            $report.Toolbar=@{Created=$created;Added=$added;AfterRestart=$persisted;PersistenceVerified=$true;FloatingPlacementVerified=$true}
        }catch{$report|ConvertTo-Json -Depth 20|Set-Content -LiteralPath (Join-Path $outputRoot 'partial-results.json') -Encoding UTF8;$_|Out-String|Set-Content -LiteralPath (Join-Path $outputRoot 'failure.txt') -Encoding UTF8;throw}
        finally{
            if($optionChanged -and $null -ne $excel){try{$state=Invoke-Bridge @{Command='read_vbe_options'};Invoke-Bridge @{Command='set_vbe_option';Pane=$option.Pane;Property=$option.Property;Value=$option.Original;ExpectedOptionsVersion=$state.OptionsVersion}|Out-Null;$restoredRead=Invoke-Bridge @{Command='read_vbe_options'};if($restoredRead.OptionsVersion -ne $before.OptionsVersion){throw 'Original preferences were not restored.'}}catch{$cleanupErrors+="Preference restoration failed: $_";Write-Warning $cleanupErrors[-1]}}
            if($null -ne $toolbarName -and $null -ne $vbe){try{
                    $state=Invoke-Bridge @{Command='toolbar_controls';ObjectName=$toolbarName}
                    foreach($button in @($state.Controls)){
                        Invoke-Bridge @{Command='remove_toolbar_command';ObjectName=$toolbarName;ControlId=$button.Id;ControlCaption=$button.Caption;InsertIndex=$button.Index;ExpectedToolbarControlsVersion=$state.ToolbarControlsVersion}|Out-Null
                        $state=Invoke-Bridge @{Command='toolbar_controls';ObjectName=$toolbarName}
                    }
                    $bars=Invoke-Bridge @{Command='list_toolbars'}
                    Invoke-Bridge @{Command='remove_toolbar';ObjectName=$toolbarName;ExpectedToolbarControlsVersion=$state.ToolbarControlsVersion;ExpectedToolbarCollectionVersion=$bars.ToolbarCollectionVersion}|Out-Null
                }catch{$cleanupErrors+="Toolbar cleanup failed: $_";Write-Warning $cleanupErrors[-1]}}
            try{Close-ProbeExcel}catch{$cleanupErrors+="Excel cleanup failed: $_";Write-Warning $cleanupErrors[-1]}
            if($hadAccess){New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value $initialAccess -PropertyType DWord -Force|Out-Null}else{Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue}
            $restoredSecurity=Get-ItemProperty -LiteralPath $securityPath
            if(($null -ne $restoredSecurity.PSObject.Properties['AccessVBOM']) -ne $hadAccess -or ($hadAccess -and $restoredSecurity.AccessVBOM -ne $initialAccess)){$cleanupErrors+='AccessVBOM restoration differs from its original state.'}
            $report.Cleanup=@{Errors=@($cleanupErrors);Verified=($cleanupErrors.Count -eq 0)}
            $report.Cleanup|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $outputRoot 'cleanup.json') -Encoding UTF8
            if($cleanupErrors.Count){throw ($cleanupErrors -join '; ')}
        }
        $report|ConvertTo-Json -Depth 20|Set-Content -LiteralPath (Join-Path $outputRoot 'registered-functional-extensions.json') -Encoding UTF8
        Write-Output 'PASS registered procedure arguments, native option reopening and toolbar persistence.'

    }
}

<#
.SYNOPSIS
Runs one explicit Excel code operation or persistence scenario.
.DESCRIPTION
Bookmarks owns an isolated assembly/Excel/session fixture. Other scenarios attach
to the selected owned disposable Excel process and retain their original unsaved
project, source and revision guards. EditPersistence performs its first SaveAs,
close and reopen checks. Use Windows PowerShell -STA for COM operations. No native
mutation is automatically retried, and each invocation runs one scenario.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidateSet('EditPersistence','Bookmarks','InsertCodeFile','RemoveProcedure','RunSub','SelectCodeRange')][string]$Scenario,
    [ValidateRange(1,2147483647)][int]$HostProcessId,
    [switch]$OwnedDisposableHost,
    [ValidateNotNullOrEmpty()][string]$AssemblyPath,
    [ValidateNotNullOrEmpty()][string]$OutputDirectory,
    [switch]$AllowTemporaryVbaAccess
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../probes/VbeProbe.Common.ps1')
$VbeProbeJsonDepth=8
$VbeProbeRawError=$false
$VbeProbeTraceRequests=$false
if($Scenario -eq 'Bookmarks'){
    $required=@('AssemblyPath','OutputDirectory');$allowed=@('AssemblyPath','OutputDirectory','AllowTemporaryVbaAccess')
}else{
    $required=@('HostProcessId','OwnedDisposableHost');$allowed=@('HostProcessId','OwnedDisposableHost')
}
Assert-VbeProbeParameters -Bound $PSBoundParameters -Required $required -Allowed $allowed

switch ($Scenario) {
    'EditPersistence' {
        $ErrorActionPreference = 'Stop'
        $hosts = @(Get-Process EXCEL -ErrorAction Stop)
        if ($hosts.Count -ne 1 -or $hosts[0].Id -ne $HostProcessId) {
            throw 'Exactly one isolated Excel process with the requested PID is required.'
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
        $workbookPath = Join-Path $env:TEMP ("VBAi-code-persistence-{0}.xlsm" -f $HostProcessId)
        $sourcePath = Join-Path $env:TEMP ("VBAi-code-persistence-{0}.bas" -f $HostProcessId)
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
            $properties = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
            $savedAs = Invoke-Vbe @{ Command = 'save_host_document_as'; Project = $project;
                ExpectedProjectVersion = $properties.Version; Path = $workbookPath }
            if (-not $savedAs.SaveAsInvoked -or $savedAs.HostPath -ne $workbookPath -or
                -not $savedAs.HostSaved -or -not $savedAs.ProjectSaved) {
                throw 'The first Excel SaveAs was not verified.'
            }
            $savedRemove = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $removeName }
            $savedInsert = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $insertName }
            if ($savedRemove.Sha256 -ne $removed.Sha256 -or $savedInsert.Sha256 -ne $inserted.Sha256) {
                throw 'Code changed during Excel SaveAs.'
            }
            $book.Close($false)
            $closedProjectCount = -1
            for ($attempt = 0; $attempt -lt 50; $attempt++) {
                $closedProjects = @(Invoke-Vbe @{ Command = 'list_projects' })
                $closedProjectCount = @($closedProjects | Where-Object {
                        $_.FileName -and [IO.Path]::GetFullPath([string]$_.FileName) -eq $workbookPath
                }).Count
                if ($closedProjectCount -eq 0) { break }
                Start-Sleep -Milliseconds 100
            }
            if ($closedProjectCount -ne 0) {
                throw 'The saved VBA project remained in VBE after its workbook closed.'
            }
            $book = $excel.Workbooks.Open($workbookPath)
            $extraBooks = @($excel.Workbooks | Where-Object {
                    [string]$_.FullName -ne $workbookPath
            })
            foreach ($extra in $extraBooks) {
                if ($extra.Path) { throw 'An unexpected saved Excel workbook is open in the isolated process.' }
                $extra.Close($false)
            }
            $reopenedProjects = @()
            $matchingProjects = @()
            for ($attempt = 0; $attempt -lt 10; $attempt++) {
                $reopenedProjects = @(Invoke-Vbe @{ Command = 'list_projects' })
                $matchingProjects = @($reopenedProjects | Where-Object {
                        $_.FileName -and [IO.Path]::GetFullPath([string]$_.FileName) -eq $workbookPath
                })
                if ($matchingProjects.Count -eq 1) { break }
                Start-Sleep -Milliseconds 100
            }
            if ($matchingProjects.Count -ne 1) {
                throw "The reopened workbook has $($matchingProjects.Count) matching VBA projects out of $($reopenedProjects.Count)."
            }
            $project = $workbookPath
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
                FirstSaveAsVerified = $true; RemovedProcedurePersisted = $true; InsertedFilePersisted = $true;
                RemoveModuleSha256 = $afterRemove.Sha256; InsertModuleSha256 = $afterInsert.Sha256 } | Format-List
        }
        finally {
            try { if ($book -ne $null) { $book.Close($false) } } catch { }
            try { $excel.DisplayAlerts = $true } catch { }
            if (Test-Path -LiteralPath $sourcePath) { Remove-Item -LiteralPath $sourcePath -Force }
            if (Test-Path -LiteralPath $workbookPath) { Remove-Item -LiteralPath $workbookPath -Force }
        }

    }
    'Bookmarks' {
        $ErrorActionPreference = 'Stop'
        . $VbeProbeInitializeAssembly

        try {
            . $VbeProbeOpenAssemblyExcel
            $vbe.MainWindow.Visible = $true
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
            $project = $book.VBProject
            $module = $project.VBComponents.Add(1); $module.Name = 'EditorCommandProbe'
            $module.CodeModule.AddFromString("Public Sub First()`r`n    Debug.Print 1`r`nEnd Sub`r`nPublic Sub Second()`r`n    Debug.Print 2`r`nEnd Sub")
            $pane = $module.CodeModule.CodePane; $pane.Show(); $pane.SetSelection(2, 1, 2, 1)
            $book.SaveAs($document, 52)
            $database = $document + '.sqlite'
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe, $database))
            $code = Invoke-Session @{ Command = 'read_module'; Project = $document; Module = $module.Name }
            $added = Invoke-Session @{ Command = 'code_bookmark'; Project = $project.Name; Module = $module.Name; Action = 'add'; Query = "Entrée été"; ExpectedSha256 = $code.Sha256; StartLine = 2; StartColumn = 5 }
            if ($added.Persistence -ne 'SQLite') { throw 'Bookmark did not use persistent storage.' }
            $book.Close($false); $book = $null
            $book = $excel.Workbooks.Open($document)
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe, $database))
            $listed = Invoke-Session @{ Command = 'code_bookmark'; Project = $document; Action = 'list' }
            if ($listed.Bookmarks.Count -ne 1) { throw 'Bookmark not found in new session.' }
            $null = Invoke-Session @{ Command = 'code_bookmark'; Project = $document; Action = 'go'; Query = 'Entrée été' }
            $pane = $vbe.ActiveCodePane
            [int]$line = 0; [int]$col = 0; [int]$endLine = 0; [int]$endCol = 0
            $pane.GetSelection([ref]$line,[ref]$col,[ref]$endLine,[ref]$endCol)
            if ($line -ne 2 -or $col -ne 5 -or $pane.CodeModule.Parent.Name -ne 'EditorCommandProbe') { throw 'Bookmark navigated to a different position.' }
            $otherModule = $book.VBProject.VBComponents.Add(1); $otherModule.Name = 'NavigationTarget'
            $otherModule.CodeModule.AddFromString("Public Sub Target()`r`n    Debug.Print 3`r`nEnd Sub")
            $otherCode = Invoke-Session @{ Command = 'read_module'; Project = $document; Module = $otherModule.Name }
            $pane.Show(); $pane.SetSelection(2,5,2,5)
            $null = Invoke-Session @{ Command = 'navigate_code'; Action = 'go'; Project = $book.VBProject.Name; Module = $otherModule.Name; ExpectedSha256 = $otherCode.Sha256; StartLine = 2; StartColumn = 1 }
            if ($vbe.ActiveCodePane.CodeModule.Parent.Name -ne 'NavigationTarget') { throw 'Navigation did not reach the second module.' }
            $null = Invoke-Session @{ Command = 'navigate_code'; Action = 'back'; Project = $book.VBProject.Name }
            if ($vbe.ActiveCodePane.CodeModule.Parent.Name -ne 'EditorCommandProbe') { throw 'Back by project name did not return to the source module.' }
            $vbe.ActiveCodePane.GetSelection([ref]$line,[ref]$col,[ref]$endLine,[ref]$endCol)
            if ($line -ne 2 -or $col -ne 5) { throw 'Back did not restore the exact position.' }
            $null = Invoke-Session @{ Command = 'navigate_code'; Action = 'forward'; Project = $document }
            if ($vbe.ActiveCodePane.CodeModule.Parent.Name -ne 'NavigationTarget') { throw 'Forward by path did not restore the target module.' }
            $null = Invoke-Session @{ Command = 'navigate_code'; Action = 'back'; Project = $document }
            $null = Invoke-Session @{ Command = 'navigate_code'; Action = 'go'; Project = $document; Module = 'EditorCommandProbe'; ExpectedSha256 = $code.Sha256; StartLine = 3; StartColumn = 1 }
            $forwardCleared = $false
            try { $null = Invoke-Session @{ Command = 'navigate_code'; Action = 'forward'; Project = $document } }
            catch { if ($_.Exception.Message.Contains('No forward')) { $forwardCleared = $true } else { throw } }
            if (-not $forwardCleared) { throw 'New navigation retained obsolete forward history.' }
            $pane = $vbe.ActiveCodePane
            $pane.CodeModule.InsertLines(1, "' external edit")
            $staleRejected = $false
            try { $null = Invoke-Session @{ Command = 'code_bookmark'; Project = $document; Action = 'go'; Query = 'Entrée été' } }
            catch { if ($_.Exception.Message.Contains('stale')) { $staleRejected = $true } else { throw } }
            if (-not $staleRejected) { throw 'Changed source bookmark was accepted.' }
            $backRejected = $false
            try { $null = Invoke-Session @{ Command = 'navigate_code'; Project = $document; Action = 'back' } }
            catch { if ($_.Exception.Message.Contains('stale')) { $backRejected = $true } else { throw } }
            if (-not $backRejected) { throw 'Back accepted changed source.' }
            $pane.CodeModule.DeleteLines(1,1)
            $null = Invoke-Session @{ Command = 'navigate_code'; Project = $document; Action = 'back' }
            $vbe.ActiveCodePane.GetSelection([ref]$line,[ref]$col,[ref]$endLine,[ref]$endCol)
            if ($line -ne 2 -or $col -ne 5) { throw 'Refused back consumed the history entry.' }

            [pscustomobject]@{ Document = $document; Database = $database; Persistence = $added.Persistence; ReopenedInNewSession = $true; Selection = "$line`:$col"; StaleRejected = $staleRejected; BackForwardVerified = $true; ProjectAliasVerified = $true; ForwardClearedByNewNavigation = $forwardCleared; StaleBackRetained = $backRejected } |
            ConvertTo-Json | Set-Content -LiteralPath ($document + '.proof.json') -Encoding UTF8
            Get-Content -LiteralPath ($document + '.proof.json')
        } finally {
            # Restore trust before Quit, which can block in some COM teardown paths.
            if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
            else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
            if ($null -ne $otherBook) { $otherBook.Close($false) }
            if ($null -ne $book) { $book.Close($false) }
            if ($null -ne $excel) { $excel.Quit() }
            if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) {
                # This process was created after the no-Excel precondition and contains only
                # the disposable workbook. Quit can leave it alive because of COM references.
                Stop-Process -Id $probeProcess.Id -Force
            }
        }

    }
    'InsertCodeFile' {
        $ErrorActionPreference = 'Stop'
        if ((Get-Process -Id $HostProcessId -ErrorAction Stop).ProcessName -ne 'EXCEL') {
            throw 'An isolated Excel process is required.'
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

    }
    'RemoveProcedure' {
        $ErrorActionPreference = 'Stop'
        if ((Get-Process -Id $HostProcessId -ErrorAction Stop).ProcessName -ne 'EXCEL') {
            throw 'An isolated Excel process is required.'
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

    }
    'RunSub' {
        $ErrorActionPreference = 'Stop'
        $hosts = @(Get-Process EXCEL -ErrorAction Stop)
        if ($hosts.Count -ne 1 -or $hosts[0].Id -ne $HostProcessId) {
            throw 'Exactly one isolated Excel process is required.'
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

    }
    'SelectCodeRange' {
        $ErrorActionPreference = 'Stop'
        if ((Get-Process -Id $HostProcessId -ErrorAction Stop).ProcessName -ne 'EXCEL') {
            throw 'An isolated Excel process is required.'
        }





        $projects = @(Invoke-Vbe @{ Command = 'list_projects' })
        if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2 -or $projects[0].FileName) {
            throw 'Use one unsaved disposable design-mode VBA project.'
        }
        $project = [string]$projects[0].Name
        $module = 'CodexRangeProbe'
        $created = Invoke-Vbe @{ Command = 'create_module'; Project = $project; Module = $module; ExpectedMode = 2 }
        $source = "Option Explicit`r`nPublic Sub Sample()`r`n    Debug.Print 42`r`nEnd Sub"
        Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
            ExpectedSha256 = $created.Sha256; StartLine = 1; Count = $created.Lines; Text = $source } | Out-Null
        $before = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
        $selected = Invoke-Vbe @{ Command = 'select_code_range'; Project = $project; Module = $module;
            ExpectedSha256 = $before.Sha256; StartLine = 2; StartColumn = 1; EndLine = 3; EndColumn = 11 }
        if (-not $selected.Verified -or $selected.StartLine -ne 2 -or $selected.EndLine -ne 3) {
            throw 'The native multi-line selection was not verified.'
        }
        $pane = Invoke-Vbe @{ Command = 'code_panes' }
        $selection = $pane.ActiveCodePane.Properties.Selection
        if ($selection.StartLine -ne 2 -or $selection.StartColumn -ne 1 -or
            $selection.EndLine -ne 3 -or $selection.EndColumn -ne 11) {
            throw 'A separate CodePanes read did not retain the selection.'
        }
        $invalid = Invoke-VbeRaw @{ Command = 'select_code_range'; Project = $project; Module = $module;
            ExpectedSha256 = $before.Sha256; StartLine = 3; StartColumn = 1; EndLine = 2; EndColumn = 1 }
        if ($invalid.Ok) { throw 'A reversed code range was accepted.' }
        $stale = Invoke-VbeRaw @{ Command = 'select_code_range'; Project = $project; Module = $module;
            ExpectedSha256 = ('0' * 64); StartLine = 2; StartColumn = 1; EndLine = 3; EndColumn = 11 }
        if ($stale.Ok) { throw 'A stale module SHA was accepted.' }
        [pscustomobject]@{ HostProcessId = $HostProcessId; ModuleSha256 = $before.Sha256;
            Verified = $selected.Verified; SeparateReadVerified = $true;
            ReversedRangeRejected = -not $invalid.Ok; StaleShaRejected = -not $stale.Ok } | Format-List

    }
}

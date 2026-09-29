param(
    [Parameter(Mandatory = $true)][string]$AssemblyPath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [switch]$AllowTemporaryVbaAccess
)
$ErrorActionPreference = 'Stop'
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close existing Excel instances before this isolated test.' }
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$directory = [IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($directory) | Out-Null
$document = Join-Path $directory ('Vbai-EditorProbe-' + [Guid]::NewGuid().ToString('N') + '.xlsm')
$securityPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$prior = Get-ItemProperty -LiteralPath $securityPath -ErrorAction Stop
$hadAccess = $null -ne $prior.PSObject.Properties['AccessVBOM']
$priorAccess = $prior.AccessVBOM
if ($priorAccess -ne 1 -and -not $AllowTemporaryVbaAccess) { throw 'Explicit -AllowTemporaryVbaAccess is required to temporarily enable trusted VBA access.' }
$probeProcess = $null
$otherBook = $null
$excel = $null; $book = $null; $form = $null; $session = $null
function Invoke-Session([hashtable]$Fields) {
    $request = New-Object VBAi.Request
    foreach ($key in $Fields.Keys) { $request.$key = $Fields[$key] }
    $response = $script:session.Execute($request)
    if (-not $response.Ok) { throw $response.Error }
    return $response.Data
}
try {
    if ($priorAccess -ne 1) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null }
    $excel = New-Object -ComObject Excel.Application
    $probeProcess = Get-Process EXCEL -ErrorAction Stop
    if (@($probeProcess).Count -ne 1) { throw "Excel isolation was lost." }
    $excel.Visible = $true
    $book = $excel.Workbooks.Add()
    $vbe = $excel.GetType().InvokeMember('VBE', [Reflection.BindingFlags]::GetProperty, $null, $excel, $null)
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

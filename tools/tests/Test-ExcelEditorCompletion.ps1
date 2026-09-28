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
    $request = New-Object CodexVBE.Request
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
    $session = [Activator]::CreateInstance($assembly.GetType('CodexVBE.VbeSession', $true), [object[]]@($vbe))
    $sourceWindow = @($vbe.Windows | Where-Object { $_.Type -eq 5 })[0]
    $oldFrame = $sourceWindow.LinkedWindowFrame
    if ($null -eq $oldFrame -or $oldFrame.LinkedWindows.Count -lt 2) { throw 'Window probe needs an existing shared frame to restore safely.' }
    $originalMembers = @($oldFrame.LinkedWindows | ForEach-Object { "$($_.Type):$($_.Caption)" } | Sort-Object)
    $anchor = @($oldFrame.LinkedWindows | Where-Object { $_.Caption -ne $sourceWindow.Caption -and $_.Type -ne 11 -and $_.Type -ne 12 })[0]
    if ($null -eq $anchor) { throw 'No stable peer for restoring the window probe.' }
    try {
        $layoutBefore = Invoke-Session @{ Command = 'window_layout'; WindowCaption = $sourceWindow.Caption; WindowType = $sourceWindow.Type }
        $detached = Invoke-Session @{ Command = 'link_vbe_window'; WindowCaption = $sourceWindow.Caption; WindowType = $sourceWindow.Type; ExpectedWindowVersion = $layoutBefore.WindowVersion; Action = 'unlink' }
        if (-not $detached.Verified) { throw ('Native window detachment was not verified: ' + ($detached | ConvertTo-Json -Depth 10 -Compress)) }
        $sourceLayout = Invoke-Session @{ Command = 'window_layout'; WindowCaption = $sourceWindow.Caption; WindowType = $sourceWindow.Type }
        $targetLayout = Invoke-Session @{ Command = 'window_layout'; WindowCaption = $anchor.Caption; WindowType = $anchor.Type }
        $reattached = Invoke-Session @{ Command = 'link_vbe_window'; WindowCaption = $sourceWindow.Caption; WindowType = $sourceWindow.Type; ExpectedWindowVersion = $sourceLayout.WindowVersion; Action = 'link'; TargetWindowCaption = $anchor.Caption; TargetWindowType = $anchor.Type; ExpectedTargetWindowVersion = $targetLayout.WindowVersion }
        if (-not $reattached.Verified) { throw 'Native reattachment was not verified.' }
    } finally {
        $oldFrame.LinkedWindows.Add($sourceWindow)
        $restoredMembers = @($oldFrame.LinkedWindows | ForEach-Object { "$($_.Type):$($_.Caption)" } | Sort-Object)
        if (@(Compare-Object $originalMembers $restoredMembers).Count) { throw 'Original linked-window membership was not restored.' }
    }
    $project = $book.VBProject
    $form = $project.VBComponents.Add(3)
    $form.Name = 'VbaiProbeForm'
    $frame = $form.Designer.Controls.Add('Forms.Frame.1', 'Frame1', $true)
    $frame.Width = 200; $frame.Height = 150
    $list = $frame.Controls.Add('Forms.ListBox.1', 'ProbeList', $true)
    $list.ColumnCount = 2
    $module = $project.VBComponents.Add(1)
    $module.Name = 'VbaiProbeModule'
    $module.CodeModule.AddFromString(@"
Option Explicit
Public Function CheckGeneratedList() As String
    Load VbaiProbeForm
    With VbaiProbeForm.Controls("Frame1").Controls("ProbeList")
        CheckGeneratedList = CStr(.ListCount) & "|" & .List(0, 0) & "|" & .List(0, 1) & "|" & .List(1, 0) & "|" & .List(1, 1)
    End With
    Unload VbaiProbeForm
End Function
"@)
    $tree = Invoke-Session @{ Command = 'form_tree'; Project = $project.Name; Form = $form.Name }
    $code = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $form.Name }
    $initialized = Invoke-Session @{ Command = 'set_form_list_initializer'; Project = $project.Name; Form = $form.Name; ControlPath = 'Controls/Frame1/Controls/ProbeList'; ExpectedTreeVersion = $tree.TreeVersion; ExpectedSha256 = $code.Sha256; Rows = [string[][]]@(@('first','a"b'), @('second','last')) }
    if (-not $initialized.Verified) { throw ('Initializer verification failed: ' + ($initialized | ConvertTo-Json -Depth 6 -Compress)) }
    $buttonA = $frame.Controls.Add('Forms.CommandButton.1', 'ButtonA', $true)
    $buttonB = $frame.Controls.Add('Forms.CommandButton.1', 'ButtonB', $true)
    $list.Left = 5; $list.Top = 5; $list.Width = 80; $list.Height = 40
    $buttonA.Left = 40; $buttonA.Top = 70; $buttonA.Width = 60; $buttonA.Height = 20
    $buttonB.Left = 10; $buttonB.Top = 100; $buttonB.Width = 60; $buttonB.Height = 20
    $tree = Invoke-Session @{ Command = 'form_tree'; Project = $project.Name; Form = $form.Name }
    $layout = Invoke-Session @{ Command = 'apply_form_layout'; Project = $project.Name; Form = $form.Name; ExpectedTreeVersion = $tree.TreeVersion; Items = [string[]]@('Controls/Frame1/Controls/ButtonA', 'Controls/Frame1/Controls/ButtonB'); Action = 'align_left' }
    if (-not $layout.Verified -or $buttonB.Left -ne $buttonA.Left) { throw 'Native group alignment was not retained.' }
    $tree = Invoke-Session @{ Command = 'form_tree'; Project = $project.Name; Form = $form.Name }
    $tabOrder = Invoke-Session @{ Command = 'set_form_tab_order'; Project = $project.Name; Form = $form.Name; ExpectedTreeVersion = $tree.TreeVersion; ParentPath = 'Controls/Frame1'; Items = [string[]]@('ButtonB', 'ProbeList', 'ButtonA') }
    if (-not $tabOrder.Verified -or $buttonB.TabIndex -ne 0 -or $list.TabIndex -ne 1 -or $buttonA.TabIndex -ne 2) { throw 'Native tab order was not retained.' }
    $original = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $module.Name }
    $splitCommand = @(Invoke-Session @{ Command = 'list_commands'; Query = 'Fractionner' } | Where-Object { $_.Id -eq 302 -and $_.Enabled })[0]
    if ($null -eq $splitCommand) { throw 'Expected native split command missing from normalized menu search.' }
    $splitRequest = @{ Command = 'set_code_split'; Project = $project.Name; Module = $module.Name; ExpectedSha256 = $original.Sha256; StartLine = 2; ExpectedMode = 2; Action = 'split'; ControlCaption = $splitCommand.Caption }
    $splitResult = Invoke-Session $splitRequest
    if (-not $splitResult.Verified -or $splitResult.PaneCount -ne 2) { throw 'Native split was not verified.' }
    $splitAgain = Invoke-Session $splitRequest
    if ($splitAgain.Applied -or -not $splitAgain.Verified) { throw 'Repeated split was not idempotent.' }
    $splitRequest.Action = 'unsplit'
    $unsplit = Invoke-Session $splitRequest
    if (-not $unsplit.Verified -or $unsplit.PaneCount -ne 1) { throw 'Native unsplit was not verified.' }
    $edited = Invoke-Session @{ Command = 'apply_code_edit'; Project = $project.Name; Module = $module.Name; ExpectedSha256 = $original.Sha256; Action = 'comment'; StartLine = 1; Count = 1 }
    $undone = Invoke-Session @{ Command = 'undo_code_edit'; Project = $project.Name; Module = $module.Name; ExpectedSha256 = $edited.Sha256 }
    if ($undone.Sha256 -ne $original.Sha256) { throw 'Native undo did not restore the source hash.' }
    $redone = Invoke-Session @{ Command = 'redo_code_edit'; Project = $project.Name; Module = $module.Name; ExpectedSha256 = $undone.Sha256 }
    if ($redone.Sha256 -ne $edited.Sha256) { throw 'Native redo did not restore the edited hash.' }
    $restored = Invoke-Session @{ Command = 'undo_code_edit'; Project = $project.Name; Module = $module.Name; ExpectedSha256 = $redone.Sha256 }
    if ($restored.Sha256 -ne $original.Sha256) { throw 'Probe code restoration failed.' }
    $symbols = Invoke-Session @{ Command = 'project_symbols'; Project = $project.Name; Query = 'CheckGeneratedList'; WholeWord = $true }
    if ($symbols.Total -ne 1) { throw 'Native procedure definition search failed.' }
    $navigation = Invoke-Session @{ Command = 'navigate_code'; Project = $project.Name; Module = $module.Name; ExpectedSha256 = $restored.Sha256; Action = 'go'; StartLine = $symbols.Symbols[0].Line; StartColumn = 1 }
    $pageForm = $project.VBComponents.Add(3)
    $pageForm.Name = 'VbaiPageProbe'
    $multi = $pageForm.Designer.Controls.Add('Forms.MultiPage.1', 'MultiPage1', $true)
    $multi.Width = 240; $multi.Height = 180
    $page = $multi.Pages.Item(0); $page.Name = 'DataPage'
    $pageList = $page.Controls.Add('Forms.ComboBox.1', 'PageList', $true)
    $pageList.ColumnCount = 2
    $pageTree = Invoke-Session @{ Command = 'form_tree'; Project = $project.Name; Form = $pageForm.Name }
    $pageCode = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $pageForm.Name }
    $pageInit = Invoke-Session @{ Command = 'set_form_list_initializer'; Project = $project.Name; Form = $pageForm.Name; ControlPath = 'Controls/MultiPage1/Pages/DataPage/Controls/PageList'; ExpectedTreeVersion = $pageTree.TreeVersion; ExpectedSha256 = $pageCode.Sha256; Rows = [string[][]](, [string[]]@('page row', 'page value')) }
    if (-not $pageInit.Verified) { throw 'Page initializer code was not verified.' }
    $module.CodeModule.AddFromString(@"
Public Function CheckPageList() As String
    Load VbaiPageProbe
    With VbaiPageProbe.Controls("MultiPage1").Pages("DataPage").Controls("PageList")
        CheckPageList = CStr(.ListCount) & "|" & .List(0, 0) & "|" & .List(0, 1)
    End With
    Unload VbaiPageProbe
End Function
"@)
    $pageBefore = [string]$excel.Run('CheckPageList')
    if ($pageBefore -ne '1|page row|page value') { throw "Incorrect MultiPage list: $pageBefore" }
    $pageGenerated = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $pageForm.Name }
    $before = [string]$excel.Run('CheckGeneratedList')
    if ($before -ne '2|first|a"b|second|last') { throw "Unexpected runtime list: $before" }
    $generated = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $form.Name }
    $book.SaveAs($document, 52)
    $book.Close($false); $book = $null
    $book = $excel.Workbooks.Open($document)
    $after = [string]$excel.Run("'" + $book.Name + "'!CheckGeneratedList")
    $pageAfter = [string]$excel.Run("'" + $book.Name + "'!CheckPageList")
    $pagePersisted = Invoke-Session @{ Command = 'read_module'; Project = $document; Module = 'VbaiPageProbe' }
    if ($pageAfter -ne $pageBefore -or $pagePersisted.Sha256 -ne $pageGenerated.Sha256) { throw 'MultiPage list or code differs after reopen.' }
    $reopenedFrame = $book.VBProject.VBComponents.Item('VbaiProbeForm').Designer.Controls.Item('Frame1')
    if ($reopenedFrame.Controls.Item('ButtonA').Left -ne 40 -or $reopenedFrame.Controls.Item('ButtonB').Left -ne 40 -or $reopenedFrame.Controls.Item('ButtonB').TabIndex -ne 0 -or $reopenedFrame.Controls.Item('ProbeList').TabIndex -ne 1 -or $reopenedFrame.Controls.Item('ButtonA').TabIndex -ne 2) { throw 'Control layout or tab order differs after reopen.' }
    $persisted = Invoke-Session @{ Command = 'read_module'; Project = $document; Module = 'VbaiProbeForm' }
    if ($after -ne $before -or $persisted.Sha256 -ne $generated.Sha256) { throw 'List runtime or form code differs after reopen.' }
    # Convert the same protected list block to a worksheet binding in the macro workbook.
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = "Source 'quoted' data"
    $sheet.Range('A1').Value2 = 'bound'; $sheet.Range('B1').Value2 = 'column'
    $sheet.Range('A2').Value2 = 'row2'; $sheet.Range('B2').Value2 = 'value2'
    $tree = Invoke-Session @{ Command = 'form_tree'; Project = $document; Form = 'VbaiProbeForm' }
    $code = Invoke-Session @{ Command = 'read_module'; Project = $document; Module = 'VbaiProbeForm' }
    $bound = Invoke-Session @{ Command = 'set_form_list_binding'; Project = $document; Form = 'VbaiProbeForm'; ControlPath = 'Controls/Frame1/Controls/ProbeList'; ExpectedHostPath = $document; ExpectedTreeVersion = $tree.TreeVersion; ExpectedSha256 = $code.Sha256; SheetName = $sheet.Name; RangeAddress = 'A1:B2' }
    if (-not $bound.Verified) { throw 'Binding source generation was not verified.' }
    $otherBook = $excel.Workbooks.Add()
    $otherBook.Worksheets.Item(1).Name = $sheet.Name
    $otherBook.Worksheets.Item(1).Range('A1').Value2 = 'WRONG WORKBOOK'
    $boundBefore = [string]$excel.Run("'" + $book.Name + "'!CheckGeneratedList")
    if ($boundBefore -ne '2|bound|column|row2|value2') { throw "Binding used incorrect source: $boundBefore" }
    $otherBook.Close($false); $otherBook = $null
    $sheet.Range('A1').Value2 = 'changed'
    $changed = [string]$excel.Run("'" + $book.Name + "'!CheckGeneratedList")
    if ($changed -ne '2|changed|column|row2|value2') { throw 'Binding did not reflect the source cell update.' }
    $boundCode = Invoke-Session @{ Command = 'read_module'; Project = $document; Module = 'VbaiProbeForm' }
    $book.Save(); $book.Close($false); $book = $null
    $book = $excel.Workbooks.Open($document)
    $boundAfter = [string]$excel.Run("'" + $book.Name + "'!CheckGeneratedList")
    $boundPersisted = Invoke-Session @{ Command = 'read_module'; Project = $document; Module = 'VbaiProbeForm' }
    if ($boundAfter -ne $changed -or $boundPersisted.Sha256 -ne $boundCode.Sha256) { throw 'Bound list differs after reopen.' }
    $result = [pscustomobject]@{ ExcelVersion = $excel.Version; Document = $document; NativeInitializerVerified = $true; BeforeSave = $before; AfterReopen = $after; FormSha256 = $persisted.Sha256; PersistenceVerified = $true; WindowLinkageVerified = $reattached.Verified; LayoutPersistenceVerified = $true; AlignmentVerified = $layout.Verified; TabOrderVerified = $tabOrder.Verified; UndoRedoVerified = $true; DefinitionNavigationVerified = $true; BoundListVerified = $true; BoundListBeforeSave = $boundBefore; BoundListAfterReopen = $boundAfter; BoundFormSha256 = $boundPersisted.Sha256; CrossWorkbookIsolationVerified = $true; MultiPageListPersistenceVerified = $true; MultiPageListAfterReopen = $pageAfter; MultiPageFormSha256 = $pagePersisted.Sha256; CodeSplitVerified = $true; CodeSplitIdempotenceVerified = $true; Invocation = 'Current build VbeSession against live Excel COM' }
    $result | ConvertTo-Json | Set-Content -LiteralPath ([IO.Path]::ChangeExtension($document, '.proof.json')) -Encoding UTF8
    $result | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'excel-editor-proof.json') -Encoding UTF8
    $result | Format-List
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

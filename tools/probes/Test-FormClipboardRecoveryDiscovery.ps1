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
    Write-Output ("Request: " + $request.Command + " " + $request.Action) | Out-Host
    $response = $script:session.Execute($request)
    if (-not $response.Ok) { throw $response.Error }
    return $response.Data
}
Add-Type -AssemblyName System.Windows.Forms
$backup = New-Object System.Windows.Forms.DataObject
$backupValues = @{}
$currentClipboard = [System.Windows.Forms.Clipboard]::GetDataObject()
if ($null -ne $currentClipboard) {
    foreach ($format in $currentClipboard.GetFormats($false)) {
        $value = $currentClipboard.GetData($format, $false)
        if ($null -eq $value) { continue }
        if ($value -is [IO.MemoryStream]) { $value = New-Object IO.MemoryStream -ArgumentList (,($value.ToArray())) }
        elseif ($value -isnot [string]) { throw "Clipboard test refused: unsupported backup format $format." }
        $backup.SetData($format, $false, $value); $backupValues[$format] = $value
    }
}
$clipboardTouched = $false
$proof = $null
try {
    if ($priorAccess -ne 1) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null }
    $excel = New-Object -ComObject Excel.Application
    $probeProcess = Get-Process EXCEL -ErrorAction Stop
    if (@($probeProcess).Count -ne 1) { throw "Excel isolation was lost." }
    $excel.Visible = $true
    $book = $excel.Workbooks.Add()
    $vbe = $excel.GetType().InvokeMember('VBE', [Reflection.BindingFlags]::GetProperty, $null, $excel, $null)
    $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
    $project = $book.VBProject
    $vbe.MainWindow.Visible = $true
    $form = $project.VBComponents.Add(3); $form.Name = 'ClipboardDesignerSource'
    $button = $form.Designer.Controls.Add('Forms.CommandButton.1','SourceButton',$true)
    $form.CodeModule.AddFromString("Private Sub SourceButton_Click()`r`n    Debug.Print ""handler retained""`r`nEnd Sub")
    $sourceCode = $form.CodeModule.Lines(1,$form.CodeModule.CountOfLines)
    $button.Caption = 'Clipboard native test'; $button.Left = 30; $button.Top = 30
    $form.DesignerWindow().Visible = $true; $form.DesignerWindow().SetFocus()
    function Pump-Designer { for ($i=0; $i -lt 5; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 50 } }
    Pump-Designer
    function Read-DesignerState($name,$parent) { Invoke-Session @{ Command='form_clipboard_state'; Project=$project.Name; Form=$name; ParentPath=$parent } }
    function Invoke-DesignerClipboard($name,$action,$state) {
        Invoke-Session @{ Command='native_form_clipboard'; Project=$project.Name; Form=$name; Action=$action; ExpectedDesignerSelectionVersion=$state.SelectionVersion; ExpectedClipboardVersion=$state.ClipboardVersion; ParentPath=$state.ParentPath }
    }
    function Select-Designer($name,$parent,[string[]]$names,$state) {
        Invoke-Session @{ Command='select_form_controls'; Project=$project.Name; Form=$name; ParentPath=$parent; Items=$names; ExpectedDesignerSelectionVersion=$state.SelectionVersion }
    }
    $initialSelection = Select-Designer $form.Name '' @('SourceButton') (Read-DesignerState $form.Name)
    if (-not $initialSelection.Verified) { throw 'Initial programmatic selection failed.' }
    $source = Read-DesignerState $form.Name
    if (@($source.Selected).Count -ne 1 -or $source.Selected[0] -ne 'SourceButton') { throw 'Designer selection not read correctly.' }
    $button.Caption = 'Changed caption'
    $staleTree = $false
    try { Invoke-DesignerClipboard $form.Name 'copy' $source | Out-Null }
    catch { $staleTree = $_.Exception.Message -match 'tree or selection changed' }
    $source = Read-DesignerState $form.Name
    $clipboardTouched = $true
    [System.Windows.Forms.Clipboard]::SetText('Temporary clipboard guard test')
    $staleClipboard = $false
    try { Invoke-DesignerClipboard $form.Name 'copy' $source | Out-Null }
    catch { $staleClipboard = $_.Exception.Message -match 'Clipboard changed' -or $_.Exception.Message -match 'tree or selection changed' }
    $copy = Invoke-DesignerClipboard $form.Name 'copy' (Read-DesignerState $form.Name)
    $detachedData = New-Object System.Windows.Forms.DataObject
    $formatInfo = @()
    $nativeData = [System.Windows.Forms.Clipboard]::GetDataObject()
    foreach ($format in $nativeData.GetFormats($false)) {
        $value = $nativeData.GetData($format,$false)
        $formatInfo += [pscustomobject]@{ Format=$format; Type=$(if ($null -eq $value) { "null" } else { $value.GetType().FullName }) }
        if ($null -eq $value) { continue }
        if ($value -is [IO.MemoryStream]) { $value=New-Object IO.MemoryStream -ArgumentList (,($value.ToArray())) }
        elseif ($value -isnot [string]) { throw "Unsupported native clipboard format type: $format" }
        $detachedData.SetData($format,$false,$value)
    }
    $formatInfo | ConvertTo-Json | Set-Content (Join-Path $directory 'native-form-clipboard-formats.json') -Encoding UTF8
    [System.Windows.Forms.Clipboard]::SetText('Clipboard overwritten after native copy')
    [System.Windows.Forms.Clipboard]::SetDataObject($detachedData,$true)
    $formatInfo | ConvertTo-Json | Set-Content (Join-Path $directory 'native-form-clipboard-formats.json') -Encoding UTF8
    $target = $project.VBComponents.Add(3); $target.Name = 'ClipboardDesignerTarget'
    $target.DesignerWindow().Visible = $true
    $form.DesignerWindow().SetFocus(); Pump-Designer
    $paste = Invoke-DesignerClipboard $target.Name 'paste' (Read-DesignerState $target.Name)
    Pump-Designer
    $countAfterPaste = $target.Designer.Controls.Count
    $pastedCaption = $target.Designer.Controls.Item(0).Caption
    $target.DesignerWindow().SetFocus(); Pump-Designer
    $cut = Invoke-DesignerClipboard $form.Name 'cut' (Read-DesignerState $form.Name)
    Pump-Designer
    $countAfterCut = $form.Designer.Controls.Count
    $tree = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
    $undoAvailable = [bool](($tree.Properties | Where-Object Name -eq CanUndo).Value)
    $undo = $null; $undoRefused = $false; $recovery = $null
    if ($undoAvailable) {
        $undo = Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='undo'; ExpectedTreeVersion=$tree.TreeVersion }
    } else {
        try { Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='undo'; ExpectedTreeVersion=$tree.TreeVersion } | Out-Null }
        catch { $undoRefused = $_.Exception.Message -match 'unavailable' }
        if (-not $undoRefused) { throw 'Unavailable undo was not refused.' }
        $recovery = Invoke-DesignerClipboard $form.Name 'paste' (Read-DesignerState $form.Name)
    }
    $restoredCount = $form.Designer.Controls.Count
    $frame = $form.Designer.Controls.Add('Forms.Frame.1','SourceFrame',$true)
    $first = $frame.Controls.Add('Forms.Label.1','FirstNested',$true); $first.Caption='First nested caption'
    $second = $frame.Controls.Add('Forms.CommandButton.1','SecondNested',$true); $second.Caption='Second nested caption'
    $multi = $target.Designer.Controls.Add('Forms.MultiPage.1','TargetPages',$true)
    $page = $multi.Pages.Item(0)
    $sourcePath = 'Controls/SourceFrame'
    $targetPath = 'Controls/TargetPages/Pages/' + $page.Name
    $oneSelection = Select-Designer $form.Name $sourcePath @('FirstNested') (Read-DesignerState $form.Name $sourcePath)
    $previousSelection = Read-DesignerState $form.Name $sourcePath
    $twoSelection = Select-Designer $form.Name $sourcePath @('FirstNested','SecondNested') $previousSelection
    $staleSelection = $false
    try { Invoke-DesignerClipboard $form.Name 'cut' $previousSelection | Out-Null }
    catch { $staleSelection = $_.Exception.Message -match 'tree or selection changed' }
    $unknownRefused = $false
    try { Select-Designer $form.Name $sourcePath @('MissingControl') (Read-DesignerState $form.Name $sourcePath) | Out-Null }
    catch { $unknownRefused = $_.Exception.Message -match 'outside the selected container' }
    $preservedSelection = Read-DesignerState $form.Name $sourcePath
    if (-not $oneSelection.Verified -or -not $twoSelection.Verified -or -not $staleSelection -or -not $unknownRefused -or @($preservedSelection.Selected).Count -ne 2) { throw 'Nested selection guards failed.' }
    $nestedCopy = Invoke-DesignerClipboard $form.Name 'copy' $preservedSelection
    $nestedPaste = Invoke-DesignerClipboard $target.Name 'paste' (Read-DesignerState $target.Name $targetPath)
    $pastedNames = @($page.Controls | ForEach-Object { $_.Name })
    if ($page.Controls.Count -ne 2 -or $page.Controls.Item('FirstNested').Caption -cne 'First nested caption' -or $page.Controls.Item('SecondNested').Caption -cne 'Second nested caption') { throw 'Frame to Page native copy failed.' }
    $nestedCut = Invoke-DesignerClipboard $form.Name 'cut' (Read-DesignerState $form.Name $sourcePath)
    if ($frame.Controls.Count -ne 0) { throw 'Nested cut did not remove both selected controls.' }
    $nestedRecovery = Invoke-DesignerClipboard $form.Name 'paste' (Read-DesignerState $form.Name $sourcePath)
    if ($frame.Controls.Count -ne 2) { throw 'Nested recovery failed.' }
    $clear = Select-Designer $form.Name $sourcePath @() (Read-DesignerState $form.Name $sourcePath)
    if (-not $clear.Verified -or @($clear.After.Selected).Count -ne 0) { throw 'Clear nested selection failed.' }
    if ($form.CodeModule.Lines(1,$form.CodeModule.CountOfLines) -cne $sourceCode -or $target.CodeModule.CountOfLines -ne 0) { throw 'Source event code changed or was unexpectedly copied.' }
    $book.SaveAs($document,52)
    $book.Close($false); $book = $null
    $book = $excel.Workbooks.Open($document)
    $reopenedSource = $book.VBProject.VBComponents.Item('ClipboardDesignerSource')
    $reopenedTarget = $book.VBProject.VBComponents.Item('ClipboardDesignerTarget')
    $savedFrame = $reopenedSource.Designer.Controls.Item('SourceFrame')
    $savedPage = $reopenedTarget.Designer.Controls.Item('TargetPages').Pages.Item(0)
    if ($savedFrame.Controls.Count -ne 2 -or $savedPage.Controls.Count -ne 2 -or $savedPage.Controls.Item('FirstNested').Caption -cne 'First nested caption' -or $savedPage.Controls.Item('SecondNested').Caption -cne 'Second nested caption') { throw 'Saved nested control state was not preserved.' }
    if ($reopenedSource.CodeModule.Lines(1,$reopenedSource.CodeModule.CountOfLines) -cne $sourceCode -or $reopenedTarget.CodeModule.CountOfLines -ne 0) { throw 'Saved event code was not preserved.' }
    $proof = [pscustomobject]@{ SaveReopenVerified=$true; EventCodePreserved=$true; NestedCopy=$nestedCopy; NestedPaste=$nestedPaste; NestedCut=$nestedCut; NestedRecovery=$nestedRecovery; NestedSelectionVerified=$true; StaleSelectionRejected=$staleSelection; UnknownControlRejected=$unknownRefused; Copy=$copy; Paste=$paste; Cut=$cut; Undo=$undo; UndoAvailable=$undoAvailable; UndoRefused=$undoRefused; RecoveryPaste=$recovery; TargetCount=$countAfterPaste; PastedCaption=$pastedCaption; SourceAfterCut=$countAfterCut; SourceAfterRecovery=$restoredCount; StaleTreeRejected=$staleTree; StaleClipboardRejected=$staleClipboard; ClipboardRestored=$false; Mvid=$assembly.ManifestModule.ModuleVersionId.ToString() }
    if ($countAfterPaste -ne 1 -or $pastedCaption -cne 'Changed caption' -or $countAfterCut -ne 0 -or $restoredCount -ne 1 -or -not $staleTree -or -not $staleClipboard -or -not $copy.ClipboardChanged -or -not $paste.DesignerChangeObserved -or -not $cut.DesignerChangeObserved) { throw 'Native Designer clipboard verification failed.' }
} finally {
    try {
        if ($clipboardTouched) {
            if ($backupValues.Count -eq 0) { [System.Windows.Forms.Clipboard]::Clear() }
            else { [System.Windows.Forms.Clipboard]::SetDataObject($backup, $true) }
            $restoredClipboard = [System.Windows.Forms.Clipboard]::GetDataObject()
            foreach ($format in $backupValues.Keys) {
                $expected = $backupValues[$format]; $actual = $restoredClipboard.GetData($format,$false)
                if ($expected -is [IO.MemoryStream]) {
                    if ($actual -isnot [IO.MemoryStream] -or [Convert]::ToBase64String($expected.ToArray()) -cne [Convert]::ToBase64String($actual.ToArray())) { throw "Clipboard format restoration failed: $format" }
                } elseif ($actual -cne $expected) { throw "Clipboard format restoration failed: $format" }
            }
            if ($null -ne $proof) { $proof.ClipboardRestored = $true }
        }
    } finally {
        if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
        else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
        try { if ($null -ne $book) { $book.Close($false) } }
        finally {
            try { if ($null -ne $excel) { $excel.Quit() } }
            finally { if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) { Stop-Process -Id $probeProcess.Id -Force } }
        }
    }
}
$proof | ConvertTo-Json -Depth 40 | Set-Content (Join-Path $directory 'native-form-clipboard-recovery-discovery.json') -Encoding UTF8
$proof | Select-Object SaveReopenVerified,EventCodePreserved,NestedSelectionVerified,StaleSelectionRejected,UnknownControlRejected,TargetCount,PastedCaption,SourceAfterCut,SourceAfterRecovery,StaleTreeRejected,StaleClipboardRejected,ClipboardRestored,UndoAvailable,UndoRefused | ConvertTo-Json

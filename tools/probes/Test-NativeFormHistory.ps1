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
    Write-Output ("Request: " + $request.Command + " " + $request.Action) | Out-Host
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
    $project = $book.VBProject
    $form = $project.VBComponents.Add(3); $form.Name = 'NativeDesignerHistory'
    $form.DesignerWindow().Visible = $true
    $button = $form.Designer.Controls.Add('Forms.CommandButton.1', 'HistoryButton', $true)
    $button.Left = 20; $button.Top = 20; $button.Width = 70; $button.Height = 24
    $button.Left = 55
    $second = $form.Designer.Controls.Add('Forms.CommandButton.1', 'SecondButton', $true)
    $second.Left = 100; $second.Top = 60; $second.Width = 70; $second.Height = 24
    $form.DesignerWindow().SetFocus()
    Add-Type -AssemblyName System.Windows.Forms
    function Pump-Designer { for ($i=0; $i -lt 5; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 50 } }
    Pump-Designer
    $vbe.CommandBars.FindControl(1,756).Execute()
    Pump-Designer
    $original = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
    $align = $vbe.CommandBars.FindControl(1,664)
    if (-not $align.Enabled) { throw 'Native Align Left is disabled.' }
    $align.Execute()
    Pump-Designer
    $before = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
    $commands = Invoke-Session @{ Command='list_commands' }
    $commands | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $directory 'native-form-history-commands.json') -Encoding UTF8
    $otherBook = $excel.Workbooks.Add()
    $otherBook.VBProject.Name = 'OtherDesignerProject'
    $otherForm = $otherBook.VBProject.VBComponents.Add(3); $otherForm.Name = 'OtherDesigner'
    $otherButton = $otherForm.Designer.Controls.Add('Forms.CommandButton.1', 'UntouchedButton', $true)
    $otherButton.Left = 80; $otherButton.Top = 40
    $otherForm.DesignerWindow().Visible = $true; $otherForm.DesignerWindow().SetFocus()
    Pump-Designer
    $otherBefore = Invoke-Session @{ Command='form_tree'; Project='OtherDesignerProject'; Form=$otherForm.Name }
    $before = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
    $nativeResult = Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='undo'; ExpectedTreeVersion=$before.TreeVersion }
    Pump-Designer
    $after = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
    $staleRejected = $false
    try { Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='redo'; ExpectedTreeVersion=$before.TreeVersion } | Out-Null }
    catch { $staleRejected = $_.Exception.Message -match 'changed since' }
    $redoResult = Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='redo'; ExpectedTreeVersion=$after.TreeVersion }
    Pump-Designer
    $redone = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
    $otherAfter = Invoke-Session @{ Command='form_tree'; Project='OtherDesignerProject'; Form=$otherForm.Name }
    $unavailableRejected = $false
    try { Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='redo'; ExpectedTreeVersion=$redone.TreeVersion } | Out-Null }
    catch { $unavailableRejected = $_.Exception.Message -match 'unavailable' }
    $form.DesignerWindow().SetFocus()
    Pump-Designer
    $vbe.CommandBars.FindControl(1,756).Execute()
    Pump-Designer
    $clear = $vbe.CommandBars.FindControl(1,478)
    if (-not $clear.Enabled) { throw 'Native Delete is disabled.' }
    $clear.Execute()
    Pump-Designer
    $deleted = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
    $undoDelete = Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='undo'; ExpectedTreeVersion=$deleted.TreeVersion }
    $restoredControls = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
    [pscustomobject]@{ Deleted=$deleted; UndoDelete=$undoDelete; Restored=$restoredControls } | ConvertTo-Json -Depth 30 | Set-Content (Join-Path $directory 'native-form-delete-stage.json') -Encoding UTF8
    $redoDeletionAvailable = [bool](($restoredControls.Properties | Where-Object Name -eq CanRedo).Value)
    $redoDelete = $null; $redoDeleteRefused = $false
    if ($redoDeletionAvailable) {
        $redoDelete = Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='redo'; ExpectedTreeVersion=$restoredControls.TreeVersion }
    } else {
        try { Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='redo'; ExpectedTreeVersion=$restoredControls.TreeVersion } | Out-Null }
        catch { $redoDeleteRefused = $_.Exception.Message -match 'unavailable' }
        if (-not $redoDeleteRefused) { throw 'Unavailable native redo was not refused.' }
    }
    $deletedAgain = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
    if ($deleted.NodeCount -ne 0 -or $restoredControls.NodeCount -ne 2 -or -not $undoDelete.Verified) { throw 'Native control deletion undo was not verified.' }
    if ($redoDeletionAvailable -and ($deletedAgain.NodeCount -ne 0 -or -not $redoDelete.Verified)) { throw 'Available native redo was not verified.' }
    if (-not $redoDeletionAvailable -and $deletedAgain.NodeCount -ne 2) { throw 'Refused redo changed the controls.' }
    [pscustomobject]@{ RedoDeletionAvailable=$redoDeletionAvailable; RedoDeletionRefused=$redoDeleteRefused; UndoDelete=$undoDelete; RedoDelete=$redoDelete; RestoredControls=$restoredControls.NodeCount; DeletedAgain=$deletedAgain.NodeCount; StaleRejected=$staleRejected; UnavailableRejected=$unavailableRejected; OtherFormPreserved=($otherBefore.TreeVersion -eq $otherAfter.TreeVersion); UndoReturn=$nativeResult; RedoReturn=$redoResult; Redone=$redone; Original=$original; Before=$before; After=$after; Commands=$commands } | ConvertTo-Json -Depth 25 | Set-Content (Join-Path $directory 'native-form-history.json') -Encoding UTF8
    if (-not $nativeResult.Verified -or -not $redoResult.Verified -or -not $staleRejected -or -not $unavailableRejected -or $otherBefore.TreeVersion -ne $otherAfter.TreeVersion) { throw 'Designer history targeting or guards failed.' }
    $leftAfter = ($after.Controls | Where-Object Name -eq SecondButton).Properties | Where-Object Name -eq Left
    $leftRedone = ($redone.Controls | Where-Object Name -eq SecondButton).Properties | Where-Object Name -eq Left
    if ($leftAfter.Value -ne 100 -or $leftRedone.Value -ne 55) { throw 'Designer geometry was not restored.' }
} finally {
    if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
    else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
    try { if ($null -ne $otherBook) { $otherBook.Close($false) }; if ($null -ne $book) { $book.Close($false) } }
    finally {
        try { if ($null -ne $excel) { $excel.Quit() } }
        finally { if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) { Stop-Process -Id $probeProcess.Id -Force } }
    }
}

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
    $module = $project.VBComponents.Add(1); $module.Name = 'NativeHistoryProbe'
    $module.CodeModule.AddFromString("Option Explicit`r`nPublic Sub Example()`r`nEnd Sub")
    $module.CodeModule.CodePane.Show()
    $otherModule = $project.VBComponents.Add(1); $otherModule.Name = 'LastEditedHistory'
    $otherModule.CodeModule.AddFromString("Option Explicit`r`n' Last edit in a different module")
    $before = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
    $otherBefore = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$otherModule.Name }
    $module.CodeModule.CodePane.Show()
    $module.CodeModule.CodePane.Window.SetFocus()
    Add-Type -AssemblyName System.Windows.Forms
    for ($i=0; $i -lt 10; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 50 }
    $settings = [Activator]::CreateInstance($assembly.GetType('CodexVBE.LlmSettings'), $true)
    $settings.VbeEditApproval = 'Automatic'
    $llm = [Activator]::CreateInstance($assembly.GetType('CodexVBE.LlmVbeTools'), [object[]]@($session,$null,$settings))
    $llm.BoundProject = $project.Name
    Add-Type -TypeDefinition 'public static class HistoryDiffCounter { public static int Count; public static string Modules = ""; public static void Changed(object change) { Count++; Modules += change.GetType().GetProperty("Module").GetValue(change, null).ToString() + ";"; } }'
    $editedEvent = $llm.GetType().GetEvent('CodeEdited')
    $editedHandler = [Delegate]::CreateDelegate($editedEvent.EventHandlerType, [HistoryDiffCounter].GetMethod('Changed'))
    $editedEvent.AddEventHandler($llm, $editedHandler)
    function Invoke-HistoryTool([hashtable]$Fields) {
        $Fields.Remove('Command')
        $response = $llm.Invoke('native_code_history', ($Fields | ConvertTo-Json -Compress)) | ConvertFrom-Json
        if (-not $response.Ok) { throw $response.Error }
        return $response.Data
    }
    $state = Invoke-Session @{ Command='native_code_history_state'; Project=$project.Name }
    $undoCaption = ($state.Commands | Where-Object { $_.Id -eq 128 -and $_.Enabled } | Select-Object -First 1).Caption
    $undo = Invoke-HistoryTool @{ Command='native_code_history'; Project=$project.Name; Action='undo'; ExpectedMode=2; ExpectedProjectVersion=$state.HistoryVersion; ControlCaption=$undoCaption }
    $empty = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$otherModule.Name }
    $undoState = Invoke-Session @{ Command='native_code_history_state'; Project=$project.Name }
    $redoCaption = ($undoState.Commands | Where-Object { $_.Id -eq 129 -and $_.Enabled } | Select-Object -First 1).Caption
    $redo = Invoke-HistoryTool @{ Command='native_code_history'; Project=$project.Name; Action='redo'; ExpectedMode=2; ExpectedProjectVersion=$undoState.HistoryVersion; ControlCaption=$redoCaption }
    $restored = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$otherModule.Name }
    $currentState = Invoke-Session @{ Command='native_code_history_state'; Project=$project.Name }
    $undoCaption = ($currentState.Commands | Where-Object { $_.Id -eq 128 -and $_.Enabled } | Select-Object -First 1).Caption
    $staleRejected = $false
    try { Invoke-Session @{ Command='native_code_history'; Project=$project.Name; Action='undo'; ExpectedMode=2; ExpectedProjectVersion=$undoState.HistoryVersion; ControlCaption=$undoCaption } | Out-Null }
    catch { $staleRejected = $_.Exception.Message -match 'changed since' }
    $badCaptionRejected = $false
    try { Invoke-Session @{ Command='native_code_history'; Project=$project.Name; Action='undo'; ExpectedMode=2; ExpectedProjectVersion=$currentState.HistoryVersion; ControlCaption='invalid' } | Out-Null }
    catch { $badCaptionRejected = $_.Exception.Message -match 'absent or disabled' }
    $afterGuards = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$otherModule.Name }
    $after = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
    $otherBook = $excel.Workbooks.Add()
    $otherBook.VBProject.Name = 'OtherHistoryProject'
    $multiRejected = $false
    try { Invoke-Session @{ Command='native_code_history'; Project=$project.Name; Action='undo'; ExpectedMode=2; ExpectedProjectVersion=$currentState.HistoryVersion; ControlCaption=$undoCaption } | Out-Null }
    catch { $multiRejected = $_.Exception.Message -match 'close other projects' }
    $otherBook.Close($false); $otherBook = $null
    $form = $project.VBComponents.Add(3)
    $formRejected = $false
    try { Invoke-Session @{ Command='native_code_history'; Project=$project.Name; Action='undo'; ExpectedMode=2; ExpectedProjectVersion=$currentState.HistoryVersion; ControlCaption=$undoCaption } | Out-Null }
    catch { $formRejected = $_.Exception.Message -match 'containing UserForms' }
    [pscustomobject]@{ ChatDiffCount=[HistoryDiffCounter]::Count; ChatDiffModules=[HistoryDiffCounter]::Modules; Undo=$undo; Redo=$redo; Before=$otherBefore; Empty=$empty; Restored=$restored; StaleRejected=$staleRejected; BadCaptionRejected=$badCaptionRejected; MultipleProjectsRejected=$multiRejected; UserFormRejected=$formRejected; ActiveModulePreserved=($before.Sha256 -eq $after.Sha256); GuardsPreservedCode=($afterGuards.Sha256 -eq $restored.Sha256); Mvid=$assembly.ManifestModule.ModuleVersionId.ToString() } | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $directory 'native-code-history.json') -Encoding UTF8
    if ([HistoryDiffCounter]::Count -ne 2 -or [HistoryDiffCounter]::Modules -ne 'LastEditedHistory;LastEditedHistory;') { throw 'Chat diffs did not match native history changes.' }
    if (-not $undo.Verified -or -not $redo.Verified -or $empty.Code -ne '' -or $restored.Sha256 -ne $otherBefore.Sha256 -or $undo.Changes[0].Module -ne $otherModule.Name -or $before.Sha256 -ne $after.Sha256 -or -not $staleRejected -or -not $badCaptionRejected -or -not $multiRejected -or -not $formRejected -or $afterGuards.Sha256 -ne $restored.Sha256) { throw 'Native shared history cycle or guards failed.' }
} finally {
    if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
    else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
    try { if ($null -ne $otherBook) { $otherBook.Close($false) }; if ($null -ne $book) { $book.Close($false) } }
    finally {
        try { if ($null -ne $excel) { $excel.Quit() } }
        finally { if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) { Stop-Process -Id $probeProcess.Id -Force } }
    }
}

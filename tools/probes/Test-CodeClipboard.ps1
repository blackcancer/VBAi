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
Add-Type -AssemblyName System.Windows.Forms
$backup = New-Object System.Windows.Forms.DataObject
$backupValues = @{}
$currentClipboard = [System.Windows.Forms.Clipboard]::GetDataObject()
if ($null -ne $currentClipboard) {
    foreach ($format in $currentClipboard.GetFormats($false)) {
        $value = $currentClipboard.GetData($format, $false)
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
    $session = [Activator]::CreateInstance($assembly.GetType('CodexVBE.VbeSession', $true), [object[]]@($vbe))
    $project = $book.VBProject
    $module = $project.VBComponents.Add(1); $module.Name = 'ClipboardProbe'
    $module.CodeModule.AddFromString("' alpha beta`r`n' second")
    $before = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
    $settings = [Activator]::CreateInstance($assembly.GetType('CodexVBE.LlmSettings'), $true)
    $settings.VbeEditApproval = 'Automatic'
    $llm = [Activator]::CreateInstance($assembly.GetType('CodexVBE.LlmVbeTools'), [object[]]@($session,$null,$settings))
    $llm.BoundProject = $project.Name
    Add-Type -TypeDefinition 'public static class ClipboardDiffCounter { public static int Count; public static void Changed(object change) { Count++; } }'
    $editedEvent = $llm.GetType().GetEvent('CodeEdited')
    $editedHandler = [Delegate]::CreateDelegate($editedEvent.EventHandlerType, [ClipboardDiffCounter].GetMethod('Changed'))
    $editedEvent.AddEventHandler($llm, $editedHandler)
    function Invoke-ClipboardTool([string]$Name, [hashtable]$Fields) {
        $response = $llm.Invoke($Name, ($Fields | ConvertTo-Json -Compress)) | ConvertFrom-Json
        if (-not $response.Ok) { throw $response.Error }
        return $response.Data
    }
    $range = @{ Project=$project.Name; Module=$module.Name; ExpectedSha256=$before.Sha256; StartLine=1; StartColumn=3; EndLine=1; EndColumn=8 }
    $clipboardTouched = $true
    $copy = Invoke-ClipboardTool 'copy_code' $range
    $readClipboard = Invoke-Session @{ Command='read_code_clipboard' }
    if ($readClipboard.Text -cne 'alpha') { throw 'Native copy did not match the exact selection.' }
    $cut = Invoke-ClipboardTool 'cut_code' $range
    $afterCut = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
    if ($afterCut.Code -cne "'  beta`r`n' second") { throw 'Cut changed unexpected text.' }
    $paste = Invoke-ClipboardTool 'paste_code' @{ Project=$project.Name; Module=$module.Name; ExpectedSha256=$afterCut.Sha256; ExpectedClipboardVersion=$cut.ClipboardVersion; StartLine=2; StartColumn=3; EndLine=2; EndColumn=9 }
    $afterPaste = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
    if ($afterPaste.Code -cne "'  beta`r`n' alpha") { throw 'Paste changed unexpected text.' }
    Invoke-ClipboardTool 'undo_code_edit' @{ Project=$project.Name; Module=$module.Name; ExpectedSha256=$afterPaste.Sha256 } | Out-Null
    $undone = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
    if ($undone.Sha256 -ne $afterCut.Sha256) { throw 'Clipboard edit undo did not restore source.' }
    Invoke-ClipboardTool 'redo_code_edit' @{ Project=$project.Name; Module=$module.Name; ExpectedSha256=$undone.Sha256 } | Out-Null
    $redone = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
    if ($redone.Sha256 -ne $afterPaste.Sha256) { throw 'Clipboard edit redo did not restore source.' }
    [System.Windows.Forms.Clipboard]::SetText("' café`n' suite", [System.Windows.Forms.TextDataFormat]::UnicodeText)
    $staleRejected = $false
    try { Invoke-ClipboardTool 'paste_code' @{ Project=$project.Name; Module=$module.Name; ExpectedSha256=$redone.Sha256; ExpectedClipboardVersion=$cut.ClipboardVersion; StartLine=1; StartColumn=1; EndLine=1; EndColumn=1 } | Out-Null }
    catch { $staleRejected = $_.Exception.Message -match 'Clipboard changed' }
    $emptyModule = $project.VBComponents.Add(1); $emptyModule.Name = 'EmptyClipboardTarget'
    $empty = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$emptyModule.Name }
    $unicode = Invoke-Session @{ Command='read_code_clipboard' }
    $insert = Invoke-ClipboardTool 'paste_code' @{ Project=$project.Name; Module=$emptyModule.Name; ExpectedSha256=$empty.Sha256; ExpectedClipboardVersion=$unicode.Version; StartLine=1; StartColumn=1; EndLine=1; EndColumn=1 }
    $inserted = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$emptyModule.Name }
    if ($inserted.Code -cne "' café`r`n' suite" -or -not $insert.Verified -or -not $staleRejected -or [ClipboardDiffCounter]::Count -ne 5) { throw 'Clipboard native guard, Unicode or chat diff verification failed.' }
    $proof = [pscustomobject]@{ CopyVerified=$true; CutVerified=$cut.Verified; PasteVerified=$paste.Verified; UndoRedoVerified=$true; StaleClipboardRejected=$staleRejected; UnicodeEmptyModuleVerified=$true; ChatDiffCount=[ClipboardDiffCounter]::Count; Mvid=$assembly.ManifestModule.ModuleVersionId.ToString(); ClipboardRestored=$false }
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
$proof | ConvertTo-Json | Set-Content (Join-Path $directory 'native-code-clipboard.json') -Encoding UTF8
$proof | ConvertTo-Json

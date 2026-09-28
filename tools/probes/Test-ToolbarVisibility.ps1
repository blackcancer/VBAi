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
    $project = $book.VBProject
    $inventory = Invoke-Session @{ Command='list_toolbars' }
    $targetState = @($inventory.Toolbars | Where-Object { $_.Properties.BuiltIn -and $_.Properties.Enabled })[0]
    if ($null -eq $targetState) { throw 'No enabled built-in normal toolbar was found.' }
    $toolbarName = $targetState.Properties.Name
    $originalBar = $vbe.CommandBars.Item($toolbarName)
    $originalVisible = [bool]$originalBar.Visible
    $action = if ($originalVisible) { 'hide' } else { 'show' }
    $changed = Invoke-Session @{ Command='set_toolbar_visibility'; ObjectName=$toolbarName; Action=$action; ExpectedWindowVersion=$targetState.WindowVersion }
    if (-not $changed.Verified -or [bool]$originalBar.Visible -eq $originalVisible) { throw 'Native toolbar visibility did not change.' }
    $staleRejected = $false
    try { Invoke-Session @{ Command='set_toolbar_visibility'; ObjectName=$toolbarName; Action=$action; ExpectedWindowVersion=$targetState.WindowVersion } | Out-Null }
    catch { $staleRejected = $_.Exception.Message -match 'state changed' }
    $noop = Invoke-Session @{ Command='set_toolbar_visibility'; ObjectName=$toolbarName; Action=$action; ExpectedWindowVersion=$changed.After.WindowVersion }
    $restoreAction = if ($originalVisible) { 'show' } else { 'hide' }
    $restored = Invoke-Session @{ Command='set_toolbar_visibility'; ObjectName=$toolbarName; Action=$restoreAction; ExpectedWindowVersion=$changed.After.WindowVersion }
    $temporaryBar = $vbe.CommandBars.Add('VBAi isolated toolbar probe',4,$false,$true)
    $button = $temporaryBar.Controls.Add(1); $button.Caption = 'VBAi isolated test'
    $temporaryState = @((Invoke-Session @{ Command='list_toolbars' }).Toolbars | Where-Object { $_.Properties.Name -eq 'VBAi isolated toolbar probe' })[0]
    $temporaryShown = Invoke-Session @{ Command='set_toolbar_visibility'; ObjectName='VBAi isolated toolbar probe'; Action='show'; ExpectedWindowVersion=$temporaryState.WindowVersion }
    $temporaryHidden = Invoke-Session @{ Command='set_toolbar_visibility'; ObjectName='VBAi isolated toolbar probe'; Action='hide'; ExpectedWindowVersion=$temporaryShown.After.WindowVersion }
    [pscustomobject]@{ BuiltIn=$toolbarName; Changed=$changed; Restored=$restored; StaleRejected=$staleRejected; Noop=$noop; TemporaryShown=$temporaryShown; TemporaryHidden=$temporaryHidden; Mvid=$assembly.ManifestModule.ModuleVersionId.ToString() } | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $directory 'native-toolbar-visibility.json') -Encoding UTF8
    if (-not $staleRejected -or $noop.Applied -or -not $noop.Verified -or -not $restored.Verified -or $restored.After.WindowVersion -ne $targetState.WindowVersion -or -not $temporaryShown.Verified -or -not $temporaryHidden.Verified) { throw 'Native toolbar cycle or guards failed.' }
} finally {
    try {
        if ($null -ne $temporaryBar) { $temporaryBar.Delete() }
        if ($null -ne $originalBar) { $originalBar.Visible = $originalVisible }
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

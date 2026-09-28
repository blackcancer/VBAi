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
    $temporaryBar = $vbe.CommandBars.Add('VBAi isolated docking probe',4,$false,$true)
    $button = $temporaryBar.Controls.Add(1); $button.Caption = 'Docking test'
    $temporaryBar.Visible = $true
    $cases = @()
    $originalLayout = @((Invoke-Session @{ Command='list_toolbars' }).Toolbars | Where-Object { $_.Properties.Name -eq 'VBAi isolated docking probe' })[0]
    foreach ($action in @('left','top','right','bottom','float')) {
        $state = @((Invoke-Session @{ Command='list_toolbars' }).Toolbars | Where-Object { $_.Properties.Name -eq 'VBAi isolated docking probe' })[0]
        $result = Invoke-Session @{ Command='set_toolbar_position'; ObjectName='VBAi isolated docking probe'; Action=$action; ExpectedToolbarLayoutVersion=$state.ToolbarLayoutVersion }
        $cases += [pscustomobject]@{ Action=$action; Result=$result }
        if (-not $result.Verified -or -not $result.VisibilityPreserved) { throw "Native toolbar docking failed for $action." }
    }
    $state = @((Invoke-Session @{ Command='list_toolbars' }).Toolbars | Where-Object { $_.Properties.Name -eq 'VBAi isolated docking probe' })[0]
    $placement = Invoke-Session @{ Command='set_toolbar_placement'; ObjectName='VBAi isolated docking probe'; Action='float'; ToolbarLeft=250; ToolbarTop=180; ExpectedToolbarLayoutVersion=$state.ToolbarLayoutVersion }
    if (-not $placement.Verified) { throw 'Floating placement was not verified.' }
    $state = @((Invoke-Session @{ Command='list_toolbars' }).Toolbars | Where-Object { $_.Properties.Name -eq 'VBAi isolated docking probe' })[0]
    $staleRequest = @{ Command='set_toolbar_position'; ObjectName='VBAi isolated docking probe'; Action='left'; ExpectedToolbarLayoutVersion=$state.ToolbarLayoutVersion }
    $temporaryBar.Left = [int]$temporaryBar.Left + 10
    $staleRejected = $false
    try { Invoke-Session $staleRequest | Out-Null }
    catch { $staleRejected = $_.Exception.Message -match 'layout changed' }
    $temporaryBar.Visible = $false
    $state = @((Invoke-Session @{ Command='list_toolbars' }).Toolbars | Where-Object { $_.Properties.Name -eq 'VBAi isolated docking probe' })[0]
    $hiddenDock = Invoke-Session @{ Command='set_toolbar_position'; ObjectName='VBAi isolated docking probe'; Action='top'; ExpectedToolbarLayoutVersion=$state.ToolbarLayoutVersion }
    $state = @((Invoke-Session @{ Command='list_toolbars' }).Toolbars | Where-Object { $_.Properties.Name -eq 'VBAi isolated docking probe' })[0]
    $rowPlacement = Invoke-Session @{ Command='set_toolbar_placement'; ObjectName='VBAi isolated docking probe'; Action='row'; RowIndex=2; ExpectedToolbarLayoutVersion=$state.ToolbarLayoutVersion }
    if (-not $rowPlacement.Verified) { throw 'Docked row placement was not verified.' }
    $temporaryBar.Protection = 16
    $state = @((Invoke-Session @{ Command='list_toolbars' }).Toolbars | Where-Object { $_.Properties.Name -eq 'VBAi isolated docking probe' })[0]
    $protectedRejected = $false
    try { Invoke-Session @{ Command='set_toolbar_position'; ObjectName='VBAi isolated docking probe'; Action='float'; ExpectedToolbarLayoutVersion=$state.ToolbarLayoutVersion } | Out-Null }
    catch { $protectedRejected = $_.Exception.Message -match 'protection' }
    [pscustomobject]@{ Cases=$cases; FloatingPlacement=$placement; RowPlacement=$rowPlacement; StaleGeometryRejected=$staleRejected; HiddenDock=$hiddenDock; ProtectedRejected=$protectedRejected; Mvid=$assembly.ManifestModule.ModuleVersionId.ToString() } | ConvertTo-Json -Depth 9 | Set-Content (Join-Path $directory 'native-toolbar-docking.json') -Encoding UTF8
    if (-not $staleRejected -or -not $protectedRejected -or -not $hiddenDock.Verified -or $temporaryBar.Visible) { throw 'Native toolbar docking guard or hidden-state preservation failed.' }
} finally {
    try {
        if ($null -ne $temporaryBar) { $temporaryBar.Protection = 0; $temporaryBar.Delete() }
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

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
    foreach ($name in @('SecondPane', 'ThirdPane')) {
        $component = $project.VBComponents.Add(1); $component.Name = $name
        $component.CodeModule.AddFromString('Option Explicit'); $component.CodeModule.CodePane.Show()
    }
    $results = @()
    foreach ($id in @(1826, 2561, 2562)) {
        $command = $vbe.CommandBars.FindControl(1, $id)
        if (-not $command.Enabled) { throw "Arrangement command $id is disabled." }
        $action = switch ($id) { 1826 { 'cascade' } 2561 { 'tile_vertical' } 2562 { 'tile_horizontal' } }
        $before = Invoke-Session @{ Command = 'editor_layout' }
        $arranged = Invoke-Session @{ Command = 'arrange_editor_windows'; Action = $action; ExpectedWindowVersion = $before.WindowVersion; ControlCaption = $command.Caption }
        if (-not $arranged.Verified) { throw ('Arrangement was not verified: ' + ($arranged | ConvertTo-Json -Depth 10 -Compress)) }
        $snapshot = Invoke-Session @{ Command = 'vbe_windows' }
        $windows = @($snapshot.Windows | Where-Object { $_.Properties.Type -eq 0 } | ForEach-Object { $_.Properties })
        $results += [pscustomobject]@{ Id = $id; Caption = $command.Caption; Windows = $windows; Verified = $arranged.Verified; WindowVersion = $arranged.WindowVersion }
    }
    $window = $vbe.ActiveCodePane.Window
    $stateResults = @()
    foreach ($action in @('maximize', 'restore', 'minimize', 'restore')) {
        $before = Invoke-Session @{ Command = 'window_layout'; WindowCaption = $window.Caption; WindowType = 0 }
        $result = Invoke-Session @{ Command = 'set_window_state'; WindowCaption = $window.Caption; WindowType = 0; ExpectedWindowVersion = $before.WindowVersion; Action = $action }
        if (-not $result.Verified) { throw ('State not verified: ' + ($result | ConvertTo-Json -Depth 10 -Compress)) }
        $stateResults += [pscustomobject]@{ Action = $action; Verified = $result.Verified; State = $window.WindowState; Caption = $window.Caption }
    }
    $before = Invoke-Session @{ Command = 'window_layout'; WindowCaption = $window.Caption; WindowType = 0 }
    $bounds = Invoke-Session @{ Command = 'set_window_bounds'; WindowCaption = $window.Caption; WindowType = 0; ExpectedWindowVersion = $before.WindowVersion; Left = 40; Top = 50; Width = 600; Height = 400 }
    if (-not $bounds.Verified) { throw 'Bounds were not verified.' }
    [pscustomobject]@{ States = $stateResults; Bounds = $bounds } | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $directory 'window-state-bounds-probe.json') -Encoding UTF8
    $results | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'window-arrangement-probe.json') -Encoding UTF8
    $results | ConvertTo-Json -Depth 10
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

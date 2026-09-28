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
    $module = $project.VBComponents.Add(1); $module.Name = 'EditorCommandProbe'
    $module.CodeModule.AddFromString("Public Sub First()`r`n    Debug.Print 1`r`nEnd Sub`r`nPublic Sub Second()`r`n    Debug.Print 2`r`nEnd Sub")
    $pane = $module.CodeModule.CodePane; $pane.Show(); $pane.SetSelection(2, 1, 2, 1)
    $module.CodeModule.AddFromString(((1..150 | ForEach-Object { "' Scroll probe line $_" }) -join "`r`n"))
    $layout = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
    $split = $vbe.CommandBars.FindControl(1, 302)
    $splitResult = Invoke-Session @{ Command = 'set_code_split'; Project = $project.Name; Module = $module.Name; ExpectedSha256 = $layout.Panes[0].State.Sha256; ExpectedMode = 2; StartLine = 2; Action = 'split'; ControlCaption = $split.Caption }
    if (-not $splitResult.Verified) { throw 'Split was not verified.' }
    $layout = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
    if ($layout.Panes.Count -ne 2) { throw 'Expected two independently targetable panes.' }
    $results = @()
    foreach ($index in @(0, 1)) {
        # Refresh tokens because scrolling one pane can change the other viewport.
        $layout = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
        $target = $layout.Panes[$index]
        $result = Invoke-Session @{ Command = 'scroll_code_pane'; Project = $project.Name; Module = $module.Name; Pane = $target.Pane; ExpectedWindowVersion = $target.WindowVersion; ExpectedSha256 = $target.State.Sha256; ExpectedMode = 2; StartLine = (30 + 40 * $index) }
        if (-not $result.Verified) { throw ('Scroll unverified: ' + ($result | ConvertTo-Json -Depth 8 -Compress)) }
        $results += $result
    }
    function Assert-Refused($Request, [string]$Expected) {
        $message = $null
        try { $response = $script:session.Execute($Request); if (-not $response.Ok) { $message = $response.Error } }
        catch { $message = $_.Exception.ToString() }
        if (-not $message -or -not $message.Contains($Expected)) { throw "Expected refusal '$Expected', got '$message'." }
        return $message
    }
    $guards = @()
    $layout = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
    $target = $layout.Panes[0]
    $request = New-Object CodexVBE.Request
    $request.Command = 'scroll_code_pane'; $request.Project = $project.Name; $request.Module = $module.Name
    $request.Pane = $target.Pane; $request.ExpectedWindowVersion = $target.WindowVersion
    $request.ExpectedSha256 = $target.State.Sha256; $request.ExpectedMode = 2; $request.StartLine = 50
    $null = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
    $message = Assert-Refused $request 'current Pane token'
    $guards += [pscustomobject]@{ Case = 'ExpiredToken'; Error = $message }
    $layout = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
    $target = $layout.Panes[0]; $request.Pane = $target.Pane
    $request.ExpectedWindowVersion = 'stale'
    $message = Assert-Refused $request 'The pane changed'
    $guards += [pscustomobject]@{ Case = 'StaleViewport'; Error = $message }
    $request.ExpectedWindowVersion = $target.WindowVersion; $request.ExpectedSha256 = 'stale'
    $message = Assert-Refused $request 'The module changed'
    $guards += [pscustomobject]@{ Case = 'StaleSource'; Error = $message }
    $afterGuards = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
    if ($afterGuards.Panes[0].WindowVersion -ne $target.WindowVersion) { throw 'Rejected requests changed the viewport.' }
    $guards | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'code-pane-scroll-guards.json') -Encoding UTF8
    $results | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'code-pane-scroll-probe.json') -Encoding UTF8
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

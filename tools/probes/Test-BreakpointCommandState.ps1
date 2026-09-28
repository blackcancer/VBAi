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
    $module = $project.VBComponents.Add(1); $module.Name = 'BreakpointStateProbe'
    $module.CodeModule.AddFromString("Public Sub ProbeState()`r`n    Debug.Print 1`r`n    Debug.Print 2`r`nEnd Sub")
    $pane = $module.CodeModule.CodePane
    $pane.Show(); $pane.SetSelection(2, 1, 2, 1)
    $toggle = $vbe.CommandBars.FindControl(1, 51)
    $clear = $vbe.CommandBars.FindControl(1, 579)
    if ($null -eq $toggle -or $toggle.Caption -notmatch 'point|breakpoint' -or -not $toggle.Enabled) { throw 'Expected native toggle command is unavailable.' }
    $rows = @()
    $rows += [pscustomobject]@{ Stage = 'before'; Line = 2; ToggleState = $toggle.State; ClearEnabled = $clear.Enabled }
    $toggle.Execute()
    $rows += [pscustomobject]@{ Stage = 'added'; Line = 2; ToggleState = $toggle.State; ClearEnabled = $clear.Enabled }
    $pane.SetSelection(3, 1, 3, 1)
    $rows += [pscustomobject]@{ Stage = 'other-line'; Line = 3; ToggleState = $toggle.State; ClearEnabled = $clear.Enabled }
    $pane.SetSelection(2, 1, 2, 1); $toggle.Execute()
    $rows += [pscustomobject]@{ Stage = 'removed'; Line = 2; ToggleState = $toggle.State; ClearEnabled = $clear.Enabled }
    $rows | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'breakpoint-command-state.json') -Encoding UTF8
    $rows | Format-Table
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

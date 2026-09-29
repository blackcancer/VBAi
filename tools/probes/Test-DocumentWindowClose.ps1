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
    $before = [string]$module.CodeModule.Lines(1, $module.CodeModule.CountOfLines)
    $codeWindow = $pane.Window
    $codeResult = Invoke-Session @{ Command = 'close_vbe_window'; WindowCaption = $codeWindow.Caption; WindowType = 0 }
    if ($codeResult.Verification -eq 'Unverified') { throw 'Code window closure is unverified.' }
    if ([string]$module.CodeModule.Lines(1, $module.CodeModule.CountOfLines) -cne $before) { throw 'Closing the code window changed the source.' }
    $pane = $module.CodeModule.CodePane; $pane.Show()
    if (-not $pane.Window.Visible) { throw 'Code window did not reopen.' }
    $form = $project.VBComponents.Add(3); $form.Name = 'WindowCloseProbe'
    $label = $form.Designer.Controls.Add('Forms.Label.1', 'ProbeLabel', $true)
    $label.Caption = 'Preserved after closing the designer'
    $designer = $form.DesignerWindow(); $designer.Visible = $true
    $formResult = Invoke-Session @{ Command = 'close_vbe_window'; WindowCaption = $designer.Caption; WindowType = 1 }
    if ($formResult.Verification -eq 'Unverified') { throw 'Designer window closure is unverified.' }
    $designer = $form.DesignerWindow(); $designer.Visible = $true
    if (-not $designer.Visible -or $form.Designer.Controls.Item('ProbeLabel').Caption -cne 'Preserved after closing the designer') { throw 'Designer did not reopen with its original label.' }
    [pscustomobject]@{ CodeClose = $codeResult; CodeReopened = $pane.Window.Visible; SourcePreserved = $true; DesignerClose = $formResult; DesignerReopened = $designer.Visible; ControlPreserved = $true } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'document-window-close-probe.json') -Encoding UTF8
    Get-Content -LiteralPath (Join-Path $directory 'document-window-close-probe.json')
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

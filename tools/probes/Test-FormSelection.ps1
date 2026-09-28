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
    $form = $project.VBComponents.Add(3); $form.Name = 'SelectionProbe'
    $one = $form.Designer.Controls.Add('Forms.CommandButton.1','FirstButton',$true)
    $two = $form.Designer.Controls.Add('Forms.CommandButton.1','SecondButton',$true)
    $frame = $form.Designer.Controls.Add('Forms.Frame.1','Container',$true)
    $nested = $frame.Controls.Add('Forms.Label.1','NestedLabel',$true)
    $form.DesignerWindow().Visible = $true; $form.DesignerWindow().SetFocus()
    $results = @()
    foreach ($control in @($one,$two,$nested)) {
        try {
            $control.InSelection = $true
            $selectedNames = @($form.Designer.Selected | ForEach-Object { $_.Name })
            $results += [pscustomobject]@{ Name=$control.Name; InSelection=$control.InSelection; Selected=$selectedNames }
        } catch { $results += [pscustomobject]@{ Name=$control.Name; Error=$_.Exception.Message } }
    }
    try { $results += [pscustomobject]@{ FrameSelected=@($frame.Selected | ForEach-Object { $_.Name }); FrameCanPaste=$frame.CanPaste } }
    catch { $results += [pscustomobject]@{ FrameError=$_.Exception.Message } }
    $multi = $form.Designer.Controls.Add('Forms.MultiPage.1','PagesHost',$true)
    $page = $multi.Pages.Item(0)
    $pageButton = $page.Controls.Add('Forms.CommandButton.1','PageButton',$true)
    try { $pageButton.InSelection=$true; $results += [pscustomobject]@{ PageSelected=@($page.Selected | ForEach-Object { $_.Name }); PageCanPaste=$page.CanPaste } }
    catch { $results += [pscustomobject]@{ PageError=$_.Exception.Message } }
    try { $one.InSelection = $false; $two.InSelection = $false; $nested.InSelection = $false; $results += [pscustomobject]@{ Cleared=@($form.Designer.Selected | ForEach-Object { $_.Name }) } }
    catch { $results += [pscustomobject]@{ ClearError=$_.Exception.Message } }
    $results | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $directory 'native-form-selection-discovery.json') -Encoding UTF8
    $results | ConvertTo-Json -Depth 8
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

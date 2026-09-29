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
    $flags = [Reflection.BindingFlags]'Instance,NonPublic'
    $source = $session.GetType().GetMethod('ReferenceEventSource', $flags).Invoke($session, @($project.Name))
    $eventType = $assembly.GetType('VBAi.VbeReferenceEvents', $true)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -TypeDefinition 'public static class ReferenceEventCounter { public static int Count; public static void Changed() { Count++; } }'
    [ReferenceEventCounter]::Count = 0
    function Pump-Events { for ($i=0; $i -lt 10; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 50 } }
    $callback = [Delegate]::CreateDelegate([Action], [ReferenceEventCounter].GetMethod('Changed'))
    $listener = [Activator]::CreateInstance($eventType, $flags, $null, [object[]]@($callback,$null,$null), $null)
    $observe = $eventType.GetMethod('Observe', $flags)
    $observe.Invoke($listener, @($source)) | Out-Null
    $reference = $project.References.AddFromGuid('{420B2830-E718-11CF-893D-00A0C9054228}',1,0)
    Pump-Events
    $afterAdd = [ReferenceEventCounter]::Count
    $project.References.Remove($reference)
    Pump-Events
    $afterRemove = [ReferenceEventCounter]::Count
    $otherBook = $excel.Workbooks.Add()
    $otherBook.VBProject.Name = 'OtherEventScope'
    $otherSource = $session.GetType().GetMethod('ReferenceEventSource', $flags).Invoke($session, @('OtherEventScope'))
    $observe.Invoke($listener, @($otherSource)) | Out-Null
    $reference = $project.References.AddFromGuid('{420B2830-E718-11CF-893D-00A0C9054228}',1,0)
    $project.References.Remove($reference)
    Pump-Events
    $afterOldScope = [ReferenceEventCounter]::Count
    $reference = $otherBook.VBProject.References.AddFromGuid('{420B2830-E718-11CF-893D-00A0C9054228}',1,0)
    $otherBook.VBProject.References.Remove($reference)
    Pump-Events
    $afterNewScope = [ReferenceEventCounter]::Count
    $listener.Dispose()
    $reference = $otherBook.VBProject.References.AddFromGuid('{420B2830-E718-11CF-893D-00A0C9054228}',1,0)
    $otherBook.VBProject.References.Remove($reference)
    Pump-Events
    $afterDispose = [ReferenceEventCounter]::Count
    [pscustomobject]@{ AfterAdd=$afterAdd; AfterRemove=$afterRemove; AfterDispose=$afterDispose; AfterOldScope=$afterOldScope; AfterNewScope=$afterNewScope; Assembly=$assembly.Location; Mvid=$assembly.ManifestModule.ModuleVersionId.ToString() } | ConvertTo-Json | Set-Content (Join-Path $directory 'reference-events-native.json') -Encoding UTF8
    if ($afterAdd -ne 1 -or $afterRemove -ne 2 -or $afterDispose -ne 4 -or $afterOldScope -ne 2 -or $afterNewScope -ne 4) { throw 'Native reference notifications or unsubscribe failed.' }
} finally {
    if ($null -ne $listener) { $listener.Dispose() }
    if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
    else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
    try { if ($null -ne $otherBook) { $otherBook.Close($false) }; if ($null -ne $book) { $book.Close($false) } }
    finally {
        try { if ($null -ne $excel) { $excel.Quit() } }
        finally { if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) { Stop-Process -Id $probeProcess.Id -Force } }
    }
}

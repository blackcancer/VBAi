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
    $eventType = $assembly.GetType('VBAi.VbeCollectionEvents', $true)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -TypeDefinition 'public static class CollectionEventCounter { public static int Components; public static int Projects; public static void ComponentChanged() { Components++; } public static void ProjectChanged() { Projects++; } public static void Observe(object session, object listener, string method, string project) { var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic; var source = session.GetType().GetMethod(method, flags).Invoke(session, method == "ProjectsEventSource" ? new object[0] : new object[] { project }); listener.GetType().GetMethod("Observe", flags).Invoke(listener, new object[] { source }); } }'
    function Pump-Events { for ($i=0; $i -lt 10; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 50 } }
    $componentsCallback = [Delegate]::CreateDelegate([Action], [CollectionEventCounter].GetMethod('ComponentChanged'))
    $projectsCallback = [Delegate]::CreateDelegate([Action], [CollectionEventCounter].GetMethod('ProjectChanged'))
    $componentListener = [Activator]::CreateInstance($eventType, $flags, $null, [object[]]@($componentsCallback,$true,$null,$null), $null)
    $projectListener = [Activator]::CreateInstance($eventType, $flags, $null, [object[]]@($projectsCallback,$false,$null,$null), $null)
    $observe = $eventType.GetMethod('Observe', $flags)
    [CollectionEventCounter]::Observe($session, $projectListener, 'ProjectsEventSource', $null)
    [CollectionEventCounter]::Observe($session, $componentListener, 'ComponentsEventSource', $project.Name)
    $component = $project.VBComponents.Add(1)
    Pump-Events
    $componentAdded = [CollectionEventCounter]::Components
    $component.Name = 'RenamedEventProbe'
    Pump-Events
    $componentRenamed = [CollectionEventCounter]::Components
    $project.VBComponents.Remove($component)
    Pump-Events
    $componentRemoved = [CollectionEventCounter]::Components
    $projectBefore = [CollectionEventCounter]::Projects
    $otherBook = $excel.Workbooks.Add()
    Pump-Events
    $projectAdded = [CollectionEventCounter]::Projects
    $otherBook.VBProject.Name = 'OtherCollectionScope'
    Pump-Events
    $projectRenamed = [CollectionEventCounter]::Projects
    [CollectionEventCounter]::Observe($session, $componentListener, 'ComponentsEventSource', 'OtherCollectionScope')
    $component = $project.VBComponents.Add(1); $project.VBComponents.Remove($component)
    Pump-Events
    $afterOldScope = [CollectionEventCounter]::Components
    $component = $otherBook.VBProject.VBComponents.Add(1)
    Pump-Events
    $afterNewScope = [CollectionEventCounter]::Components
    $otherBook.Close($false); $otherBook = $null
    Pump-Events
    $projectRemoved = [CollectionEventCounter]::Projects
    $componentListener.Dispose(); $projectListener.Dispose()
    $componentBeforeDispose = [CollectionEventCounter]::Components
    $component = $project.VBComponents.Add(1); $project.VBComponents.Remove($component)
    $otherBook = $excel.Workbooks.Add()
    Pump-Events
    $afterDisposeComponents = [CollectionEventCounter]::Components
    $afterDisposeProjects = [CollectionEventCounter]::Projects
    $result = [pscustomobject]@{ ComponentAdded=$componentAdded; ComponentRenamed=$componentRenamed; ComponentRemoved=$componentRemoved; AfterOldScope=$afterOldScope; AfterNewScope=$afterNewScope; ProjectBefore=$projectBefore; ProjectAdded=$projectAdded; ProjectRenamed=$projectRenamed; ProjectRemoved=$projectRemoved; ComponentBeforeDispose=$componentBeforeDispose; AfterDisposeComponents=$afterDisposeComponents; AfterDisposeProjects=$afterDisposeProjects; Mvid=$assembly.ManifestModule.ModuleVersionId.ToString() }
    $result | ConvertTo-Json | Set-Content (Join-Path $directory 'collection-events-native.json') -Encoding UTF8
    $result | ConvertTo-Json
    if ($componentAdded -le 0 -or $componentRenamed -le $componentAdded -or $componentRemoved -le $componentRenamed -or $afterOldScope -ne $componentRemoved -or $afterNewScope -le $afterOldScope -or $projectAdded -le $projectBefore -or $projectRenamed -le $projectAdded -or $projectRemoved -le $projectRenamed -or $afterDisposeComponents -ne $componentBeforeDispose -or $afterDisposeProjects -ne $projectRemoved) { throw 'Native collection notifications or detach failed.' }
} finally {
    if ($null -ne $componentListener) { $componentListener.Dispose() }
    if ($null -ne $projectListener) { $projectListener.Dispose() }
    if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
    else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
    try { if ($null -ne $otherBook) { $otherBook.Close($false) }; if ($null -ne $book) { $book.Close($false) } }
    finally {
        try { if ($null -ne $excel) { $excel.Quit() } }
        finally { if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) { Stop-Process -Id $probeProcess.Id -Force } }
    }
}

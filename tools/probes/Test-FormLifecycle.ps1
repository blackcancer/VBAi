param(
    [Parameter(Mandatory = $true)][string]$AssemblyPath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [switch]$UseBridge,
    [switch]$AllowTemporaryVbaAccess
)
$ErrorActionPreference = 'Stop'
if (-not $UseBridge) { throw 'This lifecycle test requires the registered add-in bridge.' }
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Threading;
public static class NavigationProbeDialog {
    private delegate bool EnumCallback(IntPtr handle, IntPtr param);
    [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool EnumWindows(EnumCallback callback, IntPtr param);
    [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint pid);
    [System.Runtime.InteropServices.DllImport("user32.dll", CharSet=System.Runtime.InteropServices.CharSet.Unicode)] private static extern int GetWindowText(IntPtr handle, System.Text.StringBuilder text, int count);
    [System.Runtime.InteropServices.DllImport("user32.dll", CharSet=System.Runtime.InteropServices.CharSet.Unicode)] private static extern int GetClassName(IntPtr handle, System.Text.StringBuilder text, int count);
    [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool PostMessage(IntPtr handle, uint message, IntPtr wparam, IntPtr lparam);
    [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr handle);
    [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr handle);
    public static string[] WindowStates(int owner) {
        var result = new List<string>();
        EnumWindows((handle, ignored) => {
            uint pid; GetWindowThreadProcessId(handle, out pid);
            if (pid != owner || !IsWindowVisible(handle)) return true;
            var title = new System.Text.StringBuilder(256); GetWindowText(handle,title,256);
            var kind = new System.Text.StringBuilder(64); GetClassName(handle,kind,64);
            result.Add(kind + " | " + title + " | Enabled=" + IsWindowEnabled(handle));
            return true;
        }, IntPtr.Zero);
        return result.ToArray();
    }
    public static IntPtr FindOwnedForm(int owner, string caption) {
        var matches = new List<IntPtr>();
        EnumWindows((handle, ignored) => {
            uint pid; GetWindowThreadProcessId(handle, out pid);
            if (pid != owner || !IsWindowVisible(handle)) return true;
            var title = new System.Text.StringBuilder(256); GetWindowText(handle,title,256);
            var kind = new System.Text.StringBuilder(64); GetClassName(handle,kind,64);
            if (title.ToString() == caption && kind.ToString().StartsWith("Thunder", StringComparison.Ordinal)) matches.Add(handle);
            return true;
        }, IntPtr.Zero);
        if (matches.Count > 1) throw new InvalidOperationException("Ambiguous owned test form.");
        return matches.Count == 1 ? matches[0] : IntPtr.Zero;
    }
    public static bool OwnedWindowEnabled(long window, int owner) {
        uint pid; GetWindowThreadProcessId(new IntPtr(window), out pid);
        if (pid != owner) throw new InvalidOperationException("Wrong window owner.");
        return IsWindowEnabled(new IntPtr(window));
    }
    public static bool CloseOwnedForm(int owner, string caption) {
        IntPtr handle = FindOwnedForm(owner, caption);
        return handle != IntPtr.Zero && PostMessage(handle, 0x10, IntPtr.Zero, IntPtr.Zero);
    }
    public static void CloseOwnedDiagnostic(int owner) {
        EnumWindows((handle, ignored) => {
            uint pid; GetWindowThreadProcessId(handle, out pid);
            if (pid != owner) return true;
            var title = new System.Text.StringBuilder(256); GetWindowText(handle,title,256);
            var kind = new System.Text.StringBuilder(64); GetClassName(handle,kind,64);
            if (kind.ToString() == "#32770" && (title.ToString() == "Microsoft Visual Basic pour Applications" || title.ToString() == "Microsoft Visual Basic for Applications" || title.ToString() == "Microsoft Visual Basic"))
                PostMessage(handle, 0x10, IntPtr.Zero, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);
    }
}
public sealed class NavigationProbeContext : SynchronizationContext {
    private readonly Queue<Action> queue = new Queue<Action>();
    public override void Post(SendOrPostCallback callback, object state) { queue.Enqueue(() => callback(state)); }
    public void RunNext() { int count = 0; while (queue.Count > 0 && count++ < 20) queue.Dequeue()(); }
}
"@
$priorContext = [Threading.SynchronizationContext]::Current
$probeContext = New-Object NavigationProbeContext
[Threading.SynchronizationContext]::SetSynchronizationContext($probeContext)
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
    if ($UseBridge) {
        Write-Host ('Bridge request: ' + $Fields.Command)
        $requestTimeout = if ($Fields.Command -in @('read_object_browser', 'select_object_browser', 'list_object_browser')) { 30 } else { 5 }
        $response = & (Join-Path $PSScriptRoot '../Invoke-CodexVBE.ps1') -HostProcessId $probeProcess.Id -RequestJson ($Fields | ConvertTo-Json -Compress) -ResponseTimeoutSeconds $requestTimeout | ConvertFrom-Json
        if (-not $response.Ok) { throw $response.Error }
        return $response.Data
    }
    $request = New-Object CodexVBE.Request
    foreach ($key in $Fields.Keys) { $request.$key = $Fields[$key] }
    $response = $script:session.Execute($request)
    if (-not $response.Ok) { throw $response.Error }
    return $response.Data
}
try {
    if ($priorAccess -ne 1) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null }
    if ($UseBridge) {
        $scratch = Join-Path $PSScriptRoot 'CodexVBE-scratch.xlsx'
        if (-not (Test-Path -LiteralPath $scratch)) { throw 'Native Excel scratch workbook is missing.' }
        $probeProcess = Start-Process -FilePath 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE' -ArgumentList @('/x', ('"' + $scratch + '"')) -WindowStyle Hidden -PassThru
        for ($attempt = 0; $attempt -lt 30 -and $null -eq $excel; $attempt++) {
            Start-Sleep -Milliseconds 200
            try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { }
        }
        if ($null -eq $excel) { throw 'The native Excel instance did not register in the ROT.' }
    } else { $excel = New-Object -ComObject Excel.Application }
    $probeProcess = Get-Process EXCEL -ErrorAction Stop
    if (@($probeProcess).Count -ne 1) { throw "Excel isolation was lost." }
    $excel.Visible = $true
    if ($UseBridge) {
        if ($excel.Workbooks.Count -ne 1 -or [IO.Path]::GetFullPath($excel.Workbooks.Item(1).FullName) -ne [IO.Path]::GetFullPath($scratch)) { throw 'Unexpected workbook in native Excel test instance.' }
        $excel.Workbooks.Item(1).Close($false)
    }
    $book = $excel.Workbooks.Add()
    $vbe = $excel.GetType().InvokeMember('VBE', [Reflection.BindingFlags]::GetProperty, $null, $excel, $null)
    $vbe.MainWindow.Visible = $true
    $session = [Activator]::CreateInstance($assembly.GetType('CodexVBE.VbeSession', $true), [object[]]@($vbe))
    if ($UseBridge) {
        Get-Process -Id $probeProcess.Id -Module | Select-Object ModuleName,FileName | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'loaded-addin-modules.json') -Encoding UTF8
        $processInfo = Get-CimInstance Win32_Process -Filter "ProcessId=$($probeProcess.Id)"
        $owner = Invoke-CimMethod -InputObject $processInfo -MethodName GetOwner
        [pscustomobject]@{ Caller = [Security.Principal.WindowsIdentity]::GetCurrent().Name; CallerSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value; ExcelOwner = ($owner.Domain + '\' + $owner.User); ExcelPath = $processInfo.ExecutablePath; ExcelId = $probeProcess.Id; CodeBaseSeen = (Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\CLSID\{8E854243-087F-4D6C-9E0E-8622B0E50883}\InprocServer32').CodeBase } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'bridge-process-identity.json') -Encoding UTF8
        $status = Invoke-Session @{ Command = 'status' }
        $status | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'bridge-build-identity.json') -Encoding UTF8
        if ($status.AssemblyModuleVersionId -ne $assembly.ManifestModule.ModuleVersionId.ToString('D') -or $status.HostProcessId -ne $probeProcess.Id) { $status | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'bridge-identity-mismatch.json') -Encoding UTF8; throw ('The bridge loaded a different assembly or process: ' + ($status | ConvertTo-Json -Compress)) }
    }
    $project = $book.VBProject
    $results = @()
    foreach ($modal in @($false, $true)) {
        $kind = if ($modal) { 'Modal' } else { 'Modeless' }
        $caption = 'VBAi isolated ' + $kind + ' lifecycle'
        $form = $project.VBComponents.Add(3); $form.Name = $kind + 'LifecycleProbe'
        $form.Properties.Item('Caption').Value = $caption
        $modalProperty = $form.Properties.Item('ShowModal').Value
        [pscustomobject]@{ Value=$modalProperty; Type=$modalProperty.GetType().FullName } | ConvertTo-Json | Set-Content (Join-Path $directory 'form-showmodal-property.json') -Encoding UTF8
        $initialTree = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
        Invoke-Session @{ Command='set_form_property'; Project=$project.Name; Form=$form.Name; Property='ShowModal'; Value=$modal; ExpectedFormVersion=$initialTree.TreeVersion } | Out-Null
        $code = @"
Private closeAttempts As Long
Private Sub UserForm_Initialize()
    closeAttempts = 0
End Sub
Private Sub UserForm_Activate()
    ThisWorkbook.Worksheets(1).Range("A6").Value2 = "$kind active"
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    closeAttempts = closeAttempts + 1
    ThisWorkbook.Worksheets(1).Range("A7").Value2 = closeAttempts
    ThisWorkbook.Worksheets(1).Range("A8").Value2 = CloseMode
    If closeAttempts = 1 Then Cancel = True
End Sub
"@
        $form.CodeModule.AddFromString($code)
        $book.Worksheets.Item(1).Range('A6:A8').ClearContents()
        $form.DesignerWindow().Visible = $true; $form.DesignerWindow().SetFocus()
        $beforeCode = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$form.Name }
        $beforeTree = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
        $observedShowModal = ($beforeTree.Properties | Where-Object Name -eq 'ShowModal').Value
        if ([bool]$observedShowModal -ne $modal) { throw 'ShowModal readback did not match the requested mode.' }
        $runCommand = $vbe.CommandBars.FindControl(1,186)
        $beforeRuntimeWindows = Invoke-Session @{ Command='read_runtime_forms' }
        if ($beforeRuntimeWindows.DetectedCount -ne 0 -or -not $beforeRuntimeWindows.EnumerationSucceeded) { throw 'A Designer was misclassified as a runtime form.' }
        $started = Invoke-Session @{ Command='run_form'; Project=$project.Name; Form=$form.Name; ExpectedMode=2; ExpectedSha256=$beforeCode.Sha256; ExpectedTreeVersion=$beforeTree.TreeVersion; ControlCaption=$runCommand.Caption }
        for ($attempt=0; $attempt -lt 30 -and [NavigationProbeDialog]::FindOwnedForm($probeProcess.Id,$caption) -eq [IntPtr]::Zero; $attempt++) { Start-Sleep -Milliseconds 100 }
        if ([NavigationProbeDialog]::FindOwnedForm($probeProcess.Id,$caption) -eq [IntPtr]::Zero) { throw "$kind runtime form was not visible." }
        $runningStatus = Invoke-Session @{ Command='form_run_status'; Project=$project.Name; Query=$started.OperationId }
        $runningState = Invoke-Session @{ Command='debug_state'; Project=$project.Name }
        $runtimeWindows = Invoke-Session @{ Command='read_runtime_forms' }
        $matchingRuntime = @($runtimeWindows.Windows | Where-Object Caption -eq $caption)
        if ($matchingRuntime.Count -ne 1 -or $runtimeWindows.HostProcessId -ne $probeProcess.Id -or $matchingRuntime[0].ProjectIdentityVerified -or -not $runtimeWindows.EnumerationSucceeded) { throw 'Runtime window inventory did not identify exactly one observed form.' }
        $hostEnabled = [NavigationProbeDialog]::OwnedWindowEnabled([long]$excel.Hwnd,$probeProcess.Id)
        if ($book.Worksheets.Item(1).Range('A6').Value2 -ne "$kind active" -or $runningStatus.State -eq 'Failed' -or $runningStatus.RuntimeVerified) { throw 'Form activation or reported status was incorrect.' }
        if ($hostEnabled -eq $modal) { throw 'The Excel window enabled state did not match ShowModal.' }
        if (-not [NavigationProbeDialog]::CloseOwnedForm($probeProcess.Id,$caption)) { throw 'First close was not sent.' }
        Start-Sleep -Milliseconds 500
        $firstAttempts = $book.Worksheets.Item(1).Range('A7').Value2
        $firstCloseMode = $book.Worksheets.Item(1).Range('A8').Value2
        $stillVisible = [NavigationProbeDialog]::FindOwnedForm($probeProcess.Id,$caption) -ne [IntPtr]::Zero
        $cancelledStatus = Invoke-Session @{ Command='form_run_status'; Project=$project.Name; Query=$started.OperationId }
        $afterCancelledWindows = Invoke-Session @{ Command='read_runtime_forms' }
        if (@($afterCancelledWindows.Windows | Where-Object Handle -eq $matchingRuntime[0].Handle).Count -ne 1) { throw 'Cancelled form disappeared from runtime inventory.' }
        if (-not $stillVisible -or $firstAttempts -ne 1 -or $firstCloseMode -ne 0 -or $cancelledStatus.RuntimeVerified) { throw 'QueryClose cancellation was not respected.' }
        if (-not [NavigationProbeDialog]::CloseOwnedForm($probeProcess.Id,$caption)) { throw 'Second close was not sent.' }
        for ($attempt=0; $attempt -lt 30 -and [NavigationProbeDialog]::FindOwnedForm($probeProcess.Id,$caption) -ne [IntPtr]::Zero; $attempt++) { Start-Sleep -Milliseconds 100 }
        $finalState = Invoke-Session @{ Command='debug_state'; Project=$project.Name }
        $finalCode = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$form.Name }
        $secondAttempts = $book.Worksheets.Item(1).Range('A7').Value2
        if ([NavigationProbeDialog]::FindOwnedForm($probeProcess.Id,$caption) -ne [IntPtr]::Zero -or $secondAttempts -ne 2 -or $finalState.Mode -ne 2 -or $finalCode.Sha256 -ne $beforeCode.Sha256) { throw 'Second close or final source/state verification failed.' }
        $afterRuntimeWindows = Invoke-Session @{ Command='read_runtime_forms' }
        if ($afterRuntimeWindows.DetectedCount -ne 0 -or -not $afterRuntimeWindows.EnumerationSucceeded) { throw 'Closed runtime form remained in inventory.' }
        $results += [pscustomobject]@{ RuntimeWindows=$runtimeWindows; FinalRuntimeWindows=$afterRuntimeWindows; Kind=$kind; ShowModal=$observedShowModal; ExcelWindowEnabled=$hostEnabled; RunningStatus=$runningStatus; RunningState=$runningState; FirstCloseCancelled=$stillVisible; FirstCloseMode=$firstCloseMode; FirstCloseAttempts=$firstAttempts; SecondCloseAttempts=$secondAttempts; FinalState=$finalState; CodeUnchanged=$true }
        [pscustomobject]@{ Cases=$results; BridgeIdentity=$status } | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $directory 'installed-form-lifecycle.json') -Encoding UTF8
    }
    $results | Select-Object Kind,ShowModal,ExcelWindowEnabled,FirstCloseCancelled,SecondCloseAttempts,CodeUnchanged | ConvertTo-Json
} catch {
    $_ | Format-List * -Force | Out-String | Set-Content (Join-Path $directory "form-lifecycle-error.txt") -Encoding UTF8
    throw
} finally {
    [Threading.SynchronizationContext]::SetSynchronizationContext($priorContext)
    if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
    else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
    if ($null -ne $probeProcess) {
        [NavigationProbeDialog]::CloseOwnedDiagnostic($probeProcess.Id)
        foreach ($caption in @('VBAi isolated Modal lifecycle','VBAi isolated Modeless lifecycle')) {
            [NavigationProbeDialog]::CloseOwnedForm($probeProcess.Id,$caption) | Out-Null
            Start-Sleep -Milliseconds 100
            [NavigationProbeDialog]::CloseOwnedForm($probeProcess.Id,$caption) | Out-Null
        }
    }
    try { if ($null -ne $book) { $book.Close($false) } }
    finally {
        try { if ($null -ne $excel) { $excel.Quit() } }
        finally { if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) { Stop-Process -Id $probeProcess.Id -Force } }
    }
}

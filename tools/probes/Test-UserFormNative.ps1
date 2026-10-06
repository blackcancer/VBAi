param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Selection','Lifecycle','Clipboard','ClipboardRecoveryDiscovery','History')]
    [string] $Scenario,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $AssemblyPath,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $OutputDirectory,
    [switch] $UseBridge,
    [switch] $AllowTemporaryVbaAccess
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'VbeProbe.Common.ps1')
$scenarioParameters = @{
    'Clipboard' = @{ Allowed = @('AssemblyPath','OutputDirectory','AllowTemporaryVbaAccess'); Required = @('AssemblyPath','OutputDirectory') }
    'Selection' = @{ Allowed = @('AssemblyPath','OutputDirectory','AllowTemporaryVbaAccess'); Required = @('AssemblyPath','OutputDirectory') }
    'ClipboardRecoveryDiscovery' = @{ Allowed = @('AssemblyPath','OutputDirectory','AllowTemporaryVbaAccess'); Required = @('AssemblyPath','OutputDirectory') }
    'Lifecycle' = @{ Allowed = @('AssemblyPath','OutputDirectory','UseBridge','AllowTemporaryVbaAccess'); Required = @('AssemblyPath','OutputDirectory') }
    'History' = @{ Allowed = @('AssemblyPath','OutputDirectory','AllowTemporaryVbaAccess'); Required = @('AssemblyPath','OutputDirectory') }
}
$contract = $scenarioParameters[$Scenario]
Assert-VbeProbeParameters -Bound $PSBoundParameters -Required $contract.Required -Allowed $contract.Allowed
# Every scenario retains its original isolated-host, trust, clipboard, and cleanup guards.
# There is deliberately no implicit multi-scenario native execution.
function Initialize-NativeUserFormProbe {
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
}

switch ($Scenario) {
    'Selection' {
        $ErrorActionPreference = 'Stop'
        . Initialize-NativeUserFormProbe
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
    }
    'Lifecycle' {
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
        . Initialize-NativeUserFormProbe
        $probeProcess = $null
        $otherBook = $null
        $excel = $null; $book = $null; $form = $null; $session = $null
        function Invoke-Session([hashtable]$Fields) {
            if ($UseBridge) {
                Write-Host ('Bridge request: ' + $Fields.Command)
                $requestTimeout = if ($Fields.Command -in @('read_object_browser', 'select_object_browser', 'list_object_browser')) { 30 } else { 5 }
                $response = & (Join-Path $PSScriptRoot '../Invoke-VBAi.ps1') -HostProcessId $probeProcess.Id -RequestJson ($Fields | ConvertTo-Json -Compress) -ResponseTimeoutSeconds $requestTimeout | ConvertFrom-Json
                if (-not $response.Ok) { throw $response.Error }
                return $response.Data
            }
            $request = New-Object VBAi.Request
            foreach ($key in $Fields.Keys) { $request.$key = $Fields[$key] }
            $response = $script:session.Execute($request)
            if (-not $response.Ok) { throw $response.Error }
            return $response.Data
        }
        try {
            if ($priorAccess -ne 1) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null }
            if ($UseBridge) {
                $scratch = Join-Path $PSScriptRoot 'VBAi-scratch.xlsx'
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
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
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
    }
    'Clipboard' {
        $ErrorActionPreference = 'Stop'
        . Initialize-NativeUserFormProbe
        $probeProcess = $null
        $otherBook = $null
        $excel = $null; $book = $null; $form = $null; $session = $null
        function Invoke-Session([hashtable]$Fields) {
            $request = New-Object VBAi.Request
            foreach ($key in $Fields.Keys) { $request.$key = $Fields[$key] }
            Write-Output ("Request: " + $request.Command + " " + $request.Action) | Out-Host
            $response = $script:session.Execute($request)
            if (-not $response.Ok) { throw $response.Error }
            return $response.Data
        }
        Add-Type -AssemblyName System.Windows.Forms
        $backup = New-Object System.Windows.Forms.DataObject
        $backupValues = @{}
        $currentClipboard = [System.Windows.Forms.Clipboard]::GetDataObject()
        if ($null -ne $currentClipboard) {
            foreach ($format in $currentClipboard.GetFormats($false)) {
                $value = $currentClipboard.GetData($format, $false)
                if ($value -is [IO.MemoryStream]) { $value = New-Object IO.MemoryStream -ArgumentList (,($value.ToArray())) }
                elseif ($value -isnot [string]) { throw "Clipboard test refused: unsupported backup format $format." }
                $backup.SetData($format, $false, $value); $backupValues[$format] = $value
            }
        }
        $clipboardTouched = $false
        $proof = $null
        try {
            if ($priorAccess -ne 1) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null }
            $excel = New-Object -ComObject Excel.Application
            $probeProcess = Get-Process EXCEL -ErrorAction Stop
            if (@($probeProcess).Count -ne 1) { throw "Excel isolation was lost." }
            $excel.Visible = $true
            $book = $excel.Workbooks.Add()
            $vbe = $excel.GetType().InvokeMember('VBE', [Reflection.BindingFlags]::GetProperty, $null, $excel, $null)
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
            $settings = [Activator]::CreateInstance($assembly.GetType('VBAi.LlmSettings'),$true); $settings.VbeEditApproval='Automatic'
            $llm = [Activator]::CreateInstance($assembly.GetType('VBAi.LlmVbeTools'),[object[]]@($session,$null,$settings))
            Add-Type -TypeDefinition 'public static class FormCutCardCapture { public static object Last; public static void Capture(object value) { Last=value; } }'
            $cutEvent = $llm.GetType().GetEvent('FormCut')
            $cutDelegate = [Delegate]::CreateDelegate($cutEvent.EventHandlerType,[FormCutCardCapture].GetMethod('Capture'))
            $cutEvent.AddEventHandler($llm,$cutDelegate)
            $project = $book.VBProject
            $llm.BoundProject=$project.Name
            $vbe.MainWindow.Visible = $true
            $form = $project.VBComponents.Add(3); $form.Name = 'ClipboardDesignerSource'
            $button = $form.Designer.Controls.Add('Forms.CommandButton.1','SourceButton',$true)
            $form.CodeModule.AddFromString("Private Sub SourceButton_Click()`r`n    Debug.Print ""handler retained""`r`nEnd Sub")
            $sourceCode = $form.CodeModule.Lines(1,$form.CodeModule.CountOfLines)
            $button.Caption = 'Clipboard native test'; $button.Left = 30; $button.Top = 30
            $form.DesignerWindow().Visible = $true; $form.DesignerWindow().SetFocus()
            function Pump-Designer { for ($i=0; $i -lt 5; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 50 } }
            Pump-Designer
            function Read-DesignerState($name,$parent) { Invoke-Session @{ Command='form_clipboard_state'; Project=$project.Name; Form=$name; ParentPath=$parent } }
            function Invoke-DesignerClipboard($name,$action,$state) {
                if ($action -eq 'cut') {
                    $args = @{ Project=$project.Name; Form=$name; Action=$action; ExpectedDesignerSelectionVersion=$state.SelectionVersion; ExpectedClipboardVersion=$state.ClipboardVersion; ParentPath=$state.ParentPath }
                    $response = $llm.Invoke('native_form_clipboard',($args | ConvertTo-Json -Compress)) | ConvertFrom-Json
                    if (-not $response.Ok) { throw $response.Error }; return $response.Data
                }
                Invoke-Session @{ Command='native_form_clipboard'; Project=$project.Name; Form=$name; Action=$action; ExpectedDesignerSelectionVersion=$state.SelectionVersion; ExpectedClipboardVersion=$state.ClipboardVersion; ParentPath=$state.ParentPath }
            }
            function Select-Designer($name,$parent,[string[]]$names,$state) {
                Invoke-Session @{ Command='select_form_controls'; Project=$project.Name; Form=$name; ParentPath=$parent; Items=$names; ExpectedDesignerSelectionVersion=$state.SelectionVersion }
            }
            $initialSelection = Select-Designer $form.Name '' @('SourceButton') (Read-DesignerState $form.Name)
            if (-not $initialSelection.Verified) { throw 'Initial programmatic selection failed.' }
            $source = Read-DesignerState $form.Name
            if (@($source.Selected).Count -ne 1 -or $source.Selected[0] -ne 'SourceButton') { throw 'Designer selection not read correctly.' }
            $button.Caption = 'Changed caption'
            $staleTree = $false
            try { Invoke-DesignerClipboard $form.Name 'copy' $source | Out-Null }
            catch { $staleTree = $_.Exception.Message -match 'tree or selection changed' }
            $source = Read-DesignerState $form.Name
            $clipboardTouched = $true
            [System.Windows.Forms.Clipboard]::SetText('Temporary clipboard guard test')
            $staleClipboard = $false
            try { Invoke-DesignerClipboard $form.Name 'copy' $source | Out-Null }
            catch { $staleClipboard = $_.Exception.Message -match 'Clipboard changed' -or $_.Exception.Message -match 'tree or selection changed' }
            $copy = Invoke-DesignerClipboard $form.Name 'copy' (Read-DesignerState $form.Name)
            $target = $project.VBComponents.Add(3); $target.Name = 'ClipboardDesignerTarget'
            $target.DesignerWindow().Visible = $true
            $form.DesignerWindow().SetFocus(); Pump-Designer
            $paste = Invoke-DesignerClipboard $target.Name 'paste' (Read-DesignerState $target.Name)
            Pump-Designer
            $countAfterPaste = $target.Designer.Controls.Count
            $pastedCaption = $target.Designer.Controls.Item(0).Caption
            $target.DesignerWindow().SetFocus(); Pump-Designer
            $cut = Invoke-DesignerClipboard $form.Name 'cut' (Read-DesignerState $form.Name)
            Pump-Designer
            $countAfterCut = $form.Designer.Controls.Count
            $tree = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
            $undoAvailable = [bool](($tree.Properties | Where-Object Name -eq CanUndo).Value)
            $undo = $null; $undoRefused = $false; $recovery = $null
            if ($undoAvailable) {
                $undo = Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='undo'; ExpectedTreeVersion=$tree.TreeVersion }
            } else {
                try { Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='undo'; ExpectedTreeVersion=$tree.TreeVersion } | Out-Null }
                catch { $undoRefused = $_.Exception.Message -match 'unavailable' }
                if (-not $undoRefused) { throw 'Unavailable undo was not refused.' }
                [System.Windows.Forms.Clipboard]::SetText('Clipboard replaced after cut')
                $current = Read-DesignerState $form.Name
                $card = [FormCutCardCapture]::Last
                if ($null -eq $card -or $card.Form -ne $form.Name) { throw 'LLM did not emit the cut recovery card.' }
                $uiFlags=[Reflection.BindingFlags]'Instance,Public,NonPublic'
                $chat=[Activator]::CreateInstance($assembly.GetType('VBAi.ChatWindow'),$true)
                try {
                    $chat.GetType().GetMethod('InitializeShell',$uiFlags).Invoke($chat,@()) | Out-Null
                    $chat.GetType().GetMethod('InitializeTranscript',$uiFlags).Invoke($chat,@()) | Out-Null
                    $chat.GetType().GetField('tools',$uiFlags).SetValue($chat,$llm)
                    $view=$chat.GetType().GetMethod('RenderFormCut',$uiFlags).Invoke($chat,@($card))
                    $restoreButton=$view.Child.Children[$view.Child.Children.Count-1]
                    if (-not $restoreButton.IsEnabled) { throw 'Live recovery card button is disabled.' }
                    $restoreButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
                    if (-not $card.Restored -or $restoreButton.IsEnabled) { throw 'UI recovery click did not restore and disable the card.' }
                    $recovery=[pscustomobject]@{ RestoredNamesGeometryAndTabOrder=$card.Restored; NativeError=$null; UiButtonVerified=$true }
                } finally { $chat.Dispose() }
                if (-not $recovery.RestoredNamesGeometryAndTabOrder) { throw ('Root recovery failed: ' + $recovery.NativeError) }
                if ($form.Designer.Controls.Item('SourceButton').Left -ne 30 -or $form.Designer.Controls.Item('SourceButton').Top -ne 30) { throw 'Root recovery position differs.' }
                $current = Read-DesignerState $form.Name
                $replayRejected = $false
                try { Invoke-Session @{ Command='recover_form_cut'; Project=$project.Name; Form=$form.Name; DesignerClipboardRecoveryId=$cut.DesignerClipboardRecoveryId; ExpectedDesignerSelectionVersion=$current.SelectionVersion; ExpectedClipboardVersion=$current.ClipboardVersion } | Out-Null }
                catch { $replayRejected = $_.Exception.Message -match 'already attempted' }
                if (-not $replayRejected) { throw 'Recovery replay was not refused.' }

            }
            $restoredCount = $form.Designer.Controls.Count
            $frame = $form.Designer.Controls.Add('Forms.Frame.1','SourceFrame',$true)
            $first = $frame.Controls.Add('Forms.Label.1','FirstNested',$true); $first.Caption='First nested caption'
            $second = $frame.Controls.Add('Forms.CommandButton.1','SecondNested',$true); $second.Caption='Second nested caption'
            $frame.Width = 210; $frame.Height = 150
            $first.Left=18; $first.Top=24; $first.Width=51; $first.Height=18
            $second.Left=91; $second.Top=60; $second.Width=71; $second.Height=25
            $survivor = $frame.Controls.Add('Forms.Label.1','Survivor',$true); $survivor.Top=110; $survivor.TabIndex=1
            function Read-Placement($container) { @($container.Controls | ForEach-Object { [pscustomobject]@{ Name=$_.Name; Left=$_.Left; Top=$_.Top; Width=$_.Width; Height=$_.Height; TabIndex=$_.TabIndex } } | Sort-Object Name) | ConvertTo-Json -Compress }
            $originalPlacement = Read-Placement $frame
            $multi = $target.Designer.Controls.Add('Forms.MultiPage.1','TargetPages',$true)
            $page = $multi.Pages.Item(0)
            $sourcePath = 'Controls/SourceFrame'
            $targetPath = 'Controls/TargetPages/Pages/' + $page.Name
            $oneSelection = Select-Designer $form.Name $sourcePath @('FirstNested') (Read-DesignerState $form.Name $sourcePath)
            $previousSelection = Read-DesignerState $form.Name $sourcePath
            $twoSelection = Select-Designer $form.Name $sourcePath @('FirstNested','SecondNested') $previousSelection
            $staleSelection = $false
            try { Invoke-DesignerClipboard $form.Name 'cut' $previousSelection | Out-Null }
            catch { $staleSelection = $_.Exception.Message -match 'tree or selection changed' }
            $unknownRefused = $false
            try { Select-Designer $form.Name $sourcePath @('MissingControl') (Read-DesignerState $form.Name $sourcePath) | Out-Null }
            catch { $unknownRefused = $_.Exception.Message -match 'outside the selected container' }
            $preservedSelection = Read-DesignerState $form.Name $sourcePath
            if (-not $oneSelection.Verified -or -not $twoSelection.Verified -or -not $staleSelection -or -not $unknownRefused -or @($preservedSelection.Selected).Count -ne 2) { throw 'Nested selection guards failed.' }
            $nestedCopy = Invoke-DesignerClipboard $form.Name 'copy' $preservedSelection
            $nestedPaste = Invoke-DesignerClipboard $target.Name 'paste' (Read-DesignerState $target.Name $targetPath)
            $pastedNames = @($page.Controls | ForEach-Object { $_.Name })
            if ($page.Controls.Count -ne 2 -or $page.Controls.Item('FirstNested').Caption -cne 'First nested caption' -or $page.Controls.Item('SecondNested').Caption -cne 'Second nested caption') { throw 'Frame to Page native copy failed.' }
            $nestedCut = Invoke-DesignerClipboard $form.Name 'cut' (Read-DesignerState $form.Name $sourcePath)
            if ($frame.Controls.Count -ne 1) { throw 'Nested cut did not remove both selected controls.' }
            [System.Windows.Forms.Clipboard]::SetText('Clipboard replaced after nested cut')
            $wrongScope = $false
            $current = Read-DesignerState $target.Name $targetPath
            try { Invoke-Session @{ Command='restore_form_clipboard'; Project=$project.Name; Form=$target.Name; ParentPath=$targetPath; DesignerClipboardRecoveryId=$nestedCut.DesignerClipboardRecoveryId; ExpectedDesignerSelectionVersion=$current.SelectionVersion; ExpectedClipboardVersion=$current.ClipboardVersion } | Out-Null }
            catch { $wrongScope = $_.Exception.Message -match 'another live form or container' }
            if (-not $wrongScope) { throw 'Recovery did not reject another form.' }
            $current = Read-DesignerState $form.Name $sourcePath
            $originalCaption = $frame.Caption; $frame.Caption = 'Intervening user edit'
            $changed = Read-DesignerState $form.Name $sourcePath
            $interveningRejected = $false
            try { Invoke-Session @{ Command='recover_form_cut'; Project=$project.Name; Form=$form.Name; ParentPath=$sourcePath; DesignerClipboardRecoveryId=$nestedCut.DesignerClipboardRecoveryId; ExpectedDesignerSelectionVersion=$changed.SelectionVersion; ExpectedClipboardVersion=$changed.ClipboardVersion } | Out-Null }
            catch { $interveningRejected = $_.Exception.Message -match 'changed after the cut' }
            if (-not $interveningRejected -or $frame.Controls.Count -ne 1) { throw 'Intervening edit guard failed.' }
            $frame.Caption = $originalCaption
            $current = Read-DesignerState $form.Name $sourcePath
            $nestedRecovery = Invoke-Session @{ Command='recover_form_cut'; Project=$project.Name; Form=$form.Name; ParentPath=$sourcePath; DesignerClipboardRecoveryId=$nestedCut.DesignerClipboardRecoveryId; ExpectedDesignerSelectionVersion=$current.SelectionVersion; ExpectedClipboardVersion=$current.ClipboardVersion }
            if (-not $nestedRecovery.RestoredNamesGeometryAndTabOrder -or $frame.Controls.Count -ne 3 -or (Read-Placement $frame) -cne $originalPlacement) { throw ('Nested geometry recovery failed: ' + $nestedRecovery.NativeError) }
            $clear = Select-Designer $form.Name $sourcePath @() (Read-DesignerState $form.Name $sourcePath)
            if (-not $clear.Verified -or @($clear.After.Selected).Count -ne 0) { throw 'Clear nested selection failed.' }
            if ($form.CodeModule.Lines(1,$form.CodeModule.CountOfLines) -cne $sourceCode -or $target.CodeModule.CountOfLines -ne 0) { throw 'Source event code changed or was unexpectedly copied.' }
            $book.SaveAs($document,52)
            $book.Close($false); $book = $null
            $book = $excel.Workbooks.Open($document)
            $reopenedSource = $book.VBProject.VBComponents.Item('ClipboardDesignerSource')
            $reopenedTarget = $book.VBProject.VBComponents.Item('ClipboardDesignerTarget')
            $savedFrame = $reopenedSource.Designer.Controls.Item('SourceFrame')
            $savedPage = $reopenedTarget.Designer.Controls.Item('TargetPages').Pages.Item(0)
            if ($savedFrame.Controls.Count -ne 3 -or $savedPage.Controls.Count -ne 2 -or $savedPage.Controls.Item('FirstNested').Caption -cne 'First nested caption' -or $savedPage.Controls.Item('SecondNested').Caption -cne 'Second nested caption') { throw 'Saved nested control state was not preserved.' }
            if ((Read-Placement $savedFrame) -cne $originalPlacement) { throw 'Saved recovered geometry or tab order differs.' }
            if ($reopenedSource.CodeModule.Lines(1,$reopenedSource.CodeModule.CountOfLines) -cne $sourceCode -or $reopenedTarget.CodeModule.CountOfLines -ne 0) { throw 'Saved event code was not preserved.' }
            $proof = [pscustomobject]@{ RecoveryAfterClipboardReplacement=$true; WrongRecoveryScopeRejected=$wrongScope; RecoveryCardClickVerified=$true; GeometryAndTabOrderRestored=$true; RecoveryReplayRejected=$replayRejected; InterveningEditRejected=$interveningRejected; SaveReopenVerified=$true; EventCodePreserved=$true; NestedCopy=$nestedCopy; NestedPaste=$nestedPaste; NestedCut=$nestedCut; NestedRecovery=$nestedRecovery; NestedSelectionVerified=$true; StaleSelectionRejected=$staleSelection; UnknownControlRejected=$unknownRefused; Copy=$copy; Paste=$paste; Cut=$cut; Undo=$undo; UndoAvailable=$undoAvailable; UndoRefused=$undoRefused; RecoveryPaste=$recovery; TargetCount=$countAfterPaste; PastedCaption=$pastedCaption; SourceAfterCut=$countAfterCut; SourceAfterRecovery=$restoredCount; StaleTreeRejected=$staleTree; StaleClipboardRejected=$staleClipboard; ClipboardRestored=$false; Mvid=$assembly.ManifestModule.ModuleVersionId.ToString() }
            if ($countAfterPaste -ne 1 -or $pastedCaption -cne 'Changed caption' -or $countAfterCut -ne 0 -or $restoredCount -ne 1 -or -not $staleTree -or -not $staleClipboard -or -not $copy.ClipboardChanged -or -not $paste.DesignerChangeObserved -or -not $cut.DesignerChangeObserved) { throw 'Native Designer clipboard verification failed.' }
        } finally {
            try {
                if ($clipboardTouched) {
                    if ($backupValues.Count -eq 0) { [System.Windows.Forms.Clipboard]::Clear() }
                    else { [System.Windows.Forms.Clipboard]::SetDataObject($backup, $true) }
                    $restoredClipboard = [System.Windows.Forms.Clipboard]::GetDataObject()
                    foreach ($format in $backupValues.Keys) {
                        $expected = $backupValues[$format]; $actual = $restoredClipboard.GetData($format,$false)
                        if ($expected -is [IO.MemoryStream]) {
                            if ($actual -isnot [IO.MemoryStream] -or [Convert]::ToBase64String($expected.ToArray()) -cne [Convert]::ToBase64String($actual.ToArray())) { throw "Clipboard format restoration failed: $format" }
                        } elseif ($actual -cne $expected) { throw "Clipboard format restoration failed: $format" }
                    }
                    if ($null -ne $proof) { $proof.ClipboardRestored = $true }
                }
            } finally {
                if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
                else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
                try { if ($null -ne $book) { $book.Close($false) } }
                finally {
                    try { if ($null -ne $excel) { $excel.Quit() } }
                    finally { if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) { Stop-Process -Id $probeProcess.Id -Force } }
                }
            }
        }
        $proof | ConvertTo-Json -Depth 40 | Set-Content (Join-Path $directory 'native-form-clipboard.json') -Encoding UTF8
        $proof | Select-Object RecoveryCardClickVerified,GeometryAndTabOrderRestored,RecoveryReplayRejected,InterveningEditRejected,RecoveryAfterClipboardReplacement,WrongRecoveryScopeRejected,SaveReopenVerified,EventCodePreserved,NestedSelectionVerified,StaleSelectionRejected,UnknownControlRejected,TargetCount,PastedCaption,SourceAfterCut,SourceAfterRecovery,StaleTreeRejected,StaleClipboardRejected,ClipboardRestored,UndoAvailable,UndoRefused | ConvertTo-Json
    }
    'ClipboardRecoveryDiscovery' {
        $ErrorActionPreference = 'Stop'
        . Initialize-NativeUserFormProbe
        $probeProcess = $null
        $otherBook = $null
        $excel = $null; $book = $null; $form = $null; $session = $null
        function Invoke-Session([hashtable]$Fields) {
            $request = New-Object VBAi.Request
            foreach ($key in $Fields.Keys) { $request.$key = $Fields[$key] }
            Write-Output ("Request: " + $request.Command + " " + $request.Action) | Out-Host
            $response = $script:session.Execute($request)
            if (-not $response.Ok) { throw $response.Error }
            return $response.Data
        }
        Add-Type -AssemblyName System.Windows.Forms
        $backup = New-Object System.Windows.Forms.DataObject
        $backupValues = @{}
        $currentClipboard = [System.Windows.Forms.Clipboard]::GetDataObject()
        if ($null -ne $currentClipboard) {
            foreach ($format in $currentClipboard.GetFormats($false)) {
                $value = $currentClipboard.GetData($format, $false)
                if ($null -eq $value) { continue }
                if ($value -is [IO.MemoryStream]) { $value = New-Object IO.MemoryStream -ArgumentList (,($value.ToArray())) }
                elseif ($value -isnot [string]) { throw "Clipboard test refused: unsupported backup format $format." }
                $backup.SetData($format, $false, $value); $backupValues[$format] = $value
            }
        }
        $clipboardTouched = $false
        $proof = $null
        try {
            if ($priorAccess -ne 1) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null }
            $excel = New-Object -ComObject Excel.Application
            $probeProcess = Get-Process EXCEL -ErrorAction Stop
            if (@($probeProcess).Count -ne 1) { throw "Excel isolation was lost." }
            $excel.Visible = $true
            $book = $excel.Workbooks.Add()
            $vbe = $excel.GetType().InvokeMember('VBE', [Reflection.BindingFlags]::GetProperty, $null, $excel, $null)
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
            $project = $book.VBProject
            $vbe.MainWindow.Visible = $true
            $form = $project.VBComponents.Add(3); $form.Name = 'ClipboardDesignerSource'
            $button = $form.Designer.Controls.Add('Forms.CommandButton.1','SourceButton',$true)
            $form.CodeModule.AddFromString("Private Sub SourceButton_Click()`r`n    Debug.Print ""handler retained""`r`nEnd Sub")
            $sourceCode = $form.CodeModule.Lines(1,$form.CodeModule.CountOfLines)
            $button.Caption = 'Clipboard native test'; $button.Left = 30; $button.Top = 30
            $form.DesignerWindow().Visible = $true; $form.DesignerWindow().SetFocus()
            function Pump-Designer { for ($i=0; $i -lt 5; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 50 } }
            Pump-Designer
            function Read-DesignerState($name,$parent) { Invoke-Session @{ Command='form_clipboard_state'; Project=$project.Name; Form=$name; ParentPath=$parent } }
            function Invoke-DesignerClipboard($name,$action,$state) {
                Invoke-Session @{ Command='native_form_clipboard'; Project=$project.Name; Form=$name; Action=$action; ExpectedDesignerSelectionVersion=$state.SelectionVersion; ExpectedClipboardVersion=$state.ClipboardVersion; ParentPath=$state.ParentPath }
            }
            function Select-Designer($name,$parent,[string[]]$names,$state) {
                Invoke-Session @{ Command='select_form_controls'; Project=$project.Name; Form=$name; ParentPath=$parent; Items=$names; ExpectedDesignerSelectionVersion=$state.SelectionVersion }
            }
            $initialSelection = Select-Designer $form.Name '' @('SourceButton') (Read-DesignerState $form.Name)
            if (-not $initialSelection.Verified) { throw 'Initial programmatic selection failed.' }
            $source = Read-DesignerState $form.Name
            if (@($source.Selected).Count -ne 1 -or $source.Selected[0] -ne 'SourceButton') { throw 'Designer selection not read correctly.' }
            $button.Caption = 'Changed caption'
            $staleTree = $false
            try { Invoke-DesignerClipboard $form.Name 'copy' $source | Out-Null }
            catch { $staleTree = $_.Exception.Message -match 'tree or selection changed' }
            $source = Read-DesignerState $form.Name
            $clipboardTouched = $true
            [System.Windows.Forms.Clipboard]::SetText('Temporary clipboard guard test')
            $staleClipboard = $false
            try { Invoke-DesignerClipboard $form.Name 'copy' $source | Out-Null }
            catch { $staleClipboard = $_.Exception.Message -match 'Clipboard changed' -or $_.Exception.Message -match 'tree or selection changed' }
            $copy = Invoke-DesignerClipboard $form.Name 'copy' (Read-DesignerState $form.Name)
            $detachedData = New-Object System.Windows.Forms.DataObject
            $formatInfo = @()
            $nativeData = [System.Windows.Forms.Clipboard]::GetDataObject()
            foreach ($format in $nativeData.GetFormats($false)) {
                $value = $nativeData.GetData($format,$false)
                $formatInfo += [pscustomobject]@{ Format=$format; Type=$(if ($null -eq $value) { "null" } else { $value.GetType().FullName }) }
                if ($null -eq $value) { continue }
                if ($value -is [IO.MemoryStream]) { $value=New-Object IO.MemoryStream -ArgumentList (,($value.ToArray())) }
                elseif ($value -isnot [string]) { throw "Unsupported native clipboard format type: $format" }
                $detachedData.SetData($format,$false,$value)
            }
            $formatInfo | ConvertTo-Json | Set-Content (Join-Path $directory 'native-form-clipboard-formats.json') -Encoding UTF8
            [System.Windows.Forms.Clipboard]::SetText('Clipboard overwritten after native copy')
            [System.Windows.Forms.Clipboard]::SetDataObject($detachedData,$true)
            $formatInfo | ConvertTo-Json | Set-Content (Join-Path $directory 'native-form-clipboard-formats.json') -Encoding UTF8
            $target = $project.VBComponents.Add(3); $target.Name = 'ClipboardDesignerTarget'
            $target.DesignerWindow().Visible = $true
            $form.DesignerWindow().SetFocus(); Pump-Designer
            $paste = Invoke-DesignerClipboard $target.Name 'paste' (Read-DesignerState $target.Name)
            Pump-Designer
            $countAfterPaste = $target.Designer.Controls.Count
            $pastedCaption = $target.Designer.Controls.Item(0).Caption
            $target.DesignerWindow().SetFocus(); Pump-Designer
            $cut = Invoke-DesignerClipboard $form.Name 'cut' (Read-DesignerState $form.Name)
            Pump-Designer
            $countAfterCut = $form.Designer.Controls.Count
            $tree = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
            $undoAvailable = [bool](($tree.Properties | Where-Object Name -eq CanUndo).Value)
            $undo = $null; $undoRefused = $false; $recovery = $null
            if ($undoAvailable) {
                $undo = Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='undo'; ExpectedTreeVersion=$tree.TreeVersion }
            } else {
                try { Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='undo'; ExpectedTreeVersion=$tree.TreeVersion } | Out-Null }
                catch { $undoRefused = $_.Exception.Message -match 'unavailable' }
                if (-not $undoRefused) { throw 'Unavailable undo was not refused.' }
                $recovery = Invoke-DesignerClipboard $form.Name 'paste' (Read-DesignerState $form.Name)
            }
            $restoredCount = $form.Designer.Controls.Count
            $frame = $form.Designer.Controls.Add('Forms.Frame.1','SourceFrame',$true)
            $first = $frame.Controls.Add('Forms.Label.1','FirstNested',$true); $first.Caption='First nested caption'
            $second = $frame.Controls.Add('Forms.CommandButton.1','SecondNested',$true); $second.Caption='Second nested caption'
            $multi = $target.Designer.Controls.Add('Forms.MultiPage.1','TargetPages',$true)
            $page = $multi.Pages.Item(0)
            $sourcePath = 'Controls/SourceFrame'
            $targetPath = 'Controls/TargetPages/Pages/' + $page.Name
            $oneSelection = Select-Designer $form.Name $sourcePath @('FirstNested') (Read-DesignerState $form.Name $sourcePath)
            $previousSelection = Read-DesignerState $form.Name $sourcePath
            $twoSelection = Select-Designer $form.Name $sourcePath @('FirstNested','SecondNested') $previousSelection
            $staleSelection = $false
            try { Invoke-DesignerClipboard $form.Name 'cut' $previousSelection | Out-Null }
            catch { $staleSelection = $_.Exception.Message -match 'tree or selection changed' }
            $unknownRefused = $false
            try { Select-Designer $form.Name $sourcePath @('MissingControl') (Read-DesignerState $form.Name $sourcePath) | Out-Null }
            catch { $unknownRefused = $_.Exception.Message -match 'outside the selected container' }
            $preservedSelection = Read-DesignerState $form.Name $sourcePath
            if (-not $oneSelection.Verified -or -not $twoSelection.Verified -or -not $staleSelection -or -not $unknownRefused -or @($preservedSelection.Selected).Count -ne 2) { throw 'Nested selection guards failed.' }
            $nestedCopy = Invoke-DesignerClipboard $form.Name 'copy' $preservedSelection
            $nestedPaste = Invoke-DesignerClipboard $target.Name 'paste' (Read-DesignerState $target.Name $targetPath)
            $pastedNames = @($page.Controls | ForEach-Object { $_.Name })
            if ($page.Controls.Count -ne 2 -or $page.Controls.Item('FirstNested').Caption -cne 'First nested caption' -or $page.Controls.Item('SecondNested').Caption -cne 'Second nested caption') { throw 'Frame to Page native copy failed.' }
            $nestedCut = Invoke-DesignerClipboard $form.Name 'cut' (Read-DesignerState $form.Name $sourcePath)
            if ($frame.Controls.Count -ne 0) { throw 'Nested cut did not remove both selected controls.' }
            $nestedRecovery = Invoke-DesignerClipboard $form.Name 'paste' (Read-DesignerState $form.Name $sourcePath)
            if ($frame.Controls.Count -ne 2) { throw 'Nested recovery failed.' }
            $clear = Select-Designer $form.Name $sourcePath @() (Read-DesignerState $form.Name $sourcePath)
            if (-not $clear.Verified -or @($clear.After.Selected).Count -ne 0) { throw 'Clear nested selection failed.' }
            if ($form.CodeModule.Lines(1,$form.CodeModule.CountOfLines) -cne $sourceCode -or $target.CodeModule.CountOfLines -ne 0) { throw 'Source event code changed or was unexpectedly copied.' }
            $book.SaveAs($document,52)
            $book.Close($false); $book = $null
            $book = $excel.Workbooks.Open($document)
            $reopenedSource = $book.VBProject.VBComponents.Item('ClipboardDesignerSource')
            $reopenedTarget = $book.VBProject.VBComponents.Item('ClipboardDesignerTarget')
            $savedFrame = $reopenedSource.Designer.Controls.Item('SourceFrame')
            $savedPage = $reopenedTarget.Designer.Controls.Item('TargetPages').Pages.Item(0)
            if ($savedFrame.Controls.Count -ne 2 -or $savedPage.Controls.Count -ne 2 -or $savedPage.Controls.Item('FirstNested').Caption -cne 'First nested caption' -or $savedPage.Controls.Item('SecondNested').Caption -cne 'Second nested caption') { throw 'Saved nested control state was not preserved.' }
            if ($reopenedSource.CodeModule.Lines(1,$reopenedSource.CodeModule.CountOfLines) -cne $sourceCode -or $reopenedTarget.CodeModule.CountOfLines -ne 0) { throw 'Saved event code was not preserved.' }
            $proof = [pscustomobject]@{ SaveReopenVerified=$true; EventCodePreserved=$true; NestedCopy=$nestedCopy; NestedPaste=$nestedPaste; NestedCut=$nestedCut; NestedRecovery=$nestedRecovery; NestedSelectionVerified=$true; StaleSelectionRejected=$staleSelection; UnknownControlRejected=$unknownRefused; Copy=$copy; Paste=$paste; Cut=$cut; Undo=$undo; UndoAvailable=$undoAvailable; UndoRefused=$undoRefused; RecoveryPaste=$recovery; TargetCount=$countAfterPaste; PastedCaption=$pastedCaption; SourceAfterCut=$countAfterCut; SourceAfterRecovery=$restoredCount; StaleTreeRejected=$staleTree; StaleClipboardRejected=$staleClipboard; ClipboardRestored=$false; Mvid=$assembly.ManifestModule.ModuleVersionId.ToString() }
            if ($countAfterPaste -ne 1 -or $pastedCaption -cne 'Changed caption' -or $countAfterCut -ne 0 -or $restoredCount -ne 1 -or -not $staleTree -or -not $staleClipboard -or -not $copy.ClipboardChanged -or -not $paste.DesignerChangeObserved -or -not $cut.DesignerChangeObserved) { throw 'Native Designer clipboard verification failed.' }
        } finally {
            try {
                if ($clipboardTouched) {
                    if ($backupValues.Count -eq 0) { [System.Windows.Forms.Clipboard]::Clear() }
                    else { [System.Windows.Forms.Clipboard]::SetDataObject($backup, $true) }
                    $restoredClipboard = [System.Windows.Forms.Clipboard]::GetDataObject()
                    foreach ($format in $backupValues.Keys) {
                        $expected = $backupValues[$format]; $actual = $restoredClipboard.GetData($format,$false)
                        if ($expected -is [IO.MemoryStream]) {
                            if ($actual -isnot [IO.MemoryStream] -or [Convert]::ToBase64String($expected.ToArray()) -cne [Convert]::ToBase64String($actual.ToArray())) { throw "Clipboard format restoration failed: $format" }
                        } elseif ($actual -cne $expected) { throw "Clipboard format restoration failed: $format" }
                    }
                    if ($null -ne $proof) { $proof.ClipboardRestored = $true }
                }
            } finally {
                if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
                else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
                try { if ($null -ne $book) { $book.Close($false) } }
                finally {
                    try { if ($null -ne $excel) { $excel.Quit() } }
                    finally { if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) { Stop-Process -Id $probeProcess.Id -Force } }
                }
            }
        }
        $proof | ConvertTo-Json -Depth 40 | Set-Content (Join-Path $directory 'native-form-clipboard-recovery-discovery.json') -Encoding UTF8
        $proof | Select-Object SaveReopenVerified,EventCodePreserved,NestedSelectionVerified,StaleSelectionRejected,UnknownControlRejected,TargetCount,PastedCaption,SourceAfterCut,SourceAfterRecovery,StaleTreeRejected,StaleClipboardRejected,ClipboardRestored,UndoAvailable,UndoRefused | ConvertTo-Json
    }
    'History' {
        $ErrorActionPreference = 'Stop'
        . Initialize-NativeUserFormProbe
        $probeProcess = $null
        $otherBook = $null
        $excel = $null; $book = $null; $form = $null; $session = $null
        function Invoke-Session([hashtable]$Fields) {
            $request = New-Object VBAi.Request
            foreach ($key in $Fields.Keys) { $request.$key = $Fields[$key] }
            Write-Output ("Request: " + $request.Command + " " + $request.Action) | Out-Host
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
            $form = $project.VBComponents.Add(3); $form.Name = 'NativeDesignerHistory'
            $form.DesignerWindow().Visible = $true
            $button = $form.Designer.Controls.Add('Forms.CommandButton.1', 'HistoryButton', $true)
            $button.Left = 20; $button.Top = 20; $button.Width = 70; $button.Height = 24
            $button.Left = 55
            $second = $form.Designer.Controls.Add('Forms.CommandButton.1', 'SecondButton', $true)
            $second.Left = 100; $second.Top = 60; $second.Width = 70; $second.Height = 24
            $form.DesignerWindow().SetFocus()
            Add-Type -AssemblyName System.Windows.Forms
            function Pump-Designer { for ($i=0; $i -lt 5; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 50 } }
            Pump-Designer
            $vbe.CommandBars.FindControl(1,756).Execute()
            Pump-Designer
            $original = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
            $align = $vbe.CommandBars.FindControl(1,664)
            if (-not $align.Enabled) { throw 'Native Align Left is disabled.' }
            $align.Execute()
            Pump-Designer
            $before = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
            $commands = Invoke-Session @{ Command='list_commands' }
            $commands | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $directory 'native-form-history-commands.json') -Encoding UTF8
            $otherBook = $excel.Workbooks.Add()
            $otherBook.VBProject.Name = 'OtherDesignerProject'
            $otherForm = $otherBook.VBProject.VBComponents.Add(3); $otherForm.Name = 'OtherDesigner'
            $otherButton = $otherForm.Designer.Controls.Add('Forms.CommandButton.1', 'UntouchedButton', $true)
            $otherButton.Left = 80; $otherButton.Top = 40
            $otherForm.DesignerWindow().Visible = $true; $otherForm.DesignerWindow().SetFocus()
            Pump-Designer
            $otherBefore = Invoke-Session @{ Command='form_tree'; Project='OtherDesignerProject'; Form=$otherForm.Name }
            $before = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
            $nativeResult = Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='undo'; ExpectedTreeVersion=$before.TreeVersion }
            Pump-Designer
            $after = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
            $staleRejected = $false
            try { Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='redo'; ExpectedTreeVersion=$before.TreeVersion } | Out-Null }
            catch { $staleRejected = $_.Exception.Message -match 'changed since' }
            $redoResult = Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='redo'; ExpectedTreeVersion=$after.TreeVersion }
            Pump-Designer
            $redone = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
            $otherAfter = Invoke-Session @{ Command='form_tree'; Project='OtherDesignerProject'; Form=$otherForm.Name }
            $unavailableRejected = $false
            try { Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='redo'; ExpectedTreeVersion=$redone.TreeVersion } | Out-Null }
            catch { $unavailableRejected = $_.Exception.Message -match 'unavailable' }
            $form.DesignerWindow().SetFocus()
            Pump-Designer
            $vbe.CommandBars.FindControl(1,756).Execute()
            Pump-Designer
            $clear = $vbe.CommandBars.FindControl(1,478)
            if (-not $clear.Enabled) { throw 'Native Delete is disabled.' }
            $clear.Execute()
            Pump-Designer
            $deleted = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
            $undoDelete = Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='undo'; ExpectedTreeVersion=$deleted.TreeVersion }
            $restoredControls = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
            [pscustomobject]@{ Deleted=$deleted; UndoDelete=$undoDelete; Restored=$restoredControls } | ConvertTo-Json -Depth 30 | Set-Content (Join-Path $directory 'native-form-delete-stage.json') -Encoding UTF8
            $redoDeletionAvailable = [bool](($restoredControls.Properties | Where-Object Name -eq CanRedo).Value)
            $redoDelete = $null; $redoDeleteRefused = $false
            if ($redoDeletionAvailable) {
                $redoDelete = Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='redo'; ExpectedTreeVersion=$restoredControls.TreeVersion }
            } else {
                try { Invoke-Session @{ Command='native_form_history'; Project=$project.Name; Form=$form.Name; Action='redo'; ExpectedTreeVersion=$restoredControls.TreeVersion } | Out-Null }
                catch { $redoDeleteRefused = $_.Exception.Message -match 'unavailable' }
                if (-not $redoDeleteRefused) { throw 'Unavailable native redo was not refused.' }
            }
            $deletedAgain = Invoke-Session @{ Command='form_tree'; Project=$project.Name; Form=$form.Name }
            if ($deleted.NodeCount -ne 0 -or $restoredControls.NodeCount -ne 2 -or -not $undoDelete.Verified) { throw 'Native control deletion undo was not verified.' }
            if ($redoDeletionAvailable -and ($deletedAgain.NodeCount -ne 0 -or -not $redoDelete.Verified)) { throw 'Available native redo was not verified.' }
            if (-not $redoDeletionAvailable -and $deletedAgain.NodeCount -ne 2) { throw 'Refused redo changed the controls.' }
            [pscustomobject]@{ RedoDeletionAvailable=$redoDeletionAvailable; RedoDeletionRefused=$redoDeleteRefused; UndoDelete=$undoDelete; RedoDelete=$redoDelete; RestoredControls=$restoredControls.NodeCount; DeletedAgain=$deletedAgain.NodeCount; StaleRejected=$staleRejected; UnavailableRejected=$unavailableRejected; OtherFormPreserved=($otherBefore.TreeVersion -eq $otherAfter.TreeVersion); UndoReturn=$nativeResult; RedoReturn=$redoResult; Redone=$redone; Original=$original; Before=$before; After=$after; Commands=$commands } | ConvertTo-Json -Depth 25 | Set-Content (Join-Path $directory 'native-form-history.json') -Encoding UTF8
            if (-not $nativeResult.Verified -or -not $redoResult.Verified -or -not $staleRejected -or -not $unavailableRejected -or $otherBefore.TreeVersion -ne $otherAfter.TreeVersion) { throw 'Designer history targeting or guards failed.' }
            $leftAfter = ($after.Controls | Where-Object Name -eq SecondButton).Properties | Where-Object Name -eq Left
            $leftRedone = ($redone.Controls | Where-Object Name -eq SecondButton).Properties | Where-Object Name -eq Left
            if ($leftAfter.Value -ne 100 -or $leftRedone.Value -ne 55) { throw 'Designer geometry was not restored.' }
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
}

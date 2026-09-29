param(
    [Parameter(Mandatory = $true)][string]$AssemblyPath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [switch]$UseBridge,
    [switch]$AllowTemporaryVbaAccess
)
$ErrorActionPreference = 'Stop'
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
    $module = $project.VBComponents.Add(1); $module.Name = 'DefinitionCaller'
    $module.CodeModule.AddFromString("Public Sub Caller()`r`n    TargetMethod`r`nEnd Sub")
    $target = $project.VBComponents.Add(1); $target.Name = 'DefinitionTarget'
    $target.CodeModule.AddFromString("Public Sub TargetMethod()`r`n    Debug.Print 42`r`nEnd Sub")
    $pane = $module.CodeModule.CodePane; $pane.Show(); $pane.SetSelection(2,5,2,17)
    $command = $vbe.CommandBars.FindControl(1,939)
    if ($null -eq $command -or -not $command.Enabled) { throw 'Definition command unavailable.' }
    $code = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $module.Name }
    $queued = Invoke-Session @{ Command = 'native_code_navigation'; Action = 'definition'; Project = $project.Name; Module = $module.Name; ExpectedSha256 = $code.Sha256; ExpectedMode = 2; StartLine = 2; StartColumn = 5; EndColumn = 17; Expression = 'TargetMethod'; ControlCaption = $command.Caption }
    if ($queued.State -ne 'Queued' -or $queued.CommandCompleted) { throw 'Navigation did not return before execution.' }
    if (-not $UseBridge) { $probeContext.RunNext() }
    else { Start-Sleep -Milliseconds 300 }
    $completed = Invoke-Session @{ Command = 'native_code_navigation'; Action = 'status'; Query = $queued.OperationId }
    if ($completed.State -ne 'Completed' -or -not $completed.NavigationObserved) {
        if ($UseBridge) {
            $diagnostics = [pscustomobject]@{ Dialog = (Invoke-Session @{ Command = 'debug_dialog' }); Panes = (Invoke-Session @{ Command = 'code_panes' }); NativeActiveModule = $vbe.ActiveCodePane.CodeModule.Parent.Name }
            $diagnostics | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $directory 'installed-navigation-diagnostics.json') -Encoding UTF8
        }
        $completed | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'native-definition-failure.json') -Encoding UTF8; throw ('Definition failed: ' + ($completed | ConvertTo-Json -Depth 5)) }
    $definitionPane = $vbe.ActiveCodePane
    [int]$line = 0; [int]$column = 0; [int]$end = 0; [int]$endColumn = 0
    $definitionPane.GetSelection([ref]$line,[ref]$column,[ref]$end,[ref]$endColumn)
    if ($definitionPane.CodeModule.Parent.Name -ne 'DefinitionTarget' -or $line -lt 1 -or $line -gt 3) { throw "Definition did not reach the target procedure: $($definitionPane.CodeModule.Parent.Name):$line" }
    $definition = [pscustomobject]@{ CommandId = 939; Caption = $command.Caption; Module = $definitionPane.CodeModule.Parent.Name; Line = $line; Column = $column; SourceLine = $definitionPane.CodeModule.Lines($line,1) }
    $last = $vbe.CommandBars.FindControl(1,1822)
    if ($null -eq $last -or -not $last.Enabled) { throw 'Last Position command unavailable.' }
    $code = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $target.Name }
    $queuedLast = Invoke-Session @{ Command = 'native_code_navigation'; Action = 'last_position'; Project = $project.Name; Module = $target.Name; ExpectedSha256 = $code.Sha256; ExpectedMode = 2; StartLine = $line; StartColumn = $column; EndColumn = $endColumn; ControlCaption = $last.Caption }
    if (-not $UseBridge) { $probeContext.RunNext() }
    else { Start-Sleep -Milliseconds 300 }
    $completedLast = Invoke-Session @{ Command = 'native_code_navigation'; Action = 'status'; Query = $queuedLast.OperationId }
    if ($completedLast.State -ne 'Completed' -or -not $completedLast.NavigationObserved) { throw 'Last Position operation failed.' }
    $returnedPane = $vbe.ActiveCodePane
    $returnedPane.GetSelection([ref]$line,[ref]$column,[ref]$end,[ref]$endColumn)
    if ($returnedPane.CodeModule.Parent.Name -ne 'DefinitionCaller' -or $line -ne 2) { throw 'Last Position did not return to the call site.' }
    $lastPosition = [pscustomobject]@{ CommandId = 1822; Caption = $last.Caption; Module = $returnedPane.CodeModule.Parent.Name; Line = $line; Column = $column }
    if (-not $UseBridge) {
    $code = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $module.Name }
    $pendingRequest = @{ Command = 'native_code_navigation'; Action = 'definition'; Project = $project.Name; Module = $module.Name; ExpectedSha256 = $code.Sha256; ExpectedMode = 2; StartLine = 2; StartColumn = 5; EndColumn = 17; Expression = 'TargetMethod'; ControlCaption = $command.Caption }
    $pending = Invoke-Session $pendingRequest
    $duplicateRejected = $false
    try { $null = Invoke-Session $pendingRequest } catch { if ($_.Exception.Message.Contains('still pending')) { $duplicateRejected = $true } else { throw } }
    if (-not $duplicateRejected) { throw 'Concurrent native navigation was accepted.' }
    $module.CodeModule.InsertLines(1, "' external change before queued execution")
    if (-not $UseBridge) { $probeContext.RunNext() }
    else { Start-Sleep -Milliseconds 300 }
    $rejected = Invoke-Session @{ Command = 'native_code_navigation'; Action = 'status'; Query = $pending.OperationId }
    if ($rejected.State -ne 'Failed' -or $rejected.CommandCompleted -or -not $rejected.NativeError.Contains('module changed')) { throw 'Queued navigation did not reject changed source.' }
    }
    if ($UseBridge) {
        $viewResults = @()
        foreach ($viewAction in @('procedure','module')) {
            $viewLayout = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
            $viewPane = $viewLayout.Panes[0]
            $viewGuardRejected = $false
            try { Invoke-Session @{ Command = 'set_code_view'; Project = $project.Name; Module = $module.Name; Pane = $viewPane.Pane; ExpectedWindowVersion = 'stale'; ExpectedSha256 = $viewPane.State.Sha256; ExpectedMode = 2; StartLine = 2; Action = $viewAction } | Out-Null }
            catch { if ($_.Exception.Message -notmatch 'The pane changed') { throw }; $viewGuardRejected = $true }
            if (-not $viewGuardRejected) { throw 'Stale view revision was accepted.' }
            $viewResult = Invoke-Session @{ Command = 'set_code_view'; Project = $project.Name; Module = $module.Name; Pane = $viewPane.Pane; ExpectedWindowVersion = $viewPane.WindowVersion; ExpectedSha256 = $viewPane.State.Sha256; ExpectedMode = 2; StartLine = 2; Action = $viewAction }
            $viewResults += $viewResult
            if (-not $viewResult.Verified) { throw 'Installed native code-view switch was not verified.' }
            $sameLayout = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
            $samePane = $sameLayout.Panes[0]
            $sameView = Invoke-Session @{ Command = 'set_code_view'; Project = $project.Name; Module = $module.Name; Pane = $samePane.Pane; ExpectedWindowVersion = $samePane.WindowVersion; ExpectedSha256 = $samePane.State.Sha256; ExpectedMode = 2; StartLine = 2; Action = $viewAction }
            if (-not $sameView.Verified -or $sameView.Changed) { throw 'Repeated view request was not a verified no-op.' }
            $viewResults += [pscustomobject]@{ StaleRevisionRejected = $viewGuardRejected; Repeat = $sameView }

        }
        $viewResults | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $directory 'installed-code-view-cycle.json') -Encoding UTF8
        $splitControl = $vbe.CommandBars.FindControl(1,302)
        $splitCode = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $module.Name }
        $splitCreated = Invoke-Session @{ Command = 'set_code_split'; Project = $project.Name; Module = $module.Name; ExpectedSha256 = $splitCode.Sha256; ExpectedMode = 2; StartLine = 2; Action = 'split'; ControlCaption = $splitControl.Caption }
        Start-Sleep -Milliseconds 200
        $splitObserved = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
        [pscustomobject]@{ Command = $splitCreated; Observed = $splitObserved } | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $directory 'installed-split-view-precondition.json') -Encoding UTF8
        if ($splitObserved.Panes.Count -ne 2) { throw 'Split view fixture did not split.' }
        $staleSiblingLayout = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
        $siblingPanes = @($vbe.CodePanes | Where-Object { $_.CodeModule.Parent.Name -eq $module.Name })
        $siblingPanes[1].SetSelection(3,1,3,1)
        $staleSiblingRejected = $false
        $unchangedPane = $staleSiblingLayout.Panes[0]
        try { Invoke-Session @{ Command = 'set_code_view'; Project = $project.Name; Module = $module.Name; Pane = $unchangedPane.Pane; ExpectedWindowVersion = $unchangedPane.WindowVersion; ExpectedSha256 = $unchangedPane.State.Sha256; ExpectedMode = 2; StartLine = 2; Action = 'procedure' } | Out-Null }
        catch { if ($_.Exception.Message -notmatch 'A pane in the code window changed') { throw }; $staleSiblingRejected = $true }
        if (-not $staleSiblingRejected) { throw 'A changed sibling pane did not invalidate the window operation.' }
        $afterSiblingGuard = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
        if (@($afterSiblingGuard.Panes | Where-Object { $_.State.View -ne 1 -or $_.State.Sha256 -ne $splitCode.Sha256 }).Count) { throw 'Rejected sibling change altered the view or source.' }
        [pscustomobject]@{ Rejected = $staleSiblingRejected; After = $afterSiblingGuard } | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $directory 'installed-split-code-view-guard.json') -Encoding UTF8
        $splitViews = @()
        foreach ($viewAction in @('procedure','module')) {
            $splitLayout = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
            if ($splitLayout.Panes.Count -ne 2) { throw 'Expected two split panes.' }
            $targetIndex = if ($viewAction -eq 'procedure') { 0 } else { 1 }
            $targetPane = $splitLayout.Panes[$targetIndex]
            $splitView = Invoke-Session @{ Command = 'set_code_view'; Project = $project.Name; Module = $module.Name; Pane = $targetPane.Pane; ExpectedWindowVersion = $targetPane.WindowVersion; ExpectedSha256 = $targetPane.State.Sha256; ExpectedMode = 2; StartLine = 2; Action = $viewAction }
            $splitViews += $splitView
            $expectedView = if ($viewAction -eq 'procedure') { 0 } else { 1 }
            if (-not $splitView.Verified -or $splitView.AfterPanes.Count -ne 2 -or @($splitView.AfterPanes | Where-Object { $_.View -ne $expectedView -or $_.Sha256 -ne $splitCode.Sha256 }).Count) { throw 'Window-wide split view change was not verified.' }
        }
        $splitViews | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $directory 'installed-split-code-view-cycle.json') -Encoding UTF8
        $unsplit = Invoke-Session @{ Command = 'set_code_split'; Project = $project.Name; Module = $module.Name; ExpectedSha256 = $splitCode.Sha256; ExpectedMode = 2; StartLine = 2; Action = 'unsplit'; ControlCaption = $splitControl.Caption }
        Start-Sleep -Milliseconds 200
        Invoke-Session @{ Command = 'code_panes' } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'installed-unsplit-panes.json') -Encoding UTF8
        $unsplitObserved = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
        if ($unsplitObserved.Panes.Count -ne 1) { throw 'Split fixture did not return to a single pane.' }
        $breakModule = $project.VBComponents.Add(1); $breakModule.Name = 'BreakViewProbe'
        $breakModule.CodeModule.AddFromString("Public Sub PauseForView()`r`n    ThisWorkbook.Worksheets(1).Range(""A5"").Value2 = 10`r`n    Stop`r`n    ThisWorkbook.Worksheets(1).Range(""A5"").Value2 = ThisWorkbook.Worksheets(1).Range(""A5"").Value2 + 1`r`nEnd Sub")
        $breakCode = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $breakModule.Name }
        $breakStarted = Invoke-Session @{ Command = 'run_sub'; Project = $project.Name; Module = $breakModule.Name; Procedure = 'PauseForView'; ExpectedSha256 = $breakCode.Sha256; ExpectedMode = 2 }
        Start-Sleep -Milliseconds 200
        $breakState = Invoke-Session @{ Command = 'debug_state'; Project = $project.Name }
        if ($breakState.Mode -ne 1 -or $book.Worksheets.Item(1).Range('A5').Value2 -ne 10) { throw 'The disposable procedure did not stop at its marker.' }
        $breakViewResults = @()
        foreach ($topology in @('single','split')) {
            if ($topology -eq 'split') {
                $breakSplitControl = $vbe.CommandBars.FindControl(1,302)
                Invoke-Session @{ Command = 'set_code_split'; Project = $project.Name; Module = $breakModule.Name; ExpectedSha256 = $breakCode.Sha256; ExpectedMode = 1; StartLine = 3; Action = 'split'; ControlCaption = $breakSplitControl.Caption } | Out-Null
                Start-Sleep -Milliseconds 200
            }
            foreach ($viewAction in @('procedure','module')) {
                $breakLayout = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $breakModule.Name }
                $expectedCount = if ($topology -eq 'split') { 2 } else { 1 }
                if ($breakLayout.Panes.Count -ne $expectedCount) { throw 'Unexpected break-mode pane count.' }
                $breakPane = $breakLayout.Panes[$expectedCount-1]
                $breakView = Invoke-Session @{ Command = 'set_code_view'; Project = $project.Name; Module = $breakModule.Name; Pane = $breakPane.Pane; ExpectedWindowVersion = $breakPane.WindowVersion; ExpectedSha256 = $breakPane.State.Sha256; ExpectedMode = 1; StartLine = 3; Action = $viewAction }
                $afterBreakView = Invoke-Session @{ Command = 'debug_state'; Project = $project.Name }
                $breakMarker = $book.Worksheets.Item(1).Range('A5').Value2
                if (-not $breakView.Verified -or $afterBreakView.Mode -ne 1 -or $breakMarker -ne 10) { throw 'View change resumed execution or failed in break mode.' }
                $breakViewResults += [pscustomobject]@{ Topology = $topology; Action = $viewAction; View = $breakView; DebugState = $afterBreakView; Marker = $breakMarker }
            }
        }
        $resumeControl = $vbe.CommandBars.FindControl(1,186)
        $resumed = Invoke-Session @{ Command = 'invoke_debug'; Project = $project.Name; Module = $breakModule.Name; ExpectedSha256 = $breakCode.Sha256; ExpectedMode = 1; StartLine = 3; Action = 'continue'; ControlId = 186; ControlCaption = $resumeControl.Caption }
        Start-Sleep -Milliseconds 200
        $afterResume = Invoke-Session @{ Command = 'debug_state'; Project = $project.Name }
        $resumeMarker = $book.Worksheets.Item(1).Range('A5').Value2
        $afterResumeCode = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $breakModule.Name }
        if ($afterResume.Mode -ne 2 -or $resumeMarker -ne 11 -or $afterResumeCode.Sha256 -ne $breakCode.Sha256) { throw 'Execution did not resume exactly once after the view cycles.' }
        [pscustomobject]@{ Start = $breakStarted; InitialStop = $breakState; Cases = $breakViewResults; Resume = $resumed; AfterResume = $afterResume; Marker = $resumeMarker; CodeUnchanged = $true } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'installed-break-code-view-cycle.json') -Encoding UTF8
        $runForm = $project.VBComponents.Add(3); $runForm.Name = 'NativeRunFormProbe'
        $runForm.CodeModule.AddFromString("Private Sub UserForm_Activate()`r`n    ThisWorkbook.Worksheets(1).Range(""A1"").Value2 = ""VBAi form executed""`r`n    Unload Me`r`nEnd Sub")
        $runForm.DesignerWindow().Visible = $true
        $runForm.DesignerWindow().SetFocus()
        $runCommand = $vbe.CommandBars.FindControl(1,186)
        $formCode = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $runForm.Name }
        $formTree = Invoke-Session @{ Command = 'form_tree'; Project = $project.Name; Form = $runForm.Name }
        $formGuards = @()
        foreach ($guard in @('ExpectedSha256','ExpectedTreeVersion')) {
            $guardRequest = @{ Command = 'run_form'; Project = $project.Name; Form = $runForm.Name; ExpectedMode = 2; ExpectedSha256 = $formCode.Sha256; ExpectedTreeVersion = $formTree.TreeVersion; ControlCaption = $runCommand.Caption }
            $guardRequest[$guard] = 'stale'
            $rejectedGuard = $false
            try { Invoke-Session $guardRequest | Out-Null }
            catch { if ($_.Exception.Message -notmatch 'changed before execution') { throw }; $rejectedGuard = $true }
            if (-not $rejectedGuard -or $null -ne $book.Worksheets.Item(1).Range('A1').Value2) { throw 'A stale UserForm revision was not safely refused.' }
            $formGuards += $guard
        }
        $formStarted = Invoke-Session @{ Command = 'run_form'; Project = $project.Name; Form = $runForm.Name; ExpectedMode = 2; ExpectedSha256 = $formCode.Sha256; ExpectedTreeVersion = $formTree.TreeVersion; ControlCaption = $runCommand.Caption }
        Start-Sleep -Milliseconds 500
        $formStatus = Invoke-Session @{ Command = 'form_run_status'; Project = $project.Name; Query = $formStarted.OperationId }
        $formValue = $book.Worksheets.Item(1).Range('A1').Value2
        [pscustomobject]@{ Operation = $formStatus; RuntimeValue = $formValue; RejectedStaleGuards = $formGuards } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'native-run-form.json') -Encoding UTF8
        if ($formValue -ne 'VBAi form executed' -or $formStatus.State -ne 'CommandReturned') { throw 'Native UserForm execution was not verified.' }
        $modalForm = $project.VBComponents.Add(3); $modalForm.Name = 'ModalRunFormProbe'
        $modalForm.Properties.Item('Caption').Value = 'VBAi isolated modal test'
        $modalForm.CodeModule.AddFromString("Private Sub UserForm_Activate()`r`n    ThisWorkbook.Worksheets(1).Range(""A2"").Value2 = ""Modal activated""`r`nEnd Sub`r`nPrivate Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)`r`n    ThisWorkbook.Worksheets(1).Range(""A3"").Value2 = CloseMode`r`nEnd Sub")
        $modalForm.DesignerWindow().Visible = $true; $modalForm.DesignerWindow().SetFocus()
        $modalCode = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $modalForm.Name }
        $modalTree = Invoke-Session @{ Command = 'form_tree'; Project = $project.Name; Form = $modalForm.Name }
        [pscustomobject]@{ Name = $modalForm.Name; Caption = $modalForm.Properties.Item('Caption').Value; DesignerCaption = $modalForm.Designer.Caption; WindowCaption = $modalForm.DesignerWindow().Caption; ActiveWindow = $vbe.ActiveWindow.Caption } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'native-run-form-modal-before.json') -Encoding UTF8
        $modalStarted = Invoke-Session @{ Command = 'run_form'; Project = $project.Name; Form = $modalForm.Name; ExpectedMode = 2; ExpectedSha256 = $modalCode.Sha256; ExpectedTreeVersion = $modalTree.TreeVersion; ControlCaption = $runCommand.Caption }
        Start-Sleep -Milliseconds 500
        $modalWindows = [NavigationProbeDialog]::WindowStates($probeProcess.Id)
        $modalWindows | Set-Content -LiteralPath (Join-Path $directory 'native-run-form-modal-windows.txt') -Encoding UTF8
        if ([NavigationProbeDialog]::FindOwnedForm($probeProcess.Id, 'VBAi isolated modal test') -eq [IntPtr]::Zero) { throw 'The modal UserForm was not visible.' }
        $modalWhileOpen = Invoke-Session @{ Command = 'form_run_status'; Project = $project.Name; Query = $modalStarted.OperationId }
        if ($modalWhileOpen.State -eq 'Failed' -or $modalWhileOpen.RuntimeVerified) { throw 'Modal form status was incorrect.' }
        if (-not [NavigationProbeDialog]::CloseOwnedForm($probeProcess.Id, 'VBAi isolated modal test')) { throw 'The owned modal form close was not sent.' }
        for ($attempt = 0; $attempt -lt 20 -and [NavigationProbeDialog]::FindOwnedForm($probeProcess.Id, 'VBAi isolated modal test') -ne [IntPtr]::Zero; $attempt++) { Start-Sleep -Milliseconds 100 }
        if ([NavigationProbeDialog]::FindOwnedForm($probeProcess.Id, 'VBAi isolated modal test') -ne [IntPtr]::Zero) { throw 'The modal test form did not close.' }
        $modalAfter = Invoke-Session @{ Command = 'form_run_status'; Project = $project.Name; Query = $modalStarted.OperationId }
        $activated = $book.Worksheets.Item(1).Range('A2').Value2
        $closeMode = $book.Worksheets.Item(1).Range('A3').Value2
        if ($activated -ne 'Modal activated' -or $null -eq $closeMode -or $closeMode -ne 0 -or $modalAfter.State -ne 'CommandReturned') { throw 'Modal Activate/QueryClose runtime proof failed.' }
        [pscustomobject]@{ Windows = $modalWindows; WhileOpen = $modalWhileOpen; AfterClose = $modalAfter; Activation = $activated; QueryCloseMode = $closeMode } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'native-run-form-modal.json') -Encoding UTF8
        $errorForm = $project.VBComponents.Add(3); $errorForm.Name = 'ErrorRunFormProbe'
        $errorForm.Properties.Item('Caption').Value = 'VBAi isolated error test'
        $errorForm.CodeModule.AddFromString("Private Sub UserForm_Activate()`r`n    ThisWorkbook.Worksheets(1).Range(""A4"").Value2 = ""Error event entered""`r`n    Err.Raise 5`r`nEnd Sub")
        $errorForm.DesignerWindow().Visible = $true; $errorForm.DesignerWindow().SetFocus()
        $errorCode = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $errorForm.Name }
        $errorTree = Invoke-Session @{ Command = 'form_tree'; Project = $project.Name; Form = $errorForm.Name }
        $errorStarted = Invoke-Session @{ Command = 'run_form'; Project = $project.Name; Form = $errorForm.Name; ExpectedMode = 2; ExpectedSha256 = $errorCode.Sha256; ExpectedTreeVersion = $errorTree.TreeVersion; ControlCaption = $runCommand.Caption }
        Start-Sleep -Milliseconds 400
        $eventDiagnostic = Invoke-Session @{ Command = 'debug_dialog' }
        $eventDiagnostic | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'native-run-form-error-dialog.json') -Encoding UTF8
        $eventOperation = Invoke-Session @{ Command = 'form_run_status'; Project = $project.Name; Query = $errorStarted.OperationId }
        $eventDebugState = Invoke-Session @{ Command = 'debug_state'; Project = $project.Name }
        [pscustomobject]@{ Operation = $eventOperation; State = $eventDebugState; Windows = [NavigationProbeDialog]::WindowStates($probeProcess.Id) } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'native-run-form-error-state.json') -Encoding UTF8
        if (-not $eventDiagnostic.Available -or $eventDiagnostic.Diagnostic -notmatch "Erreur d'exécution '5'") { throw 'UserForm event runtime diagnostic was not observed.' }
        $eventWhileOpen = Invoke-Session @{ Command = 'form_run_status'; Project = $project.Name; Query = $errorStarted.OperationId }
        if ($eventWhileOpen.RuntimeVerified) { throw 'An event error was incorrectly reported as runtime success.' }
        $endButtons = @($eventDiagnostic.Buttons | Where-Object { ($_ -replace '&','') -in @('Fin','End') })
        if ($endButtons.Count -ne 1) { throw 'The exact native End button was not identified.' }
        $eventDismiss = Invoke-Session @{ Command = 'respond_debug_dialog'; Diagnostic = $eventDiagnostic.Diagnostic; Button = $endButtons[0] }
        if ($eventDismiss.Verification -ne 'DialogClosed') { throw 'The runtime error dialog was not closed.' }
        Start-Sleep -Milliseconds 200
        $eventAfter = Invoke-Session @{ Command = 'form_run_status'; Project = $project.Name; Query = $errorStarted.OperationId }
        $eventMarker = $book.Worksheets.Item(1).Range('A4').Value2
        $eventState = Invoke-Session @{ Command = 'debug_state'; Project = $project.Name }
        if ($eventMarker -ne 'Error event entered' -or $eventState.Mode -ne 2) { throw 'Event error execution and return to design mode were not verified.' }
        [pscustomobject]@{ Diagnostic = $eventDiagnostic; WhileOpen = $eventWhileOpen; Dismiss = $eventDismiss; After = $eventAfter; Marker = $eventMarker; ModeAfter = $eventState.Mode } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'native-run-form-error.json') -Encoding UTF8
        $module.CodeModule.CodePane.Show(); $module.CodeModule.CodePane.Window.SetFocus()
    }
    $missingProof = $null
    if ($UseBridge) {
        $missingModule = $project.VBComponents.Add(1); $missingModule.Name = 'MissingDefinitionProbe'
        $missingModule.CodeModule.AddFromString("Public Sub MissingCaller()`r`n    MissingTarget`r`nEnd Sub")
        $missingCode = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $missingModule.Name }
        $missing = Invoke-Session @{ Command = 'native_code_navigation'; Action = 'definition'; Project = $project.Name; Module = $missingModule.Name; ExpectedSha256 = $missingCode.Sha256; ExpectedMode = 2; StartLine = 2; StartColumn = 5; EndColumn = 18; Expression = 'MissingTarget'; ControlCaption = $command.Caption }
        Start-Sleep -Milliseconds 200
        $dialog = Invoke-Session @{ Command = 'debug_dialog' }
        $dialog | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'missing-definition-dialog.json') -Encoding UTF8
        if (-not $dialog.Available -or $dialog.Diagnostic -ne "L'identificateur sous le curseur n'est pas reconnu") { throw 'Expected missing-definition diagnostic was not observed.' }
        $whileOpen = Invoke-Session @{ Command = 'native_code_navigation'; Action = 'status'; Query = $missing.OperationId }
        $dismiss = Invoke-Session @{ Command = 'respond_debug_dialog'; Diagnostic = $dialog.Diagnostic; Button = 'OK' }
        if ($dismiss.Verification -ne 'DialogClosed') { throw 'The installed bridge did not close the native diagnostic.' }
        Start-Sleep -Milliseconds 2100
        $missingState = Invoke-Session @{ Command = 'native_code_navigation'; Action = 'status'; Query = $missing.OperationId }
        if ($missingState.State -ne 'Completed' -or $missingState.NavigationObserved -or $missingState.DefinitionResolved) { throw 'Missing definition was incorrectly reported as navigation or left pending.' }
        $missingProof = [pscustomobject]@{ Dialog = $dialog; WhileOpen = $whileOpen; Dismiss = $dismiss; AfterDismiss = $missingState }
    }
    $externalProof = $null
    if ($UseBridge) {
        $externalModule = $project.VBComponents.Add(1); $externalModule.Name = 'ExternalDefinitionProbe'
        $externalModule.CodeModule.AddFromString("Public Sub ExternalCaller()`r`n    Dim sample As Range`r`nEnd Sub")
        $externalModule.CodeModule.CodePane.Show()
        $externalModule.CodeModule.CodePane.Window.SetFocus()
        Start-Sleep -Milliseconds 250
        $externalCode = Invoke-Session @{ Command = 'read_module'; Project = $project.Name; Module = $externalModule.Name }
        $external = Invoke-Session @{ Command = 'native_code_navigation'; Action = 'definition'; Project = $project.Name; Module = $externalModule.Name; ExpectedSha256 = $externalCode.Sha256; ExpectedMode = 2; StartLine = 2; StartColumn = 19; EndColumn = 24; Expression = 'Range'; ControlCaption = $command.Caption }
        Start-Sleep -Milliseconds 400
        $externalState = Invoke-Session @{ Command = 'native_code_navigation'; Action = 'status'; Query = $external.OperationId }
        $externalWindows = Invoke-Session @{ Command = 'vbe_windows' }
        if ($externalState.State -ne 'Completed' -or -not $externalState.NavigationObserved -or @($externalWindows.Windows | Where-Object { $_.Properties.Type -eq 2 -and $_.Properties.Visible }).Count -ne 1) { throw ('External definition did not reach Object Browser: ' + ($externalState | ConvertTo-Json -Depth 6 -Compress)) }
        $externalDialog = Invoke-Session @{ Command = 'debug_dialog' }
        $externalDialog | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'external-definition-dialog.json') -Encoding UTF8
        if ($externalDialog.Available) {
            if ($externalDialog.Diagnostic -ne "Impossible d'aller à 'Range' qui est caché") { throw 'Unexpected external-definition diagnostic.' }
            $modalRejected = $false
            try { Invoke-Session @{ Command = 'select_object_browser'; ObjectName = 'Range' } | Out-Null }
            catch { if ($_.Exception.Message -notmatch 'disabled by a modal window') { throw }; $modalRejected = $true }
            if (-not $modalRejected) { throw 'Selection behind a modal window was not refused.' }
            $externalDismiss = Invoke-Session @{ Command = 'respond_debug_dialog'; Diagnostic = $externalDialog.Diagnostic; Button = 'OK' }
            if ($externalDismiss.Verification -ne 'DialogClosed') { throw 'Hidden definition diagnostic was not closed.' }
            Start-Sleep -Milliseconds 200
        }
        $browserRead = Invoke-Session @{ Command = 'read_object_browser' }
        $browserRead | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'object-browser-read.json') -Encoding UTF8
        if (-not $browserRead.Available -or @($browserRead.Descriptions | Where-Object { $_.Text }).Count -eq 0 -or @($browserRead.Selections | Where-Object { $_.SelectionReadable -and $_.Selected.Count -gt 0 }).Count -eq 0) { throw 'Native Object Browser selections or description were unreadable.' }
        $selectedBrowser = Invoke-Session @{ Command = 'select_object_browser'; ObjectName = 'Range'; Procedure = 'Address' }
        $selectedBrowser | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'object-browser-selection.json') -Encoding UTF8
        if (-not $selectedBrowser.SelectionObserved -or @($selectedBrowser.Snapshot.Descriptions | Where-Object { $_.Text -match 'Address' }).Count -eq 0) { throw 'Native member selection or description readback failed.' }
        $libraries = Invoke-Session @{ Command = 'list_object_browser'; Pane = 'libraries'; Query = 'Excel'; Limit = 2 }
        if (-not $libraries.Available -or $libraries.TotalMatches -ne 1 -or $libraries.Items[0].Label -ne 'Excel') { throw 'Excel library listing failed.' }
        $page1 = Invoke-Session @{ Command = 'list_object_browser'; Pane = 'members'; Limit = 2 }
        $page2 = Invoke-Session @{ Command = 'list_object_browser'; Pane = 'members'; Limit = 2; Offset = 2 }
        if (-not $page1.Available -or $page1.Items.Count -ne 2 -or -not $page1.HasMore -or $page2.Items.Count -ne 2 -or $page1.Items[0].Label -eq $page2.Items[0].Label) { throw 'Native member pagination failed.' }
        [NavigationProbeDialog]::WindowStates($probeProcess.Id) | Set-Content -LiteralPath (Join-Path $directory 'object-browser-top-windows.txt') -Encoding UTF8
        $libraryChange = Invoke-Session @{ Command = 'select_object_browser'; ObjectName = 'Range'; Procedure = 'Address'; Context = 'Excel' }
        if (-not $libraryChange.SelectionObserved -or @($libraryChange.Snapshot.Selections | Where-Object { $_.Type -eq 'ControlType.ComboBox' -and $_.Selected -contains 'Excel' }).Count -ne 1) { throw 'Excel library selection was not verified.' }
        [pscustomobject]@{ Libraries = $libraries; Page1 = $page1; Page2 = $page2; LibraryChange = $libraryChange } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'object-browser-lists.json') -Encoding UTF8
        $vbaLibrary = Invoke-Session @{ Command = 'select_object_browser'; Context = 'VBA' }
        if (-not $vbaLibrary.SelectionObserved) { throw 'VBA library-only selection failed.' }
        $classes = Invoke-Session @{ Command = 'list_object_browser'; Pane = 'classes'; Query = 'collection'; Limit = 2 }
        [pscustomobject]@{ Library = $vbaLibrary; Classes = $classes } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'object-browser-class-search.json') -Encoding UTF8
        if (-not $classes.Available -or $classes.TotalMatches -lt 1 -or @($classes.Items | Where-Object { $_.Label -eq 'Classes Collection' }).Count -ne 1 -or @($classes.Items | Where-Object { $_.Label -notmatch '(?i)collection' }).Count -ne 0) { throw 'VBA class substring search failed.' }
        $classPage1 = Invoke-Session @{ Command = 'list_object_browser'; Pane = 'classes'; Limit = 2 }
        $classPage2 = Invoke-Session @{ Command = 'list_object_browser'; Pane = 'classes'; Limit = 2; Offset = 2 }
        if ($classPage1.Items.Count -ne 2 -or $classPage2.Items.Count -ne 2 -or -not $classPage1.HasMore -or $classPage1.Items[0].Label -eq $classPage2.Items[0].Label) { throw 'Class pagination failed.' }
        $missingLibraryRejected = $false
        try { Invoke-Session @{ Command = 'select_object_browser'; Context = '__VbaiMissingLibrary__' } | Out-Null }
        catch { if ($_.Exception.Message -notmatch 'Expected one exact library; found 0') { throw }; $missingLibraryRejected = $true }
        if (-not $missingLibraryRejected) { throw 'Missing library was not rejected.' }
        $afterMissingLibrary = Invoke-Session @{ Command = 'read_object_browser' }
        if (@($afterMissingLibrary.Selections | Where-Object { $_.Type -eq 'ControlType.ComboBox' -and $_.Selected -contains 'VBA' }).Count -ne 1) { throw 'Missing library changed the library selection.' }
        $returnExcel = Invoke-Session @{ Command = 'select_object_browser'; Context = 'Excel'; ObjectName = 'Range'; Procedure = 'Address' }
        if (-not $returnExcel.SelectionObserved) { throw 'Return to Excel.Range.Address failed.' }
        [pscustomobject]@{ VbaLibrary = $vbaLibrary; Classes = $classes; Page1 = $classPage1; Page2 = $classPage2; MissingLibraryRejected = $missingLibraryRejected; ReturnExcel = $returnExcel } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'object-browser-library-cycle.json') -Encoding UTF8
        $missingClassRejected = $false
        try { Invoke-Session @{ Command = 'select_object_browser'; ObjectName = '__VbaiMissingClass__' } | Out-Null }
        catch { if ($_.Exception.Message -notmatch 'found 0') { throw }; $missingClassRejected = $true }
        if (-not $missingClassRejected) { throw 'Missing browser class was not rejected.' }
        $afterMissingClass = Invoke-Session @{ Command = 'read_object_browser' }
        if (@($afterMissingClass.Descriptions | Where-Object { $_.Text -match 'Membre de Excel.Range' -and $_.Text -match 'Property Address' }).Count -ne 1) { throw 'Missing class changed the existing member description.' }
        [pscustomobject]@{ MissingClassRejected = $missingClassRejected; ExistingDescriptionPreserved = $true } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'object-browser-guards.json') -Encoding UTF8
        $externalProof = [pscustomobject]@{ Operation = $externalState; ActiveWindow = $externalWindows.ActiveWindow; BrowserWindows = @($externalWindows.Windows | Where-Object { $_.Properties.Type -eq 2 }); SemanticMemberVerified = $false }
    }
    [pscustomobject]@{ Definition = $definition; LastPosition = $lastPosition; ExternalDefinition = $externalProof; MissingDefinition = $missingProof; ChangedSourceOperation = $rejected; ConcurrentRejected = $duplicateRejected; DefinitionOperation = $completed; LastPositionOperation = $completedLast; BridgeIdentity = $status; Scheduler = $(if ($UseBridge) { "Installed add-in UI context and named pipe" } else { "Queued STA test context" }) } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory $(if ($UseBridge) { 'installed-native-definition-probe.json' } else { 'native-definition-probe.json' })) -Encoding UTF8
    Get-Content -LiteralPath (Join-Path $directory $(if ($UseBridge) { 'installed-native-definition-probe.json' } else { 'native-definition-probe.json' }))
} finally {
    [Threading.SynchronizationContext]::SetSynchronizationContext($priorContext)
    # Restore trust before Quit, which can block in some COM teardown paths.
    if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
    else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
    if ($null -ne $probeProcess) {
        [NavigationProbeDialog]::CloseOwnedDiagnostic($probeProcess.Id)
        [NavigationProbeDialog]::CloseOwnedForm($probeProcess.Id, 'VBAi isolated modal test') | Out-Null
        [NavigationProbeDialog]::CloseOwnedForm($probeProcess.Id, 'VBAi isolated error test') | Out-Null
    }
    if ($null -ne $otherBook) { $otherBook.Close($false) }
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) {
        # This process was created after the no-Excel precondition and contains only
        # the disposable workbook. Quit can leave it alive because of COM references.
        Stop-Process -Id $probeProcess.Id -Force
    }
}

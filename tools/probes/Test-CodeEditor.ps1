<#
.SYNOPSIS
Runs one explicit native code/editor probe with its original guard and oracle.
.DESCRIPTION
ComponentRoundTrip requires an explicitly selected owned disposable host and
retains failed exports for recovery. Other scenarios own an isolated assembly
fixture; NativeDefinition optionally uses the registered bridge. PaneView alone
accepts -Split. Native execution is never retried by this command.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Clipboard','PaneScroll','PaneView','ComponentRoundTrip','CommandInventory','NativeHistory','NativeDefinition')]
    [string]$Scenario,
    [ValidateNotNullOrEmpty()][string]$AssemblyPath,
    [ValidateNotNullOrEmpty()][string]$OutputDirectory,
    [switch]$AllowTemporaryVbaAccess,
    [switch]$UseBridge,
    [switch]$Split,
    [ValidateRange(1,2147483647)][int]$HostProcessId,
    [switch]$OwnedDisposableHost,
    [ValidateSet('Class','Form')][string]$Kind,
    [ValidateNotNullOrEmpty()][string]$Project = 'VBAProject',
    [ValidateNotNullOrEmpty()][string]$Module
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'VbeProbe.Common.ps1')
$VbeProbeJsonDepth = 8
$VbeProbeRawError = $false
$VbeProbeTraceRequests = $Scenario -in @('Clipboard','NativeHistory')
if ($Scenario -eq 'ComponentRoundTrip') {
    $required = @('HostProcessId','OwnedDisposableHost','Kind')
    $allowed = @('HostProcessId','OwnedDisposableHost','Kind','Project','Module')
} else {
    $required = @('AssemblyPath','OutputDirectory')
    $allowed = @('AssemblyPath','OutputDirectory','AllowTemporaryVbaAccess')
    if ($Scenario -eq 'NativeDefinition') { $allowed += 'UseBridge' }
    if ($Scenario -eq 'PaneView') { $allowed += 'Split' }
}
Assert-VbeProbeParameters -Bound $PSBoundParameters -Required $required -Allowed $allowed
if ($Scenario -eq 'ComponentRoundTrip') {
    $ownedHost = Get-Process -Id $HostProcessId -ErrorAction Stop
    if ($ownedHost.ProcessName -ne 'EXCEL') { throw 'An owned disposable Excel process is required.' }
    if (-not $PSBoundParameters.ContainsKey('Module')) {
        $Module = if ($Kind -eq 'Class') { 'CodexClassRoundTrip' } else { 'CodexFormRoundTrip' }
    }
}

switch ($Scenario) {
    'Clipboard' {
        $ErrorActionPreference = 'Stop'
        . $VbeProbeInitializeAssembly

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
            . $VbeProbeOpenAssemblyExcel
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
            $project = $book.VBProject
            $module = $project.VBComponents.Add(1); $module.Name = 'ClipboardProbe'
            $module.CodeModule.AddFromString("' alpha beta`r`n' second")
            $before = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
            $settings = [Activator]::CreateInstance($assembly.GetType('VBAi.LlmSettings'), $true)
            $settings.VbeEditApproval = 'Automatic'
            $llm = [Activator]::CreateInstance($assembly.GetType('VBAi.LlmVbeTools'), [object[]]@($session,$null,$settings))
            $llm.BoundProject = $project.Name
            Add-Type -TypeDefinition 'public static class ClipboardDiffCounter { public static int Count; public static void Changed(object change) { Count++; } }'
            $editedEvent = $llm.GetType().GetEvent('CodeEdited')
            $editedHandler = [Delegate]::CreateDelegate($editedEvent.EventHandlerType, [ClipboardDiffCounter].GetMethod('Changed'))
            $editedEvent.AddEventHandler($llm, $editedHandler)
            function Invoke-ClipboardTool([string]$Name, [hashtable]$Fields) {
                $response = $llm.Invoke($Name, ($Fields | ConvertTo-Json -Compress)) | ConvertFrom-Json
                if (-not $response.Ok) { throw $response.Error }
                return $response.Data
            }
            $range = @{ Project=$project.Name; Module=$module.Name; ExpectedSha256=$before.Sha256; StartLine=1; StartColumn=3; EndLine=1; EndColumn=8 }
            $clipboardTouched = $true
            $copy = Invoke-ClipboardTool 'copy_code' $range
            $readClipboard = Invoke-Session @{ Command='read_code_clipboard' }
            if ($readClipboard.Text -cne 'alpha') { throw 'Native copy did not match the exact selection.' }
            $cut = Invoke-ClipboardTool 'cut_code' $range
            $afterCut = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
            if ($afterCut.Code -cne "'  beta`r`n' second") { throw 'Cut changed unexpected text.' }
            $paste = Invoke-ClipboardTool 'paste_code' @{ Project=$project.Name; Module=$module.Name; ExpectedSha256=$afterCut.Sha256; ExpectedClipboardVersion=$cut.ClipboardVersion; StartLine=2; StartColumn=3; EndLine=2; EndColumn=9 }
            $afterPaste = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
            if ($afterPaste.Code -cne "'  beta`r`n' alpha") { throw 'Paste changed unexpected text.' }
            Invoke-ClipboardTool 'undo_code_edit' @{ Project=$project.Name; Module=$module.Name; ExpectedSha256=$afterPaste.Sha256 } | Out-Null
            $undone = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
            if ($undone.Sha256 -ne $afterCut.Sha256) { throw 'Clipboard edit undo did not restore source.' }
            Invoke-ClipboardTool 'redo_code_edit' @{ Project=$project.Name; Module=$module.Name; ExpectedSha256=$undone.Sha256 } | Out-Null
            $redone = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
            if ($redone.Sha256 -ne $afterPaste.Sha256) { throw 'Clipboard edit redo did not restore source.' }
            [System.Windows.Forms.Clipboard]::SetText("' café`n' suite", [System.Windows.Forms.TextDataFormat]::UnicodeText)
            $staleRejected = $false
            try { Invoke-ClipboardTool 'paste_code' @{ Project=$project.Name; Module=$module.Name; ExpectedSha256=$redone.Sha256; ExpectedClipboardVersion=$cut.ClipboardVersion; StartLine=1; StartColumn=1; EndLine=1; EndColumn=1 } | Out-Null }
            catch { $staleRejected = $_.Exception.Message -match 'Clipboard changed' }
            $emptyModule = $project.VBComponents.Add(1); $emptyModule.Name = 'EmptyClipboardTarget'
            $empty = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$emptyModule.Name }
            $unicode = Invoke-Session @{ Command='read_code_clipboard' }
            $insert = Invoke-ClipboardTool 'paste_code' @{ Project=$project.Name; Module=$emptyModule.Name; ExpectedSha256=$empty.Sha256; ExpectedClipboardVersion=$unicode.Version; StartLine=1; StartColumn=1; EndLine=1; EndColumn=1 }
            $inserted = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$emptyModule.Name }
            if ($inserted.Code -cne "' café`r`n' suite" -or -not $insert.Verified -or -not $staleRejected -or [ClipboardDiffCounter]::Count -ne 5) { throw 'Clipboard native guard, Unicode or chat diff verification failed.' }
            $proof = [pscustomobject]@{ CopyVerified=$true; CutVerified=$cut.Verified; PasteVerified=$paste.Verified; UndoRedoVerified=$true; StaleClipboardRejected=$staleRejected; UnicodeEmptyModuleVerified=$true; ChatDiffCount=[ClipboardDiffCounter]::Count; Mvid=$assembly.ManifestModule.ModuleVersionId.ToString(); ClipboardRestored=$false }
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
        $proof | ConvertTo-Json | Set-Content (Join-Path $directory 'native-code-clipboard.json') -Encoding UTF8
        $proof | ConvertTo-Json

    }
    'PaneScroll' {
        $ErrorActionPreference = 'Stop'
        . $VbeProbeInitializeAssembly

        try {
            . $VbeProbeOpenAssemblyExcel
            $vbe.MainWindow.Visible = $true
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
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
            $request = New-Object VBAi.Request
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

    }
    'PaneView' {
        $ErrorActionPreference = 'Stop'
        . $VbeProbeInitializeAssembly

        try {
            . $VbeProbeOpenAssemblyExcel
            $vbe.MainWindow.Visible = $true
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
            $project = $book.VBProject
            $module = $project.VBComponents.Add(1); $module.Name = 'EditorCommandProbe'
            $module.CodeModule.AddFromString("Public Sub First()`r`n    Debug.Print 1`r`nEnd Sub`r`nPublic Sub Second()`r`n    Debug.Print 2`r`nEnd Sub")
            $pane = $module.CodeModule.CodePane; $pane.Show(); $pane.SetSelection(2, 1, 2, 1)
            $module.CodeModule.AddFromString(((1..150 | ForEach-Object { "' Scroll probe line $_" }) -join "`r`n"))
            if ($Split) {
                $layout = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
                $splitControl = $vbe.CommandBars.FindControl(1,302)
                $splitResult = Invoke-Session @{ Command = 'set_code_split'; Project = $project.Name; Module = $module.Name; ExpectedSha256 = $layout.Panes[0].State.Sha256; ExpectedMode = 2; StartLine = 2; Action = 'split'; ControlCaption = $splitControl.Caption }
                if (-not $splitResult.Verified) { throw 'Native code window did not split.' }
            }
            & (Join-Path $PSScriptRoot 'Inspect-VbeNativeWindows.ps1') -View Tree -HostProcessId $probeProcess.Id | Set-Content -LiteralPath (Join-Path $directory 'code-view-window-tree.json') -Encoding UTF8
            & (Join-Path $PSScriptRoot 'Inspect-VbeAccessibility.ps1') -Target DebugTree -HostProcessId $probeProcess.Id -NamePattern '.*' | Set-Content -LiteralPath (Join-Path $directory 'code-view-accessibility.json') -Encoding UTF8
            $nativeTree = Get-Content -LiteralPath (Join-Path $directory 'code-view-window-tree.json') -Raw | ConvertFrom-Json
            $codeWindow = @($nativeTree | Where-Object { $_.Class -eq 'VbaWindow' -and $_.Title -eq 'EditorCommandProbe (Code)' -and $_.Visible })
            if ($codeWindow.Count -ne 1) { throw 'Exact code window not found for capture.' }
            Add-Type -AssemblyName System.Drawing
            Add-Type -AssemblyName Accessibility
            Add-Type -ReferencedAssemblies System.Drawing,Accessibility -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
public static class CodeViewCapture {
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr hwnd, IntPtr dc, uint flags);
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object result);
    public static string[] Accessibility(IntPtr hwnd) {
        var lines = new System.Collections.Generic.List<string>();
        Guid iid = new Guid("618736e0-3c3d-11cf-810c-00aa00389b71");
        foreach (uint objectId in new uint[] { 0, 0xFFFFFFFC }) {
            object raw;
            int hr = AccessibleObjectFromWindow(hwnd,objectId,ref iid,out raw);
            lines.Add("ObjectId=" + objectId + " HR=" + hr);
            var accessible = raw as Accessibility.IAccessible;
            if (accessible == null) continue;
            Dump(accessible,lines,0);
        }
        return lines.ToArray();
    }
    private static void Dump(Accessibility.IAccessible accessible, System.Collections.Generic.List<string> lines, int depth) {
        if (depth > 8 || lines.Count > 500) return;
        for (int child=0;child<=Math.Min(accessible.accChildCount,100);child++) {
            try { lines.Add(depth + ":" + child + " Name=" + accessible.get_accName(child) + " Role=" + accessible.get_accRole(child) + " Action=" + accessible.get_accDefaultAction(child)); } catch {}
            if (child == 0) continue;
            try { var nested = accessible.get_accChild(child) as Accessibility.IAccessible; if(nested != null) Dump(nested,lines,depth+1); } catch {}
        }
    }
    private delegate bool EnumChild(IntPtr hwnd, IntPtr data);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr hwnd, EnumChild callback, IntPtr data);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr hwnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, System.Text.StringBuilder text, int count);
    [StructLayout(LayoutKind.Sequential)] public struct Point { public int X,Y; }
    [DllImport("user32.dll")] private static extern bool ScreenToClient(IntPtr hwnd, ref Point point);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wparam, IntPtr lparam);
    [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(Point point);
    [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr wparam, IntPtr lparam);
    [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
    public static string ClickView(IntPtr hwnd, bool procedure) {
        var bars = new System.Collections.Generic.List<Rect>();
        EnumChildWindows(hwnd, (child, ignored) => {
            var kind = new System.Text.StringBuilder(128); GetClassName(child,kind,128);
            Rect r;
            if (kind.ToString()=="ScrollBar" && IsWindowVisible(child) && GetWindowRect(child,out r) && r.Right-r.Left > r.Bottom-r.Top) bars.Add(r);
            return true;
        },IntPtr.Zero);
        if(bars.Count != 1 || !IsWindowEnabled(hwnd)) throw new InvalidOperationException("Expected one active unsplit code window.");
        Rect bar = bars[0];
        Point origin = new Point { X=bar.Left, Y=bar.Top }; ScreenToClient(hwnd,ref origin);
        int gap=origin.X, height=bar.Bottom-bar.Top;
        if (gap < height*1.5 || gap > height*3 || height<10 || height>64) throw new InvalidOperationException("Unrecognized view-button geometry: " + gap + "/" + height);
        int x = procedure ? gap/4 : 3*gap/4;
        int y = origin.Y + height/2;
        IntPtr position=new IntPtr((y<<16)|(x&65535));

                Point screen = new Point { X=bar.Left-gap+x, Y=bar.Top+height/2 };
        IntPtr actual=WindowFromPoint(screen);
        var actualClass=new System.Text.StringBuilder(128); GetClassName(actual,actualClass,128);
        if(actualClass.ToString() != "ObtbarWndClass" || !IsChild(hwnd,actual)) throw new InvalidOperationException("Exact native view toolbar was not identified.");
        Point local=screen; ScreenToClient(actual,ref local);
        IntPtr localPosition=new IntPtr((local.Y<<16)|(local.X&65535));
        if(!PostMessage(actual,0x201,new IntPtr(1),localPosition) || !PostMessage(actual,0x202,IntPtr.Zero,localPosition)) throw new InvalidOperationException("Native toolbar click failed.");
        IntPtr hit=SendMessage(hwnd,0x84,IntPtr.Zero,new IntPtr((screen.Y<<16)|(screen.X&65535)));
        return "x="+x+" y="+y+" gap="+gap+" barHeight="+height+" actualClass="+actualClass+" actual="+actual+" target="+hwnd+" hit="+hit;
    }
    public static void Save(IntPtr hwnd, string path) {
        Rect rect; if (!GetWindowRect(hwnd, out rect)) throw new InvalidOperationException("Window bounds unavailable.");
        using (var bitmap = new Bitmap(rect.Right-rect.Left, rect.Bottom-rect.Top))
        using (var graphics = Graphics.FromImage(bitmap)) {
            IntPtr dc = graphics.GetHdc();
            try { if (!PrintWindow(hwnd,dc,0)) throw new InvalidOperationException("PrintWindow failed."); }
            finally { graphics.ReleaseHdc(dc); }
            bitmap.Save(path,ImageFormat.Png);
        }
    }
}
"@
            [CodeViewCapture]::Accessibility([IntPtr]$codeWindow[0].Handle) | Set-Content -LiteralPath (Join-Path $directory 'code-view-msaa.txt') -Encoding UTF8
            [CodeViewCapture]::Save([IntPtr]$codeWindow[0].Handle, (Join-Path $directory 'code-view-window.png'))
            $allPanes = @($vbe.CodePanes | Where-Object { $_.CodeModule.Parent.Name -eq $module.Name })
            $pane = $allPanes[0]
            $pane.Show(); $pane.Window.SetFocus()
            $before = [int]$pane.CodePaneView
            $clickProcedure = [CodeViewCapture]::ClickView([IntPtr]$codeWindow[0].Handle, $true)
            Start-Sleep -Milliseconds 150
            $procedureView = [int]$pane.CodePaneView
            $procedureViews = @($allPanes | ForEach-Object { [int]$_.CodePaneView })
            $clickModule = [CodeViewCapture]::ClickView([IntPtr]$codeWindow[0].Handle, $false)
            Start-Sleep -Milliseconds 150
            $moduleView = [int]$pane.CodePaneView
            $moduleViews = @($allPanes | ForEach-Object { [int]$_.CodePaneView })
            [pscustomobject]@{ Before=$before; PaneCount=$allPanes.Count; ProcedureViews=$procedureViews; ModuleViews=$moduleViews; Procedure=$procedureView; FullModule=$moduleView; ProcedureClick=$clickProcedure; ModuleClick=$clickModule } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'code-view-native-cycle.json') -Encoding UTF8
            if ($procedureView -ne 0 -or $moduleView -ne 1) { throw 'Native view cycle was not verified.' }
            [pscustomobject]@{ CodePaneView = $pane.CodePaneView; Project = $project.Name; Module = $module.Name } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'code-view-before.json') -Encoding UTF8
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

    }
    'ComponentRoundTrip' {
        $ErrorActionPreference = 'Stop'


        $extension = if ($Kind -eq 'Class') { '.cls' } else { '.frm' }
        $path = Join-Path $env:TEMP ("$Module-" + [guid]::NewGuid().ToString('N') + $extension)
        $frxPath = [IO.Path]::ChangeExtension($path, '.frx')
        $completed = $false
        try {
            $before = Invoke-Vbe @{ Command = 'component_properties'; Project = $Project; Module = $Module }
            $codeBefore = Invoke-Vbe @{ Command = 'read_module'; Project = $Project; Module = $Module }
            $treeBefore = if ($Kind -eq 'Form') { Invoke-Vbe @{ Command = 'form_tree'; Project = $Project; Form = $Module } } else { $null }
            $export = Invoke-Vbe @{ Command = 'export_component'; Project = $Project; Module = $Module; Path = $path; ExpectedComponentVersion = $before.Version }
            $projectBeforeRemove = Invoke-Vbe @{ Command = 'project_properties'; Project = $Project }
            $componentBeforeRemove = Invoke-Vbe @{ Command = 'component_properties'; Project = $Project; Module = $Module }
            $afterRemove = Invoke-Vbe @{ Command = 'remove_component'; Project = $Project; Module = $Module; ExpectedProjectVersion = $projectBeforeRemove.Version; ExpectedComponentVersion = $componentBeforeRemove.Version }
            $import = Invoke-Vbe @{ Command = 'import_component'; Project = $Project; Path = $path; ExpectedProjectVersion = $afterRemove.Version }
            if (-not $import.Applied) { throw 'Import response did not confirm that the component was added.' }
            $importedName = $import.ImportedName
            if (-not $importedName) { throw 'Import response did not identify the added component.' }
            $componentAfter = Invoke-Vbe @{ Command = 'component_properties'; Project = $Project; Module = $importedName }
            $codeAfter = Invoke-Vbe @{ Command = 'read_module'; Project = $Project; Module = $importedName }
            $treeAfter = if ($Kind -eq 'Form') { Invoke-Vbe @{ Command = 'form_tree'; Project = $Project; Form = $importedName } } else { $null }
            [pscustomobject]@{
                Kind = $Kind
                ExportBytes = $export.Bytes
                CompanionFrxBytes = if (Test-Path -LiteralPath $frxPath) { (Get-Item -LiteralPath $frxPath).Length } else { 0 }
                Applied = $import.Applied
                Verified = $import.Verified
                VerificationPending = $import.VerificationPending
                ImportError = $import.ImportError
                ComponentReadbackError = $import.ComponentReadbackError
                ProjectReadbackError = $import.ProjectReadbackError
                ImportedName = $importedName
                ImportedType = $componentAfter.Type
                FollowupVerified = ($componentAfter.Type -eq $(if ($Kind -eq 'Form') { 3 } else { 2 }))
                CodeSame = $codeBefore.Code -ceq $codeAfter.Code
                CodeBeforeSha256 = $codeBefore.Sha256
                CodeAfterSha256 = $codeAfter.Sha256
                TreeNodeCountBefore = if ($treeBefore) { $treeBefore.NodeCount } else { $null }
                TreeNodeCountAfter = if ($treeAfter) { $treeAfter.NodeCount } else { $null }
                TreeVersionSame = if ($treeBefore) { $treeBefore.TreeVersion -ceq $treeAfter.TreeVersion } else { $null }
                TreeSame = if ($treeBefore) { ($treeBefore.Nodes | ConvertTo-Json -Depth 6 -Compress) -ceq ($treeAfter.Nodes | ConvertTo-Json -Depth 6 -Compress) } else { $null }
                ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
            } | Format-List
            $completed = $true
        }
        catch {
            Write-Error "Round trip stopped; exported files retained for diagnosis: $path ; $frxPath. $($_.Exception.Message)"
        }
        finally {
            if ($completed) {
                if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
                if (Test-Path -LiteralPath $frxPath) { Remove-Item -LiteralPath $frxPath -Force }
            }
        }

    }
    'CommandInventory' {
        $ErrorActionPreference = 'Stop'
        . $VbeProbeInitializeAssembly

        try {
            . $VbeProbeOpenAssemblyExcel
            $vbe.MainWindow.Visible = $true
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
            $project = $book.VBProject
            $module = $project.VBComponents.Add(1); $module.Name = 'EditorCommandProbe'
            $module.CodeModule.AddFromString("Public Sub First()`r`n    Debug.Print 1`r`nEnd Sub`r`nPublic Sub Second()`r`n    Debug.Print 2`r`nEnd Sub")
            $pane = $module.CodeModule.CodePane; $pane.Show(); $pane.SetSelection(2, 1, 2, 1)
            $commands = @()
            foreach ($query in @('nêtre', 'procédure', 'module')) {
                $commands += @(Invoke-Session @{ Command = 'list_commands'; Query = $query; Limit = 200 })
            }
            $commands | Sort-Object Id,Path -Unique | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'editor-command-inventory.json') -Encoding UTF8
            $commands | Sort-Object Id,Path -Unique | Format-Table Id,Caption,Enabled,Path -AutoSize
            $split = $vbe.CommandBars.FindControl(1, 302)
            $beforeCount = $vbe.CodePanes.Count
            $split.Execute()
            $afterCount = $vbe.CodePanes.Count
            $afterCaption = $split.Caption
            $split.Execute()
            $restoredCount = $vbe.CodePanes.Count
            [pscustomobject]@{ BeforePanes = $beforeCount; AfterSplitPanes = $afterCount; AfterCaption = $afterCaption; RestoredPanes = $restoredCount } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'editor-split-probe.json') -Encoding UTF8
            Get-Content -LiteralPath (Join-Path $directory 'editor-split-probe.json')
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

    }
    'NativeHistory' {
        $ErrorActionPreference = 'Stop'
        . $VbeProbeInitializeAssembly

        try {
            . $VbeProbeOpenAssemblyExcel
            $vbe.MainWindow.Visible = $true
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
            $project = $book.VBProject
            $module = $project.VBComponents.Add(1); $module.Name = 'NativeHistoryProbe'
            $module.CodeModule.AddFromString("Option Explicit`r`nPublic Sub Example()`r`nEnd Sub")
            $module.CodeModule.CodePane.Show()
            $otherModule = $project.VBComponents.Add(1); $otherModule.Name = 'LastEditedHistory'
            $otherModule.CodeModule.AddFromString("Option Explicit`r`n' Last edit in a different module")
            $before = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
            $otherBefore = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$otherModule.Name }
            $module.CodeModule.CodePane.Show()
            $module.CodeModule.CodePane.Window.SetFocus()
            Add-Type -AssemblyName System.Windows.Forms
            for ($i=0; $i -lt 10; $i++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 50 }
            $settings = [Activator]::CreateInstance($assembly.GetType('VBAi.LlmSettings'), $true)
            $settings.VbeEditApproval = 'Automatic'
            $llm = [Activator]::CreateInstance($assembly.GetType('VBAi.LlmVbeTools'), [object[]]@($session,$null,$settings))
            $llm.BoundProject = $project.Name
            Add-Type -TypeDefinition 'public static class HistoryDiffCounter { public static int Count; public static string Modules = ""; public static void Changed(object change) { Count++; Modules += change.GetType().GetProperty("Module").GetValue(change, null).ToString() + ";"; } }'
            $editedEvent = $llm.GetType().GetEvent('CodeEdited')
            $editedHandler = [Delegate]::CreateDelegate($editedEvent.EventHandlerType, [HistoryDiffCounter].GetMethod('Changed'))
            $editedEvent.AddEventHandler($llm, $editedHandler)
            function Invoke-HistoryTool([hashtable]$Fields) {
                $Fields.Remove('Command')
                $response = $llm.Invoke('native_code_history', ($Fields | ConvertTo-Json -Compress)) | ConvertFrom-Json
                if (-not $response.Ok) { throw $response.Error }
                return $response.Data
            }
            $state = Invoke-Session @{ Command='native_code_history_state'; Project=$project.Name }
            $undoCaption = ($state.Commands | Where-Object { $_.Id -eq 128 -and $_.Enabled } | Select-Object -First 1).Caption
            $undo = Invoke-HistoryTool @{ Command='native_code_history'; Project=$project.Name; Action='undo'; ExpectedMode=2; ExpectedProjectVersion=$state.HistoryVersion; ControlCaption=$undoCaption }
            $empty = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$otherModule.Name }
            $undoState = Invoke-Session @{ Command='native_code_history_state'; Project=$project.Name }
            $redoCaption = ($undoState.Commands | Where-Object { $_.Id -eq 129 -and $_.Enabled } | Select-Object -First 1).Caption
            $redo = Invoke-HistoryTool @{ Command='native_code_history'; Project=$project.Name; Action='redo'; ExpectedMode=2; ExpectedProjectVersion=$undoState.HistoryVersion; ControlCaption=$redoCaption }
            $restored = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$otherModule.Name }
            $currentState = Invoke-Session @{ Command='native_code_history_state'; Project=$project.Name }
            $undoCaption = ($currentState.Commands | Where-Object { $_.Id -eq 128 -and $_.Enabled } | Select-Object -First 1).Caption
            $staleRejected = $false
            try { Invoke-Session @{ Command='native_code_history'; Project=$project.Name; Action='undo'; ExpectedMode=2; ExpectedProjectVersion=$undoState.HistoryVersion; ControlCaption=$undoCaption } | Out-Null }
            catch { $staleRejected = $_.Exception.Message -match 'changed since' }
            $badCaptionRejected = $false
            try { Invoke-Session @{ Command='native_code_history'; Project=$project.Name; Action='undo'; ExpectedMode=2; ExpectedProjectVersion=$currentState.HistoryVersion; ControlCaption='invalid' } | Out-Null }
            catch { $badCaptionRejected = $_.Exception.Message -match 'absent or disabled' }
            $afterGuards = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$otherModule.Name }
            $after = Invoke-Session @{ Command='read_module'; Project=$project.Name; Module=$module.Name }
            $otherBook = $excel.Workbooks.Add()
            $otherBook.VBProject.Name = 'OtherHistoryProject'
            $multiRejected = $false
            try { Invoke-Session @{ Command='native_code_history'; Project=$project.Name; Action='undo'; ExpectedMode=2; ExpectedProjectVersion=$currentState.HistoryVersion; ControlCaption=$undoCaption } | Out-Null }
            catch { $multiRejected = $_.Exception.Message -match 'close other projects' }
            $otherBook.Close($false); $otherBook = $null
            $form = $project.VBComponents.Add(3)
            $formRejected = $false
            try { Invoke-Session @{ Command='native_code_history'; Project=$project.Name; Action='undo'; ExpectedMode=2; ExpectedProjectVersion=$currentState.HistoryVersion; ControlCaption=$undoCaption } | Out-Null }
            catch { $formRejected = $_.Exception.Message -match 'containing UserForms' }
            [pscustomobject]@{ ChatDiffCount=[HistoryDiffCounter]::Count; ChatDiffModules=[HistoryDiffCounter]::Modules; Undo=$undo; Redo=$redo; Before=$otherBefore; Empty=$empty; Restored=$restored; StaleRejected=$staleRejected; BadCaptionRejected=$badCaptionRejected; MultipleProjectsRejected=$multiRejected; UserFormRejected=$formRejected; ActiveModulePreserved=($before.Sha256 -eq $after.Sha256); GuardsPreservedCode=($afterGuards.Sha256 -eq $restored.Sha256); Mvid=$assembly.ManifestModule.ModuleVersionId.ToString() } | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $directory 'native-code-history.json') -Encoding UTF8
            if ([HistoryDiffCounter]::Count -ne 2 -or [HistoryDiffCounter]::Modules -ne 'LastEditedHistory;LastEditedHistory;') { throw 'Chat diffs did not match native history changes.' }
            if (-not $undo.Verified -or -not $redo.Verified -or $empty.Code -ne '' -or $restored.Sha256 -ne $otherBefore.Sha256 -or $undo.Changes[0].Module -ne $otherModule.Name -or $before.Sha256 -ne $after.Sha256 -or -not $staleRejected -or -not $badCaptionRejected -or -not $multiRejected -or -not $formRejected -or $afterGuards.Sha256 -ne $restored.Sha256) { throw 'Native shared history cycle or guards failed.' }
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
    'NativeDefinition' {
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
        . $VbeProbeInitializeAssembly
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

    }
}

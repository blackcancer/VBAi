param(
    [Parameter(Mandatory=$true)][string]$AssemblyPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [switch]$UseBridge,
    [switch]$AllowTemporaryVbaAccess
)
$ErrorActionPreference = 'Stop'
if (-not $UseBridge -or @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'An isolated Excel session is required.' }
Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class MonacoStartupWindow {
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr child);
    [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr window, out Rect rect);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr window, out Rect rect);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
    public delegate bool WindowCallback(IntPtr window, IntPtr parameter);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, WindowCallback callback, IntPtr parameter);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr window, uint command);
    public static bool IsAboveInWorkspace(IntPtr editor, IntPtr native) {
        IntPtr workspace = GetParent(editor);
        while (native != IntPtr.Zero && GetParent(native) != workspace) native = GetParent(native);
        if (native == IntPtr.Zero) return false;
        for (IntPtr child = GetWindow(workspace, 5); child != IntPtr.Zero; child = GetWindow(child, 2)) {
            if (child == editor) return true;
            if (child == native) return false;
        }
        return false;
    }
    public static IntPtr FindCodePane(IntPtr root, string caption) {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(root, (window, unused) => {
            var kind = new StringBuilder(128); GetClassName(window, kind, kind.Capacity);
            var text = new StringBuilder(256); GetWindowText(window, text, text.Capacity);
            if (kind.ToString() == "VbaWindow" && text.ToString() == caption) { found = window; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }
    public static IntPtr FindEditor(IntPtr parent) {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(parent, (window, unused) => {
            var text = new StringBuilder(256); GetWindowText(window, text, text.Capacity);
            if (text.ToString() == "VBAi editor" || text.ToString() == "\u00c9diteur VBAi") { found = window; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
'@
function Wait-Startup([scriptblock]$read) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do { $value = & $read; if ($value) { return $value }; Start-Sleep -Milliseconds 100 } while ($watch.Elapsed.TotalSeconds -lt 40)
    throw 'The automatic Monaco workspace did not become available.'
}
function Read-FillProof([IntPtr]$editorHandle) {
    $workspace = [MonacoStartupWindow]::GetParent($editorHandle)
    $name = [Text.StringBuilder]::new(128)
    [void][MonacoStartupWindow]::GetClassName($workspace,$name,$name.Capacity)
    if ($name.ToString() -notlike '*MDIClient*') { throw 'Monaco is not attached to the VBE document workspace.' }
    $client = [MonacoStartupWindow+Rect]::new(); $editor = [MonacoStartupWindow+Rect]::new()
    if (-not [MonacoStartupWindow]::GetClientRect($workspace,[ref]$client) -or -not [MonacoStartupWindow]::GetWindowRect($editorHandle,[ref]$editor)) { throw 'The owned workspace bounds are unavailable.' }
    $width = $editor.Right-$editor.Left; $height = $editor.Bottom-$editor.Top
    if ($width -ne $client.Right -or $height -ne $client.Bottom -or $width -le 0 -or $height -le 0) { return $null }
    return @{ ParentClass=$name.ToString(); Width=$width; Height=$height; FillsWorkspace=$true }
}
$excel=$null; $book=$null; $vbe=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$true; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false
    $book=$excel.Workbooks.Add()
    $vbe=$excel.VBE; $vbe.MainWindow.Visible=$true; $vbe.MainWindow.SetFocus()
    [uint32]$ownerProcess=0
    [void][MonacoStartupWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$ownerProcess)
    $response=& (Join-Path $PSScriptRoot '../Invoke-CodexVBE.ps1') -HostProcessId $ownerProcess -RequestJson '{"Command":"status"}' -ResponseTimeoutSeconds 30 | ConvertFrom-Json
    $assembly=[Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath).Path)
    if (-not $response.Ok -or $response.Data.AssemblyModuleVersionId -ne $assembly.ManifestModule.ModuleVersionId.ToString('D')) { throw ('Loaded assembly mismatch: ' + ($response.Data | ConvertTo-Json -Compress)) }
    $editorHandle=Wait-Startup { $handle=[MonacoStartupWindow]::FindEditor([IntPtr]$vbe.MainWindow.HWnd); if ($handle -ne [IntPtr]::Zero) { $handle } }
    Write-Output ('Automatic Monaco child found: ' + $editorHandle)
    if (-not [MonacoStartupWindow]::IsWindowVisible($editorHandle)) { throw 'The automatic editor is hidden.' }
    $initial=Wait-Startup { Read-FillProof $editorHandle }
    Write-Output ('Initial workspace: ' + $initial.Width + 'x' + $initial.Height)
    $pane=$vbe.VBProjects.Item(1).VBComponents.Item('ThisWorkbook').CodeModule.CodePane
    $pane.Window.Visible=$true; $pane.Window.SetFocus()
    $codeHandle=[IntPtr]$pane.Window.HWnd
    # Some VBE code windows expose HWnd=0 through COM; resolve the owned native pane by its caption.
    if ($codeHandle -eq [IntPtr]::Zero) {
        $codeHandle=Wait-Startup {
            $handle=[MonacoStartupWindow]::FindCodePane([IntPtr]$vbe.MainWindow.HWnd,[string]$pane.Window.Caption)
            if ($handle -ne [IntPtr]::Zero) { $handle }
        }
    }
    try {
        [void](Wait-Startup { if ([MonacoStartupWindow]::IsAboveInWorkspace($editorHandle,$codeHandle)) { $true } })
    } catch {
        $chain=@(); $window=$codeHandle
        while ($window -ne [IntPtr]::Zero) {
            $name=[Text.StringBuilder]::new(128)
            [void][MonacoStartupWindow]::GetClassName($window,$name,$name.Capacity)
            $chain+=@{ Handle=$window.ToInt64(); Class=$name.ToString() }
            $window=[MonacoStartupWindow]::GetParent($window)
        }
        Write-Output (@{ Editor=$editorHandle.ToInt64(); Workspace=([MonacoStartupWindow]::GetParent($editorHandle)).ToInt64(); CodeChain=$chain; ActiveType=[int]$vbe.ActiveWindow.Type } | ConvertTo-Json -Depth 5)
        throw
    }
    if (-not $pane.Window.Visible -or -not [MonacoStartupWindow]::IsWindowVisible($editorHandle)) { throw 'The native code pane and Monaco must remain visible.' }
    Write-Output 'PASS: the native code window remains open behind Monaco after native activation.'
    $width=$vbe.MainWindow.Width
    $vbe.MainWindow.Width=[Math]::Max(700,$width-120)
    $resized=Wait-Startup { $proof=Read-FillProof $editorHandle; if ($proof -and $proof.Width -ne $initial.Width) { $proof } }
    $vbe.MainWindow.Width=$width
    $restored=Wait-Startup { $proof=Read-FillProof $editorHandle; if ($proof -and $proof.Width -eq $initial.Width) { $proof } }
    $browser=@($vbe.Windows | Where-Object { $_.Type -eq 2 })[0]
    $browser.Visible=$true; $browser.SetFocus()
    $browserHandle=[IntPtr]$browser.HWnd
    $browserParent=[MonacoStartupWindow]::GetParent($browserHandle)
    $browserBounds=[MonacoStartupWindow+Rect]::new()
    [void][MonacoStartupWindow]::GetWindowRect($browserHandle,[ref]$browserBounds)
    [void](Wait-Startup { if (-not [MonacoStartupWindow]::IsWindowVisible($editorHandle)) { $true } })
    Start-Sleep -Milliseconds 600
    $browserAfter=[MonacoStartupWindow+Rect]::new()
    [void][MonacoStartupWindow]::GetWindowRect($browserHandle,[ref]$browserAfter)
    if (-not $browser.Visible -or [MonacoStartupWindow]::GetParent($browserHandle) -ne $browserParent -or
        $browserAfter.Left -ne $browserBounds.Left -or $browserAfter.Top -ne $browserBounds.Top -or
        $browserAfter.Right -ne $browserBounds.Right -or $browserAfter.Bottom -ne $browserBounds.Bottom) { throw 'Monaco disturbed the native Object Browser layout.' }
    [IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($OutputDirectory)) | Out-Null
    @{ ProcessId=$ownerProcess; OpenedAutomatically=$true; MenuInvoked=$false; Initial=$initial; Resized=$resized; Restored=$restored; NativeCodePaneVisible=$true; NativeCodePaneBehindMonaco=$true; NativeObjectBrowserVisible=$true; NativeObjectBrowserLayoutPreserved=$true; MonacoHiddenForNativeBrowser=$true } |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'monaco-startup.json') -Encoding UTF8
    Write-Output 'PASS: Monaco opened automatically, filled MDIClient, resized and restored without a menu, shortcut or pointer input.'
} finally {
    if ($null -ne $book) { try { $book.Close($false) } catch { }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $vbe) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($vbe) }
    if ($null -ne $excel) { try { $excel.Quit() } catch { }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

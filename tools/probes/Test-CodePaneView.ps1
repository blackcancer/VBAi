param(
    [Parameter(Mandatory = $true)][string]$AssemblyPath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [switch]$Split,
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
    $module.CodeModule.AddFromString(((1..150 | ForEach-Object { "' Scroll probe line $_" }) -join "`r`n"))
    if ($Split) {
        $layout = Invoke-Session @{ Command = 'code_pane_layout'; Project = $project.Name; Module = $module.Name }
        $splitControl = $vbe.CommandBars.FindControl(1,302)
        $splitResult = Invoke-Session @{ Command = 'set_code_split'; Project = $project.Name; Module = $module.Name; ExpectedSha256 = $layout.Panes[0].State.Sha256; ExpectedMode = 2; StartLine = 2; Action = 'split'; ControlCaption = $splitControl.Caption }
        if (-not $splitResult.Verified) { throw 'Native code window did not split.' }
    }
    & (Join-Path $PSScriptRoot 'Inspect-VbeNativeWindows.ps1') -View Tree -HostProcessId $probeProcess.Id | Set-Content -LiteralPath (Join-Path $directory 'code-view-window-tree.json') -Encoding UTF8
    & (Join-Path $PSScriptRoot 'Inspect-DebugTree.ps1') -HostProcessId $probeProcess.Id -NamePattern '.*' | Set-Content -LiteralPath (Join-Path $directory 'code-view-accessibility.json') -Encoding UTF8
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

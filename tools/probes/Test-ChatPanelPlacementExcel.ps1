param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '../../artifacts/chat-panel-placement')
)
$ErrorActionPreference = 'Stop'
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close existing Excel sessions before running this disposable-instance test.' }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
[void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
public static class VbeSystemCloseProbe {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wparam, IntPtr lparam);
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
    [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr hwnd, uint cmd);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    public delegate bool EnumProc(IntPtr hwnd, IntPtr param);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr param);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr param);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr hwnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder text, int count);
    public class WindowInfo { public long Hwnd, Parent; public string Text, Class; public bool Visible; public Rect Bounds; }
    public static List<WindowInfo> Windows(uint pid) {
        var result=new List<WindowInfo>();
        EnumProc add=(h,p)=> { uint process; GetWindowThreadProcessId(h,out process); if(process!=pid)return true;
            var text=new StringBuilder(512); GetWindowText(h,text,512); var cls=new StringBuilder(256); GetClassName(h,cls,256); Rect rect; GetWindowRect(h,out rect);
            result.Add(new WindowInfo { Hwnd=h.ToInt64(), Parent=GetParent(h).ToInt64(), Text=text.ToString(), Class=cls.ToString(), Visible=IsWindowVisible(h), Bounds=rect }); return true; };
        EnumWindows((h,p)=> { uint process; GetWindowThreadProcessId(h,out process); if(process==pid) { add(h,p); EnumChildWindows(h,add,p); } return true; },IntPtr.Zero); return result;
    }
}
'@
Add-Type -AssemblyName System.Drawing

$excel = $null; $book = $null; $vbe = $null
$states = [Collections.Generic.List[object]]::new()
function Read-State([string]$stage) {
    $native = @([VbeSystemCloseProbe]::Windows($script:excelProcessId) | Where-Object { $_.Text -eq 'VBAi' -and $_.Class -eq 'GenericPane' })
    $state = [pscustomobject]@{
        Stage=$stage; Time=(Get-Date).ToString('o'); ExcelPid=$script:excelProcessId
        EditorVisible=[bool]$script:vbe.MainWindow.Visible; NativeChat=$native
        AddIns=@($script:vbe.AddIns | ForEach-Object { [pscustomobject]@{ ProgId=$_.ProgId; Connect=$_.Connect } })
    }
    $script:states.Add($state)
    $script:states | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'states.json') -Encoding UTF8
    $state | ConvertTo-Json -Depth 6 -Compress
    $connected = @($state.AddIns | Where-Object { $_.ProgId -eq 'VBAi.AddIn' -and $_.Connect })
    if ($connected.Count -ne 1) { throw 'The add-in is not connected.' }
    if ($stage -eq 'closed') {
        if ($state.EditorVisible -or @($native | Where-Object Visible).Count) { throw 'SC_CLOSE did not hide the editor and its pane.' }
    } else {
        if ($native.Count -ne 1 -or -not $native[0].Visible -or
            ($native[0].Bounds.Right-$native[0].Bounds.Left) -lt 440 -or
            ($native[0].Bounds.Bottom-$native[0].Bounds.Top) -lt 560) {
            throw 'The native chat pane is absent, hidden or smaller than the chat minimum.'
        }
    }
}
function Export-PaneImage([string]$stage) {
    $pane = @([VbeSystemCloseProbe]::Windows($script:excelProcessId) | Where-Object { $_.Text -eq 'VBAi' -and $_.Class -eq 'GenericPane' })[0]
    $handle = [VbeSystemCloseProbe]::GetAncestor([IntPtr]$pane.Hwnd,2)
    $rect = New-Object VbeSystemCloseProbe+Rect
    [void][VbeSystemCloseProbe]::GetWindowRect($handle,[ref]$rect)
    $bitmap = New-Object Drawing.Bitmap(($rect.Right-$rect.Left),($rect.Bottom-$rect.Top))
    try {
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $hdc = $graphics.GetHdc()
            try {
                if (-not [VbeSystemCloseProbe]::PrintWindow($handle,$hdc,2)) { throw 'PrintWindow failed.' }
            } finally { $graphics.ReleaseHdc($hdc) }
        } finally { $graphics.Dispose() }
        $bitmap.Save((Join-Path $OutputDirectory ($stage + '.png')),[Drawing.Imaging.ImageFormat]::Png)
    } finally { $bitmap.Dispose() }
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true; $excel.DisplayAlerts = $false
    [uint32]$excelProcessId = 0
    [void][VbeSystemCloseProbe]::GetWindowThreadProcessId([IntPtr][long]$excel.Hwnd,[ref]$excelProcessId)
    $book = $excel.Workbooks.Add()
    $excel.CommandBars.ExecuteMso('VisualBasic')
    $vbe = $excel.VBE
    Start-Sleep -Seconds 3
    Read-State 'initial'
    Export-PaneImage 'initial'
    # SC_CLOSE is the system command used by the VBE title-bar cross. Do not send WM_CLOSE.
    $editorHandle = [IntPtr][long]$vbe.MainWindow.HWnd
    if (-not [VbeSystemCloseProbe]::PostMessage($editorHandle,0x0112,[IntPtr]0xF060,[IntPtr]::Zero)) { throw 'SC_CLOSE could not be posted.' }
    Start-Sleep -Seconds 2
    Read-State 'closed'
    $excel.CommandBars.ExecuteMso('VisualBasic')
    Start-Sleep -Seconds 3
    Read-State 'reopened'
    Export-PaneImage 'reopened'
}
finally {
    try { if ($vbe) { $vbe.MainWindow.Visible = $false } } catch { Write-Warning $_.Exception.Message }
    $vbe = $null
    try { if ($book) { $book.Close($false) } } catch { Write-Warning $_.Exception.Message }
    $book = $null
    try { if ($excel) { $excel.Quit() } } catch { Write-Warning $_.Exception.Message }
    $excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($excelProcessId) {
        $remaining = Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
        if ($remaining -and -not $remaining.WaitForExit(5000)) { throw "Disposable Excel PID=$excelProcessId did not exit; no forced termination was attempted." }
    }
}

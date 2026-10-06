# Capture an existing, explicitly selected native product window.
# This module never creates fixtures, sends prompts or starts an Office host.
function Invoke-HelpCapture {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$AssemblyPath,
        [Parameter(Mandatory=$true)][ValidateRange(1,2147483647)][int]$HostProcessId,
        [Parameter(Mandatory=$true)][ValidateRange(1,9223372036854775807)][long]$WindowHandle,
        [Parameter(Mandatory=$true)][ValidatePattern('^[a-z][a-z0-9-]*$')][string]$CaptureId,
        [string]$OutputDirectory='artifacts/help-captures',
        [ValidateSet('Native','Screen')][string]$CaptureRenderer='Native',
        [string]$Culture='fr-FR'
    )
    $ErrorActionPreference='Stop'
    Add-Type -AssemblyName System.Drawing
    if(-not ('VbaiHelpWindowCapture' -as [type])) {
        Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
public static class VbaiHelpWindowCapture {
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left,Top,Right,Bottom; }
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr window,out Rect bounds);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr window,uint flag);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr window,StringBuilder text,int count);
    [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr window,IntPtr dc,uint flags);
    public static string Title(IntPtr window) {
        var text=new StringBuilder(2048);GetWindowText(window,text,text.Capacity);
        if(text.Length==0) { GetWindowText(GetAncestor(window,2),text,text.Capacity); }
        return text.ToString();
    }
    public static void Save(IntPtr window,string path,bool screen) {
        Rect r;if(!GetWindowRect(window,out r)||r.Right<=r.Left||r.Bottom<=r.Top)throw new InvalidOperationException("Window has no capture bounds.");
        using(var bitmap=new Bitmap(r.Right-r.Left,r.Bottom-r.Top)) {
            using(var graphics=Graphics.FromImage(bitmap)) {
                if(screen) graphics.CopyFromScreen(r.Left,r.Top,0,0,bitmap.Size);
                else { var dc=graphics.GetHdc();try { if(!PrintWindow(window,dc,2))throw new InvalidOperationException("Native renderer refused capture."); }finally {graphics.ReleaseHdc(dc);} }
            }
            bool varied=false;int first=bitmap.GetPixel(0,0).ToArgb();
            for(int y=0;y<bitmap.Height&&!varied;y+=Math.Max(1,bitmap.Height/80))
                for(int x=0;x<bitmap.Width;x+=Math.Max(1,bitmap.Width/80))
                    if(bitmap.GetPixel(x,y).ToArgb()!=first){varied=true;break;}
            if(!varied)throw new InvalidOperationException("Renderer returned a uniform image; no capture accepted.");
            bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);
        }
    }
}
"@
    }
    $candidate=[IO.Path]::GetFullPath($AssemblyPath)
    if(-not (Test-Path -LiteralPath $candidate -PathType Leaf)){throw 'Candidate assembly does not exist.'}
    $process=Get-Process -Id $HostProcessId -ErrorAction Stop
    $birth=$process.StartTime.ToUniversalTime().ToString('o')
    $hwnd=[IntPtr]$WindowHandle;$actualPid=[uint32]0
    [void][VbaiHelpWindowCapture]::GetWindowThreadProcessId($hwnd,[ref]$actualPid)
    if($actualPid -ne $HostProcessId -or -not [VbaiHelpWindowCapture]::IsWindowVisible($hwnd)){throw 'The selected window is not visible in the selected process.'}
    $title=[VbaiHelpWindowCapture]::Title($hwnd)
    if([string]::IsNullOrWhiteSpace($title)){throw 'The real window title is unavailable.'}
    # Managed LoadFrom assemblies are not reliably listed in Process.Modules.
    # Ask the PID-scoped product bridge for the exact loaded assembly identity.
    $raw=& (Join-Path $PSScriptRoot '../Invoke-VBAi.ps1') -HostProcessId $HostProcessId -Command status
    $status=$raw|ConvertFrom-Json
    if(-not $status.Ok -or $status.Data.HostProcessId -ne $HostProcessId -or
        -not [string]::Equals([IO.Path]::GetFullPath($status.Data.AssemblyPath),$candidate,[StringComparison]::OrdinalIgnoreCase)){
        throw 'The selected product bridge does not identify the specified candidate assembly.'
    }
    $directory=[IO.Path]::GetFullPath($OutputDirectory)
    [void][IO.Directory]::CreateDirectory($directory)
    $path=Join-Path $directory ($CaptureId+'.png')
    if(Test-Path -LiteralPath $path){throw 'Capture ID already exists; choose a new ID to preserve evidence.'}
    [VbaiHelpWindowCapture]::Save($hwnd,$path,($CaptureRenderer -eq 'Screen'))
    [void][VbaiHelpWindowCapture]::GetWindowThreadProcessId($hwnd,[ref]$actualPid)
    if($actualPid -ne $HostProcessId -or $process.HasExited){throw 'Window ownership changed during capture; inspect the unaccepted PNG.'}
    $bitmap=[Drawing.Bitmap]::FromFile($path)
    try {
        $record=[pscustomobject]@{id=$CaptureId;file=($CaptureId+'.png');provenance='live-interface';synthetic=$false;
            windowTitle=$title;processId=$HostProcessId;processStartedUtc=$birth;windowHandle=$WindowHandle;
            capturedUtc=[DateTime]::UtcNow.ToString('o');renderer=$CaptureRenderer;culture=$Culture;
            width=$bitmap.Width;height=$bitmap.Height;sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();
            assemblyMvid=$status.Data.AssemblyModuleVersionId;assemblySha256=(Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash.ToLowerInvariant();
            review='Pending visual and workflow review. A varied image does not establish a successful operation.'}
    }finally{$bitmap.Dispose();$process.Dispose()}
    $record|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $directory ($CaptureId+'.json')) -Encoding UTF8
    $record
}
Export-ModuleMember -Function Invoke-HelpCapture

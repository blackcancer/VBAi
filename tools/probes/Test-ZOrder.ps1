param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [string] $Project = 'VBAProject',
    [string] $Form = 'CodexZOrderSurvey'
)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') {
    throw "The probe requires an Excel process; PID $HostProcessId is $($hostProcess.ProcessName)."
}

Add-Type -ReferencedAssemblies 'System.Drawing' -TypeDefinition @'
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

public static class ZOrderCapture {
    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr state);
    [StructLayout(LayoutKind.Sequential)] private struct Rect {
        public int Left, Top, Right, Bottom;
    }
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr state);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder value, int maxCount);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);

    public static string Capture(int processId, string path) {
        IntPtr editor = IntPtr.Zero;
        EnumWindows((hwnd, state) => {
            uint owner;
            GetWindowThreadProcessId(hwnd, out owner);
            if (owner != (uint)processId || !IsWindowVisible(hwnd)) return true;
            var name = new StringBuilder(128);
            GetClassName(hwnd, name, name.Capacity);
            if (!name.ToString().StartsWith("wndclass_desked", StringComparison.OrdinalIgnoreCase)) return true;
            editor = hwnd;
            return false;
        }, IntPtr.Zero);
        if (editor == IntPtr.Zero) throw new InvalidOperationException("Visible VBE window not found for target process.");
        Rect rect;
        if (!GetWindowRect(editor, out rect)) throw new InvalidOperationException("Cannot read VBE window bounds.");
        int width = rect.Right - rect.Left, height = rect.Bottom - rect.Top;
        if (width <= 0 || height <= 0 || width > 8192 || height > 8192)
            throw new InvalidOperationException("VBE window size is outside the capture limit.");
        using (var image = new Bitmap(width, height))
        using (var graphics = Graphics.FromImage(image)) {
            graphics.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(width, height));
            image.Save(path, ImageFormat.Png);
        }
        return path;
    }
}
'@

function Invoke-Vbe([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $reply = & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $reply.Ok) { throw "$($Request.Command): $($reply.Error)" }
    return $reply.Data
}

function Read-Tree {
    Invoke-Vbe @{ Command = 'form_tree'; Project = $Project; Form = $Form }
}

function Set-Node([string] $Path, [string] $Property, $Value) {
    $tree = Read-Tree
    Invoke-Vbe @{ Command = 'set_form_node_property'; Project = $Project; Form = $Form;
        ControlPath = $Path; Property = $Property; Value = $Value;
        ExpectedTreeVersion = $tree.TreeVersion } | Out-Null
}

$existing = @(Invoke-Vbe @{ Command = 'list_forms'; Project = $Project })
if (@($existing | Where-Object { $_.Name -eq $Form }).Count -gt 0) {
    throw "The probe form $Form already exists; choose a new Form name in a disposable workbook."
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$artifactDirectory = Join-Path $PSScriptRoot "..\..\artifacts\zorder\$HostProcessId-$stamp"
$artifactDirectory = [IO.Path]::GetFullPath($artifactDirectory)
New-Item -ItemType Directory -Path $artifactDirectory -Force | Out-Null

Invoke-Vbe @{ Command = 'create_form'; Project = $Project; Form = $Form } | Out-Null
foreach ($item in @(
    @{ Name = 'lblRed'; Caption = 'ROUGE'; Left = 24; Top = 24 },
    @{ Name = 'lblBlue'; Caption = 'BLEU'; Left = 50; Top = 38 }
)) {
    $tree = Read-Tree
    Invoke-Vbe @{ Command = 'add_form_control'; Project = $Project; Form = $Form;
        ControlType = 'Forms.Label.1'; Control = $item.Name; Caption = $item.Caption;
        Left = $item.Left; Top = $item.Top; Width = 100; Height = 40;
        ExpectedFormVersion = $tree.FormVersion } | Out-Null
}

Set-Node 'Controls/lblRed' 'BackStyle' 1
Set-Node 'Controls/lblRed' 'BackColor' '#FF0000'
Set-Node 'Controls/lblBlue' 'BackStyle' 1
Set-Node 'Controls/lblBlue' 'BackColor' '#0000FF'
Invoke-Vbe @{ Command = 'open_form'; Project = $Project; Form = $Form } | Out-Null
Start-Sleep -Milliseconds 350
$before = Read-Tree
$beforeImage = [ZOrderCapture]::Capture($HostProcessId, (Join-Path $artifactDirectory 'before.png'))

$front = Invoke-Vbe @{ Command = 'z_order_form_control'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/lblRed'; ZPosition = 0; ExpectedTreeVersion = $before.TreeVersion }
Start-Sleep -Milliseconds 350
$frontTree = Read-Tree
$frontImage = [ZOrderCapture]::Capture($HostProcessId, (Join-Path $artifactDirectory 'red-front.png'))

$back = Invoke-Vbe @{ Command = 'z_order_form_control'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/lblRed'; ZPosition = 1; ExpectedTreeVersion = $frontTree.TreeVersion }
Start-Sleep -Milliseconds 350
$backTree = Read-Tree
$backImage = [ZOrderCapture]::Capture($HostProcessId, (Join-Path $artifactDirectory 'red-back.png'))

[pscustomobject]@{
    HostProcessId = $HostProcessId
    Project = $Project
    Form = $Form
    ArtifactDirectory = $artifactDirectory
    BeforeImage = $beforeImage
    RedFrontImage = $frontImage
    RedBackImage = $backImage
    BeforeTreeVersion = $before.TreeVersion
    RedFrontTreeVersion = $frontTree.TreeVersion
    RedBackTreeVersion = $backTree.TreeVersion
    RedFrontExecuted = $front.Executed
    RedBackExecuted = $back.Executed
    RedFrontVerification = $front.Verification
    RedBackVerification = $back.Verification
    NodeCount = $backTree.NodeCount
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

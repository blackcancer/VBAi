param(
    [string]$AssemblyPath = 'artifacts/chat-designer/build/VBAi/Debug/net48/VBAi.dll',
    [string]$OutputDirectory = 'artifacts/chat-designer/render',
    [ValidateSet('Light','Dark')][string]$Theme = 'Light',
    [int]$Width = 720
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms,System.Drawing,PresentationFramework
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class ChatPreviewCapture {
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);
}
"@
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
$themeType = $assembly.GetType('VBAi.UiTheme')
$themeType.GetField('<Choice>k__BackingField',[Reflection.BindingFlags]'Static,NonPublic').SetValue($null,[Enum]::Parse($assembly.GetType('VBAi.ThemeChoice'),$Theme))
function New-Internal($name) { [Activator]::CreateInstance($assembly.GetType('VBAi.' + $name),$true) }
function Call($object,$name,$arguments) { $object.GetType().GetMethod($name,$flags).Invoke($object,$arguments) }
function Field($object,$name) { $object.GetType().GetField($name,$flags).GetValue($object) }
$window = New-Internal ChatWindow
$frame = [Windows.Forms.Form]::new()
$flow = [Windows.Forms.FlowLayoutPanel]::new()
$flow.FlowDirection = [Windows.Forms.FlowDirection]::TopDown
$flow.WrapContents = $false; $flow.AutoScroll = $true; $flow.Dock = [Windows.Forms.DockStyle]::Fill
$frame.Controls.Add($flow); $frame.Size = [Drawing.Size]::new($Width,1000)
$frame.Text = 'VBAi - Designer cards'
$themeType.GetMethod('Apply',[Reflection.BindingFlags]'Static,NonPublic').Invoke($null,@($frame))
$hosts = [Collections.Generic.List[object]]::new()
try {
    Call $window InitializeShell @()
    Call $window InitializeComposer @($null)
    Call $window InitializeTranscript @()
    $reference = New-Internal VbeChatReference
    $reference.Project = 'Budget'; $reference.Module = 'ModuleCalcul'; $reference.Name = 'CalculerTotal'; $reference.Kind = 'Function'
    $user = New-Internal ChatEntry; $user.Speaker='Vous'; $user.Text='Explain @Budget.ModuleCalcul.CalculerTotal'
    $assistant = New-Internal ChatEntry; $assistant.Speaker='Assistant'; $assistant.Text="## Result`n`nThe **last row** is now included in the total. Review @Budget.ModuleCalcul.CalculerTotal.`n`n- Code synchronized`n- Change available for rollback`n`n[Documentation](https://learn.microsoft.com)"
    $refs = [Array]::CreateInstance($reference.GetType(),1); $refs.SetValue($reference,0); $assistant.References=$refs
    $change = New-Internal CodeChange; $change.Project='Budget'; $change.Module='ModuleCalcul'; $change.Before="For i = 1 To lastRow - 1`n    total = total + Cells(i, 2).Value`nNext i"; $change.After="For i = 1 To lastRow`n    total = total + Cells(i, 2).Value`nNext i"
    $changed = New-Internal ChatEntry; $changed.Speaker='Code'; $changed.Change=$change
    foreach ($entry in @($user,$assistant,$changed)) {
        Call $window AddEntry @($entry)
        $hosted = Call $window RenderEntry @($entry)
        $view = $hosted.Child; $hosts.Add($hosted)
        $view.MinimumSize = [Drawing.Size]::Empty
        $view.MaximumSize = [Drawing.Size]::new($Width-50,0); $view.MinimumSize=$view.MaximumSize; $view.Width=$Width-50
        $hosted.Child = $null
        $flow.Controls.Add($view); $view.Visible=$true
        $view.Height = $view.GetPreferredSize([Drawing.Size]::new($Width-50,0)).Height
    }
    $frame.Show(); [Windows.Forms.Application]::DoEvents()
    $frame.Refresh(); [Windows.Forms.Application]::DoEvents()
    $bitmap = [Drawing.Bitmap]::new($frame.Width,$frame.Height)
    try {
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        $hdc = $graphics.GetHdc()
        try { if (-not [ChatPreviewCapture]::PrintWindow($frame.Handle,$hdc,2)) { throw 'Could not capture the native Designer preview.' } }
        finally { $graphics.ReleaseHdc($hdc); $graphics.Dispose() }
        $directory = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputDirectory)); [IO.Directory]::CreateDirectory($directory) | Out-Null
        $path = Join-Path $directory "designer-cards-$Theme-$Width.png"; $bitmap.Save($path,[Drawing.Imaging.ImageFormat]::Png); Write-Output $path
    } finally { $bitmap.Dispose() }
} finally { $frame.Dispose(); foreach ($hosted in $hosts) { $hosted.Dispose() }; $window.Dispose() }

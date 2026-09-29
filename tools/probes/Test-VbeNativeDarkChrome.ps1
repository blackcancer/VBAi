param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '../../artifacts/vbe-native-dark-chrome'),
    [switch]$InProcessAddInExperiment,
    [switch]$OpenCodeWindow,
    [switch]$CycleCodeWindow,
    [switch]$DebugPreview,
    [switch]$InspectPalette,
    [switch]$PreviewPalette,
    [switch]$CheckPropertyTabs,
    [switch]$ProductionPalette,
    [switch]$CheckNativeCombos,
    [switch]$SettingsThemeCycle,
    [switch]$CheckEditorInteractions,
    [switch]$CheckCommandMenus,
    [switch]$CheckNativeDialogs,
    [switch]$CheckPropertyRows,
    [string]$NativeRenderTraceDll,
    [switch]$CheckToolbarBlink,
    [switch]$LocalRefreshExperiment,
    [switch]$TraceToolbarModule,
    [switch]$ToolbarPatternPilot,
    [switch]$ToolbarOnly,
    [switch]$KeepToolbarProbeVisible
)

$ErrorActionPreference = 'Stop'
if ($TraceToolbarModule -and -not $NativeRenderTraceDll) { throw 'TraceToolbarModule requires NativeRenderTraceDll.' }
if ($ToolbarPatternPilot -and (-not $TraceToolbarModule -or -not $CheckToolbarBlink)) { throw 'ToolbarPatternPilot requires TraceToolbarModule and CheckToolbarBlink.' }
if (($ToolbarOnly -or $KeepToolbarProbeVisible) -and -not $CheckToolbarBlink) { throw 'ToolbarOnly and KeepToolbarProbeVisible require CheckToolbarBlink.' }
if ($LocalRefreshExperiment -and -not $InProcessAddInExperiment) { throw 'LocalRefreshExperiment requires the disposable in-process experiment.' }
if ($CheckToolbarBlink -and (-not $InProcessAddInExperiment -or -not $OpenCodeWindow -or -not $CheckPropertyRows)) {
    throw 'Toolbar observation requires the in-process experiment, code pane, and disposable UserForm fixture.'
}
if ($NativeRenderTraceDll -and (-not $InProcessAddInExperiment -or -not $OpenCodeWindow)) {
    throw 'NativeRenderTraceDll requires InProcessAddInExperiment and OpenCodeWindow.'
}
if ($NativeRenderTraceDll -and ($DebugPreview -or $CheckEditorInteractions)) {
    throw 'The render trace uses a non-executing code fixture; run editing/debugging probes separately.'
}
if ($NativeRenderTraceDll) {
    . (Join-Path $PSScriptRoot 'native-render-trace/Invoke-VbeRenderTrace.ps1')
}
if ($ProductionPalette -and ($PreviewPalette -or $InspectPalette)) { throw 'Choose one palette implementation per probe.' }
if ($SettingsThemeCycle -and ($InProcessAddInExperiment -or $ProductionPalette -or $PreviewPalette -or $InspectPalette)) { throw 'SettingsThemeCycle must drive the production setting without an experiment override.' }
if ($SettingsThemeCycle -and @(Get-Process SLDWORKS -ErrorAction SilentlyContinue).Count) {
    throw 'The settings/palette cycle requires SOLIDWORKS to be closed; VBE palette recovery is shared by version.'
}
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) {
    throw 'Close existing Excel sessions before running this disposable-instance probe.'
}

$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
[void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class VbeNativeDarkProbe
{
    public delegate bool EnumProc(IntPtr window, IntPtr parameter);

    [StructLayout(LayoutKind.Sequential)]
    public struct Rect { public int Left, Top, Right, Bottom; }

    public sealed class Result
    {
        public long Hwnd, Parent;
        public string Class, Caption;
        public int ThemeResult;
    }

    public sealed class WindowInfo
    {
        public long Hwnd, Parent;
        public string Class, Caption;
        public bool Visible;
        public int Style;
        public Rect Bounds;
    }

    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc callback, IntPtr parameter);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr parameter);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll", EntryPoint="GetWindowLongW")] public static extern int GetWindowStyle(IntPtr window, int index);
    [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr window);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr window);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll")] public static extern bool EnableWindow(IntPtr window, bool enabled);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr window, int command);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);
    [DllImport("user32.dll", EntryPoint="GetWindowLongPtrW")] public static extern IntPtr GetWindowLongPtr(IntPtr window, int index);
    [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr window, out Rect rect);
    [DllImport("user32.dll", SetLastError=true)] public static extern uint GetGuiResources(IntPtr process, uint flags);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr window, out Rect rect);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr window, IntPtr hdc, uint flags);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool RedrawWindow(IntPtr window, IntPtr update, IntPtr region, uint flags);
    [DllImport("user32.dll", EntryPoint="RedrawWindow")] private static extern bool RedrawRegion(IntPtr window, ref Rect update, IntPtr region, uint flags);
    [DllImport("uxtheme.dll", CharSet=CharSet.Unicode)] private static extern int SetWindowTheme(IntPtr window, string appName, string idList);
    [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);

    private const uint WmThemeChanged = 0x031A;
    private const uint RedrawFlags = 0x0001 | 0x0004 | 0x0080 | 0x0100;

    public static bool RedrawPartial(IntPtr window)
    {
        Rect rect;
        if (!GetClientRect(window, out rect) || rect.Right <= 0 || rect.Bottom <= 0) return false;
        rect.Left = rect.Right / 4;
        rect.Right = Math.Max(rect.Left + 1, rect.Right * 3 / 4);
        return RedrawRegion(window, ref rect, IntPtr.Zero, 0x0001 | 0x0100);
    }

    public static string ClassName(IntPtr window)
    {
        var text = new StringBuilder(256);
        GetClassName(window, text, text.Capacity);
        return text.ToString();
    }

    private static string Caption(IntPtr window)
    {
        var text = new StringBuilder(512);
        GetWindowText(window, text, text.Capacity);
        return text.ToString();
    }

    public static IntPtr FindEditor(uint processId)
    {
        IntPtr result = IntPtr.Zero;
        EnumWindows((window, parameter) => {
            uint owner;
            GetWindowThreadProcessId(window, out owner);
            if (owner == processId && ClassName(window) == "wndclass_desked_gsk") {
                result = window;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return result;
    }

    private static bool IsAddInWindow(string className)
    {
        return className.StartsWith("WindowsForms10.", StringComparison.Ordinal) ||
            className.StartsWith("HwndWrapper[", StringComparison.Ordinal) ||
            className == "GenericPane";
    }

    public static List<Result> Apply(IntPtr editor)
    {
        var windows = new List<IntPtr> { editor };
        EnumChildWindows(editor, (window, parameter) => { windows.Add(window); return true; }, IntPtr.Zero);
        var results = new List<Result>();
        int enabled = 1;
        foreach (IntPtr window in windows)
        {
            string className = ClassName(window);
            if (IsAddInWindow(className)) continue;
            int themeResult = SetWindowTheme(window, "DarkMode_Explorer", null);
            if (window == editor)
            {
                int dwmResult = DwmSetWindowAttribute(window, 20, ref enabled, sizeof(int));
                if (dwmResult != 0) DwmSetWindowAttribute(window, 19, ref enabled, sizeof(int));
            }
            SendMessage(window, WmThemeChanged, IntPtr.Zero, IntPtr.Zero);
            results.Add(new Result { Hwnd = window.ToInt64(), Parent = GetParent(window).ToInt64(),
                Class = className, Caption = Caption(window), ThemeResult = themeResult });
        }
        RedrawWindow(editor, IntPtr.Zero, IntPtr.Zero, RedrawFlags);
        return results;
    }

    public static List<WindowInfo> Snapshot(IntPtr editor)
    {
        var windows = new List<IntPtr> { editor };
        EnumChildWindows(editor, (window, parameter) => { windows.Add(window); return true; }, IntPtr.Zero);
        var result = new List<WindowInfo>();
        foreach (IntPtr window in windows)
        {
            Rect bounds;
            GetWindowRect(window, out bounds);
            result.Add(new WindowInfo { Hwnd = window.ToInt64(), Parent = GetParent(window).ToInt64(),
                Class = ClassName(window), Caption = Caption(window), Visible = IsWindowVisible(window), Style = GetWindowStyle(window, -16), Bounds = bounds });
        }
        return result;
    }
}
'@

$excel = $null
$book = $null
$excelProcessId = 0
$renderTrace = $null
$renderTraceStopError = $null
$settingsBackup = $null
$initialNativeTheme = $false
$initialPalettePath = $null
$settingsCycleRestored = $false
$settingsPath = Join-Path $env:APPDATA 'VBAi/settings.json'
if ($SettingsThemeCycle) {
    if (-not (Test-Path -LiteralPath $settingsPath)) { throw 'This settings-cycle probe requires an existing settings file.' }
    $initialNativeTheme = [bool](Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json).NativeVbeDarkTheme
    $paletteFiles = @(Get-ChildItem (Join-Path $env:LOCALAPPDATA 'VBAi/native-theme') -Filter '*.json' -ErrorAction SilentlyContinue)
    if ($paletteFiles.Count -gt 1 -or (-not $initialNativeTheme -and $paletteFiles.Count)) { throw 'The initial palette recovery state is ambiguous.' }
    if ($initialNativeTheme) {
        if ($paletteFiles.Count -ne 1) { throw 'An enabled theme requires its original palette recovery file for this probe.' }
        $initialPalettePath = $paletteFiles[0].FullName
        $paletteSnapshot = Join-Path $OutputDirectory 'initial-palette.json'
        [IO.File]::Copy($initialPalettePath, $paletteSnapshot, $false)
    }
    $settingsBackup = $settingsPath + '.native-theme-probe.bak'
    if (Test-Path -LiteralPath $settingsBackup) { throw 'A previous settings backup already exists.' }
    [IO.File]::Copy($settingsPath, $settingsBackup, $false)
    $wpf = Join-Path ([Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) 'WPF'
    Add-Type -Path (Join-Path $PSScriptRoot 'VbeThemeSettingsProbe.cs') -ReferencedAssemblies (Join-Path $wpf 'UIAutomationClient.dll'),(Join-Path $wpf 'UIAutomationTypes.dll'),(Join-Path $wpf 'WindowsBase.dll')
}
$previousExperiment = [Environment]::GetEnvironmentVariable('VBAi_NATIVE_DARK_EXPERIMENT', 'Process')
$previousLocalRefresh = [Environment]::GetEnvironmentVariable('VBAi_NATIVE_LOCAL_REFRESH_EXPERIMENT', 'Process')
if ($LocalRefreshExperiment) { [Environment]::SetEnvironmentVariable('VBAi_NATIVE_LOCAL_REFRESH_EXPERIMENT', '1', 'Process') }
else { [Environment]::SetEnvironmentVariable('VBAi_NATIVE_LOCAL_REFRESH_EXPERIMENT', $null, 'Process') }
if ($SettingsThemeCycle) { [Environment]::SetEnvironmentVariable('VBAi_NATIVE_DARK_EXPERIMENT', $null, 'Process') }
if ($InProcessAddInExperiment) {
    [Environment]::SetEnvironmentVariable('VBAi_NATIVE_DARK_EXPERIMENT', '1', 'Process')
}

function Export-WindowImage([IntPtr]$Handle, [string]$Name) {
    $rect = New-Object VbeNativeDarkProbe+Rect
    if (-not [VbeNativeDarkProbe]::GetWindowRect($Handle, [ref]$rect)) { throw 'GetWindowRect failed.' }
    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top
    if ($width -lt 1 -or $height -lt 1) { throw "Invalid editor dimensions: ${width}x${height}." }
    $bitmap = New-Object Drawing.Bitmap($width, $height)
    try {
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $hdc = $graphics.GetHdc()
            try {
                if (-not [VbeNativeDarkProbe]::PrintWindow($Handle, $hdc, 2)) { throw 'PrintWindow failed.' }
            } finally { $graphics.ReleaseHdc($hdc) }
        } finally { $graphics.Dispose() }
        $bitmap.Save((Join-Path $OutputDirectory ($Name + '.png')), [Drawing.Imaging.ImageFormat]::Png)
    } finally { $bitmap.Dispose() }
}

function Export-LiveWindowImage([IntPtr]$Handle, [string]$Name) {
    $liveBounds = New-Object VbeNativeDarkProbe+Rect
    if ([VbeNativeDarkProbe]::GetWindowRect($Handle, [ref]$liveBounds)) {
        $liveImage = New-Object Drawing.Bitmap ($liveBounds.Right - $liveBounds.Left),($liveBounds.Bottom - $liveBounds.Top)
        try {
            $liveGraphics = [Drawing.Graphics]::FromImage($liveImage)
            try { $liveGraphics.CopyFromScreen($liveBounds.Left, $liveBounds.Top, 0, 0, $liveImage.Size) }
            finally { $liveGraphics.Dispose() }
            $liveImage.Save((Join-Path $OutputDirectory ($Name + '.png')), [Drawing.Imaging.ImageFormat]::Png)
        } finally { $liveImage.Dispose() }
    }
}

function Get-ProbeGuiResourceSample([string]$Phase, [IntPtr]$TabHandle = [IntPtr]::Zero) {
    $hostProcess = Get-Process -Id $excelProcessId -ErrorAction Stop
    try {
        $gdi = [VbeNativeDarkProbe]::GetGuiResources($hostProcess.Handle, 0)
        if ($gdi -eq 0) { throw ('GetGuiResources failed: ' + [Runtime.InteropServices.Marshal]::GetLastWin32Error()) }
        $user = [VbeNativeDarkProbe]::GetGuiResources($hostProcess.Handle, 1)
        if ($user -eq 0) { throw ('GetGuiResources(USER) failed: ' + [Runtime.InteropServices.Marshal]::GetLastWin32Error()) }
        $paintCount = $null
        $printCount = $null
        $msoPaintChromeCount = $null
        $deferredMsoPassCount = $null
        if ($InProcessAddInExperiment -and $TabHandle -ne [IntPtr]::Zero) {
            $paintCount = [VbeNativeDarkProbe]::SendMessage($TabHandle, 0x856a, [IntPtr]::Zero, [IntPtr]::Zero)
            $printCount = [VbeNativeDarkProbe]::SendMessage($TabHandle, 0x856a, [IntPtr]1, [IntPtr]::Zero)
            $msoPaintChromeCount = [VbeNativeDarkProbe]::SendMessage($TabHandle, 0x856b, [IntPtr]::Zero, [IntPtr]::Zero)
            $deferredMsoPassCount = [VbeNativeDarkProbe]::SendMessage($TabHandle, 0x856b, [IntPtr]1, [IntPtr]::Zero)
        }
        [pscustomobject]@{
            Phase = $Phase
            Utc = [DateTime]::UtcNow.ToString('o')
            GdiObjects = $gdi
            UserObjects = $user
            CpuMilliseconds = $hostProcess.TotalProcessorTime.TotalMilliseconds
            TabPaintCount = $paintCount
            TabPrintCount = $printCount
            MsoPaintChromeCount = $msoPaintChromeCount
            DeferredMsoPassCount = $deferredMsoPassCount
        }
    } finally { $hostProcess.Dispose() }
}

function Export-PropertyTabState([IntPtr]$Handle, [IntPtr]$ParentHandle, [string]$Name) {
    # Capture the screen first: PrintWindow can itself cause a fresh paint.
    Export-LiveWindowImage $Handle ($Name + '-live')
    Export-LiveWindowImage $ParentHandle ($Name + '-pane-live')
    Export-WindowImage $Handle ($Name + '-print')
    [pscustomobject]@{
        Name = $Name
        Selected = [VbeNativeDarkProbe]::SendMessage($Handle, 0x130b, [IntPtr]::Zero, [IntPtr]::Zero)
        Visible = [VbeNativeDarkProbe]::IsWindowVisible($Handle)
        Enabled = [VbeNativeDarkProbe]::IsWindowEnabled($Handle)
    }
}

function Set-NativeThemeFromSettings([bool]$Enabled) {
    $logPath = Join-Path $env:TEMP 'VBAi-load.log'
    $expected = if ($Enabled) { 'Native editor palette applied and verified.' } else { 'Native editor palette restored and verified.' }
    $previous = @(Get-Content -LiteralPath $logPath -ErrorAction SilentlyContinue | Where-Object { $_.Contains($expected) }) | Select-Object -Last 1
    [VbeThemeSettingsProbe]::Begin($excelProcessId, $Enabled)
    $settingsCommand = $excel.VBE.CommandBars.FindControl(1, [Type]::Missing, 'VBAi.Settings')
    if (-not $settingsCommand) { throw 'The add-in Settings command was not found.' }
    $settingsCommand.Execute()
    if (-not [VbeThemeSettingsProbe]::Wait() -or [VbeThemeSettingsProbe]::Error -or -not [VbeThemeSettingsProbe]::Saved) { throw ('Settings UI failed: ' + [VbeThemeSettingsProbe]::Error) }
    $timeout = [Diagnostics.Stopwatch]::StartNew()
    do {
        Start-Sleep -Milliseconds 200
        $latest = @(Get-Content -LiteralPath $logPath -ErrorAction SilentlyContinue | Where-Object { $_.Contains($expected) }) | Select-Object -Last 1
        if ($latest -and $latest -ne $previous) {
            $stored = (Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json).NativeVbeDarkTheme
            if ([bool]$stored -ne $Enabled) { throw 'The native theme preference was not saved.' }
            return
        }
    } while ($timeout.ElapsedMilliseconds -lt 60000)
    throw 'The deferred native palette update did not report verified completion.'
}

try {
    $startupLogPath = Join-Path $env:TEMP 'VBAi-load.log'
    $startupMarker = 'Native editor palette applied and verified.'
    $previousStartup = @(Get-Content -LiteralPath $startupLogPath -ErrorAction SilentlyContinue | Where-Object { $_.Contains($startupMarker) }) | Select-Object -Last 1
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    [uint32]$owner = 0
    [void][VbeNativeDarkProbe]::GetWindowThreadProcessId([IntPtr][long]$excel.Hwnd, [ref]$owner)
    $excelProcessId = [int]$owner
    $book = $excel.Workbooks.Add()
    $excel.CommandBars.ExecuteMso('VisualBasic')
    if ($SettingsThemeCycle) {
        $productionRecovery = Join-Path $env:LOCALAPPDATA ('VBAi/native-theme/palette-' + $excel.VBE.Version + '.json')
        if ($initialNativeTheme) {
            if ($initialPalettePath -ne $productionRecovery) { throw 'The initial recovery file belongs to another VBE version.' }
            $startupWait = [Diagnostics.Stopwatch]::StartNew()
            do {
                Start-Sleep -Milliseconds 200
                $startup = @(Get-Content -LiteralPath $startupLogPath -ErrorAction SilentlyContinue | Where-Object { $_.Contains($startupMarker) }) | Select-Object -Last 1
            } while (($null -eq $startup -or $startup -eq $previousStartup) -and $startupWait.ElapsedMilliseconds -lt 60000)
            if ($null -eq $startup -or $startup -eq $previousStartup) { throw 'The initial native palette did not finish applying.' }
            Set-NativeThemeFromSettings $false
        }
        Set-NativeThemeFromSettings $true
        'Settings checkbox enabled; preference saved; deferred palette verified.' | Set-Content (Join-Path $OutputDirectory 'settings-enabled.txt')
    }
    if ($ProductionPalette) {
        $paletteAssembly = [Reflection.Assembly]::LoadFrom((Join-Path $PSScriptRoot '../../bin/Debug/net48/VBAi.dll'))
        $paletteChange = $paletteAssembly.GetType('VBAi.VbeNativePalette').GetMethod('Change', [Reflection.BindingFlags]'NonPublic,Static')
        $productionRecovery = Join-Path $OutputDirectory 'production-palette-recovery.json'
        [void]$paletteChange.Invoke($null, [object[]]@($excel.VBE.PSObject.BaseObject, $true, [string]$productionRecovery))
        'Applied and reopened verification succeeded.' | Set-Content (Join-Path $OutputDirectory 'production-palette-applied.txt')
    }
    if ($InspectPalette -or $PreviewPalette) {
        $wpf = Join-Path ([Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) 'WPF'
        Add-Type -Path (Join-Path $PSScriptRoot 'VbePaletteInspection.cs') -ReferencedAssemblies (Join-Path $wpf 'UIAutomationClient.dll'),(Join-Path $wpf 'UIAutomationTypes.dll'),(Join-Path $wpf 'WindowsBase.dll'),'System.Web.Extensions.dll','System.Drawing.dll'
        if ($PreviewPalette) { [VbePaletteInspection]::BeginPreview([uint32]$excelProcessId, (Join-Path $OutputDirectory 'palette-before.json')) }
        else { [VbePaletteInspection]::Begin([uint32]$excelProcessId) }
        $optionsCommand = $excel.VBE.CommandBars.FindControl(1, 522)
        $optionsCommand.Execute()
        if (-not [VbePaletteInspection]::Wait()) { throw 'The palette inspection worker did not finish.' }
        [VbePaletteInspection]::Result | Set-Content -LiteralPath (Join-Path $OutputDirectory 'palette-controls.json') -Encoding UTF8
        if (([VbePaletteInspection]::Result | ConvertFrom-Json).Error) { throw [VbePaletteInspection]::Result }
    }
    if ($OpenCodeWindow) {
        $module = $book.VBProject.VBComponents.Add(1)
        $module.Name = 'NativeThemeProbe'
        $module.CodeModule.AddFromString("Option Explicit`r`n`r`nPublic Sub PreviewNativeTheme()`r`n    ' Native syntax colors inspired by Visual Studio Community Dark.`r`n    Dim sampleValue As Long`r`n    sampleValue = 123`r`n    Debug.Print `"Native dark theme`", sampleValue`r`nEnd Sub")
        if ($DebugPreview) {
            $module.CodeModule.DeleteLines(1, $module.CodeModule.CountOfLines)
            $module.CodeModule.AddFromString("Option Explicit`r`nPublic Sub PreviewNativeTheme()`r`n    Dim sampleValue As Long`r`n    sampleValue = 123`r`n    Debug.Print sampleValue`r`n    Stop`r`n    Debug.Print sampleValue + 1`r`nEnd Sub")
        }
        if ($CheckEditorInteractions -or $NativeRenderTraceDll -or $CheckToolbarBlink) {
            if ($DebugPreview) { throw 'Use separate disposable fixtures for editing and debugging.' }
            $module.CodeModule.DeleteLines(1, $module.CodeModule.CountOfLines)
            $longCode = @('Option Explicit', '', 'Public Sub PreviewNativeTheme()')
            $longCode += @(1..120 | ForEach-Object { "    ' Theme scroll line $_" })
            $longCode += @('    Debug.Print "Native theme edit probe"', 'End Sub')
            $module.CodeModule.AddFromString(($longCode -join "`r`n"))
        }
        $module.CodeModule.CodePane.Show()
    }
    if ($CheckPropertyRows -and (-not $NativeRenderTraceDll -or $CheckToolbarBlink)) {
        $propertyForm = $book.VBProject.VBComponents.Add(3)
        $propertyForm.Name = 'NativePropertyRowsProbe'
        $propertyForm.DesignerWindow().Visible = $true
    }
    Start-Sleep -Seconds 3
    $editor = [VbeNativeDarkProbe]::FindEditor([uint32]$excelProcessId)
    if ($editor -eq [IntPtr]::Zero) { throw 'The VBE main window was not found.' }
    if ($NativeRenderTraceDll) {
        $traceSource = [IntPtr]::Zero
        if ($TraceToolbarModule) {
            $traceToolbar = @([VbeNativeDarkProbe]::Snapshot($editor) | Where-Object { $_.Class -eq 'MsoCommandBar' -and $_.Visible -and $_.Caption -eq 'Standard' })
            if ($traceToolbar.Count -ne 1) { throw 'Expected one visible Standard toolbar to identify the trace module.' }
            $traceSource = [IntPtr]$traceToolbar[0].Hwnd
        }
        $renderTrace = Start-VbeRenderTrace -Excel $excel -Workbook $book -Editor $editor -HostProcessId $excelProcessId -LibraryPath $NativeRenderTraceDll -OutputDirectory $OutputDirectory -SourceWindow $traceSource -PatternPilot:$ToolbarPatternPilot
    }
    Export-WindowImage $editor 'before'
    [VbeNativeDarkProbe]::Snapshot($editor) | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'windows.json') -Encoding UTF8
    $results = if ($InProcessAddInExperiment -or $SettingsThemeCycle) { @() } else { [VbeNativeDarkProbe]::Apply($editor) }
    Start-Sleep -Seconds 2
    Export-WindowImage $editor 'after'
    Export-LiveWindowImage $editor 'live'
    if ($CheckToolbarBlink) {
        Add-Type -Path (Join-Path $PSScriptRoot 'VbeToolbarPaintProbe.cs') -ReferencedAssemblies 'System.Drawing.dll','System.Web.Extensions.dll'
        $standard = @([VbeNativeDarkProbe]::Snapshot($editor) | Where-Object { $_.Class -eq 'MsoCommandBar' -and $_.Visible -and $_.Caption -eq 'Standard' })
        if ($standard.Count -ne 1) { throw 'Expected exactly one visible Standard toolbar for the blink observation.' }
        $probeWasTopmost = ([VbeNativeDarkProbe]::GetWindowLongPtr($editor, -20).ToInt64() -band 8) -ne 0
        $probePinned = $false
        $captureStarted = $false
        try {
            if ($KeepToolbarProbeVisible -and -not $probeWasTopmost) {
                $probePinned = [VbeNativeDarkProbe]::SetWindowPos($editor, [IntPtr](-1), 0, 0, 0, 0, 0x13)
                if (-not $probePinned) { throw 'Unable to keep the disposable VBE visible without activation.' }
            }
            [VbeToolbarPaintProbe]::Begin([IntPtr]$standard[0].Hwnd, (Join-Path $OutputDirectory 'toolbar-samples'))
            $captureStarted = $true
            for ($activation = 0; $activation -lt 4; $activation++) {
                $module.CodeModule.CodePane.Show()
                $module.CodeModule.CodePane.SetSelection(4 + $activation, 7, 4 + $activation, 18)
                $module.CodeModule.CodePane.TopLine = $(if ($activation % 2) { 60 } else { 1 })
                Start-Sleep -Milliseconds 300
                $propertyForm.DesignerWindow().Visible = $true
                Start-Sleep -Milliseconds 300
            }
        } finally {
            try {
                if ($captureStarted) {
                    $toolbarReport = [VbeToolbarPaintProbe]::End() | ConvertFrom-Json
                    if (-not $toolbarReport.Completed) { $toolbarReport = [VbeToolbarPaintProbe]::End() | ConvertFrom-Json }
                    $toolbarReport | Add-Member NoteProperty ProbeKeptVisible ([bool]$probePinned)
                    $toolbarReport | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'toolbar-observation.json') -Encoding UTF8
                }
            } finally {
                if ($probePinned -and -not [VbeNativeDarkProbe]::SetWindowPos($editor, [IntPtr](-2), 0, 0, 0, 0, 0x13)) { throw 'Unable to restore the disposable VBE z-order.' }
            }
        }
        if (-not $toolbarReport.Completed) { throw 'The toolbar capture worker did not finish.' }
        if ($toolbarReport.ValidSamples -eq 0) { Write-Warning 'No unobscured toolbar samples: visual blinking remains unverified; continuing the independent repaint counters.' }
        Export-LiveWindowImage $editor 'toolbar-observation-final-live'
        if ($ToolbarOnly) { return }
    }
    if ($NativeRenderTraceDll) {
        $traceCodeBefore = [string]$module.CodeModule.Lines(1, $module.CodeModule.CountOfLines)
        $module.CodeModule.CodePane.Show()
        $module.CodeModule.CodePane.SetSelection(4, 7, 4, 18)
        $startLine = 0; $startColumn = 0; $endLine = 0; $endColumn = 0
        $module.CodeModule.CodePane.GetSelection([ref]$startLine, [ref]$startColumn, [ref]$endLine, [ref]$endColumn)
        if ($startLine -ne 4 -or $startColumn -ne 7 -or $endLine -ne 4 -or $endColumn -ne 18) { throw 'The render-trace code selection was not retained.' }
        Start-Sleep -Milliseconds 250
        Export-LiveWindowImage $editor 'trace-code-selected-live'
        Export-WindowImage $editor 'trace-code-selected-print'
        $module.CodeModule.CodePane.TopLine = 60
        Start-Sleep -Milliseconds 250
        $traceTopLine = [int]$module.CodeModule.CodePane.TopLine
        if ($traceTopLine -le 1) { throw 'The render-trace code viewport did not advance.' }
        Export-LiveWindowImage $editor 'trace-code-scrolled-live'
        Export-WindowImage $editor 'trace-code-scrolled-print'
        $module.CodeModule.CodePane.TopLine = 1
        $module.CodeModule.CodePane.SetSelection(4, 1, 4, 1)
        if ([string]$module.CodeModule.Lines(1, $module.CodeModule.CountOfLines) -cne $traceCodeBefore) { throw 'The trace interaction changed the disposable module text.' }
        @{ SelectionVerified = $true; ScrolledTopLine = $traceTopLine; FinalTopLine = [int]$module.CodeModule.CodePane.TopLine; TextChanged = $false } |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutputDirectory 'trace-code-states.json') -Encoding UTF8
        if ($CheckPropertyRows) {
            # Reuse the earlier toolbar fixture, or create the designer here.
            if (-not $propertyForm) {
                $propertyForm = $book.VBProject.VBComponents.Add(3)
                $propertyForm.Name = 'NativePropertyRowsProbe'
            }
            $propertyForm.DesignerWindow().Visible = $true
        }
        Start-Sleep -Milliseconds 250
    }
    if ($CheckEditorInteractions) {
        if (-not $OpenCodeWindow) { throw 'Editor interaction checks require OpenCodeWindow.' }
        $codeHandle = [IntPtr][long]$module.CodeModule.CodePane.Window.HWnd
        $originalLine = [string]$module.CodeModule.Lines(4, 1)
        $column = $originalLine.Length + 1
        $module.CodeModule.CodePane.SetSelection(4, $column, 4, $column)
        $typedMarker = ' - typed theme check'
        foreach ($character in $typedMarker.ToCharArray()) {
            [void][VbeNativeDarkProbe]::SendMessage($codeHandle, 0x102, [IntPtr][int][char]$character, [IntPtr]1)
        }
        [void][VbeNativeDarkProbe]::SendMessage($codeHandle, 0x102, [IntPtr]13, [IntPtr]1)
        Start-Sleep -Milliseconds 400
        $typedLine = [string]$module.CodeModule.Lines(4, 1)
        if (-not $typedLine.EndsWith($typedMarker)) { throw 'Targeted character messages did not edit the disposable code line.' }
        Export-LiveWindowImage $editor 'editor-typed-live'
        $module.CodeModule.CodePane.SetSelection(4, 7, 4, 18)
        $startLine = 0; $startColumn = 0; $endLine = 0; $endColumn = 0
        $module.CodeModule.CodePane.GetSelection([ref]$startLine, [ref]$startColumn, [ref]$endLine, [ref]$endColumn)
        if ($startLine -ne 4 -or $startColumn -ne 7 -or $endLine -ne 4 -or $endColumn -ne 18) { throw 'The native code selection was not retained.' }
        Start-Sleep -Milliseconds 300
        Export-LiveWindowImage $editor 'editor-selection-live'
        $module.CodeModule.CodePane.TopLine = 60
        Start-Sleep -Milliseconds 300
        $topBefore = $module.CodeModule.CodePane.TopLine
        $vertical = [VbeNativeDarkProbe]::Snapshot($editor) | Where-Object {
            $_.Parent -eq $codeHandle.ToInt64() -and $_.Class -eq 'ScrollBar' -and $_.Visible -and
            ($_.Bounds.Bottom - $_.Bounds.Top) -gt ($_.Bounds.Right - $_.Bounds.Left)
        } | Select-Object -First 1
        if (-not $vertical) { throw 'The native code scrollbar was not found.' }
        [void][VbeNativeDarkProbe]::SendMessage($codeHandle, 0x115, [IntPtr]1, [IntPtr]$vertical.Hwnd)
        Start-Sleep -Milliseconds 300
        $topAfter = $module.CodeModule.CodePane.TopLine
        if ($topAfter -le $topBefore) { throw 'The native code scrollbar did not advance the viewport.' }
        Export-LiveWindowImage $editor 'editor-scroll-live'
        $module.CodeModule.CodePane.TopLine = 1
        Start-Sleep -Milliseconds 300
        Export-LiveWindowImage $editor 'editor-scroll-return-live'
        @{ TypedLineVerified = $true; SelectionVerified = $true; TopBefore = $topBefore; TopAfter = $topAfter } |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutputDirectory 'editor-interactions.json') -Encoding UTF8
    }
    if ($CheckNativeDialogs) {
        Add-Type -Path (Join-Path $PSScriptRoot 'VbePopupThemeProbe.cs') -ReferencedAssemblies 'System.Drawing.dll','System.Core.dll'
        $dialogCommands = @(@{ Name = 'options-dialog'; Control = $excel.VBE.CommandBars.FindControl(1, 522) })
        $propertyCommands = @($excel.VBE.CommandBars.FindControls(1) | Where-Object { $_.Caption -match '^(Propri.t.s de |.*Project Properties)' } | Group-Object Id | ForEach-Object { $_.Group[0] })
        $propertyCommands | Select-Object Id,Caption | ConvertTo-Json | Set-Content (Join-Path $OutputDirectory 'property-commands.json')
        if ($propertyCommands.Count -eq 1 -and $propertyCommands[0].Enabled) {
            $dialogCommands += @{ Name = 'project-properties-dialog'; Control = $propertyCommands[0] }
        }
        foreach ($dialogCommand in $dialogCommands) {
            [VbePopupThemeProbe]::Begin([uint32]$excelProcessId, $editor, (Join-Path $OutputDirectory $dialogCommand.Name), $true)
            $dialogCommand.Control.Execute()
            if (-not [VbePopupThemeProbe]::Wait() -or [VbePopupThemeProbe]::Error) { throw ('Dialog capture failed: ' + [VbePopupThemeProbe]::Error) }
        }
    }
    if ($CheckCommandMenus) {
        Add-Type -Path (Join-Path $PSScriptRoot 'VbePopupThemeProbe.cs') -ReferencedAssemblies 'System.Drawing.dll','System.Core.dll'
        $menuBar = @($excel.VBE.CommandBars | Where-Object { $_.Type -eq 1 }) | Select-Object -First 1
        if (-not $menuBar) { throw 'Native menu bar not found.' }
        $menuIndex = 0
        foreach ($popup in @($menuBar.Controls | Where-Object { $_.Type -eq 10 } | Select-Object -First 3)) {
            if ($popup.CommandBar.Position -ne 5) { throw 'Expected a popup command bar.' }
            $bounds = New-Object VbeNativeDarkProbe+Rect
            [void][VbeNativeDarkProbe]::GetWindowRect($editor, [ref]$bounds)
            foreach ($opening in 1..2) {
                [VbePopupThemeProbe]::BeginMenu([uint32]$excelProcessId, $editor, (Join-Path $OutputDirectory ('menu-' + $menuIndex + '-opening-' + $opening)))
                $popup.accDoDefaultAction(0)
                if (-not [VbePopupThemeProbe]::Wait() -or [VbePopupThemeProbe]::Error) { throw ('Popup capture failed: ' + [VbePopupThemeProbe]::Error) }
            }
            $menuIndex++
        }
        if ($menuIndex -eq 0) { throw 'No native popup controls found.' }
    }
    if ($CheckNativeCombos) {
        $combos = @([VbeNativeDarkProbe]::Snapshot($editor) | Where-Object { $_.Class -eq 'ComboBox' -and $_.Visible })
        $comboResults = @()
        $comboIndex = 0
        foreach ($combo in $combos) {
            $comboHandle = [IntPtr]$combo.Hwnd
            $selectedBefore = [VbeNativeDarkProbe]::SendMessage($comboHandle, 0x147, [IntPtr]::Zero, [IntPtr]::Zero)
            try {
                [void][VbeNativeDarkProbe]::SendMessage($comboHandle, 0x14f, [IntPtr]1, [IntPtr]::Zero)
                Start-Sleep -Milliseconds 350
                $opened = [VbeNativeDarkProbe]::SendMessage($comboHandle, 0x157, [IntPtr]::Zero, [IntPtr]::Zero)
                if (-not $opened) { throw 'The native combo list did not open.' }
                Export-LiveWindowImage $editor ('combo-open-' + $comboIndex)
            } finally {
                [void][VbeNativeDarkProbe]::SendMessage($comboHandle, 0x14f, [IntPtr]::Zero, [IntPtr]::Zero)
            }
            $selectedAfter = [VbeNativeDarkProbe]::SendMessage($comboHandle, 0x147, [IntPtr]::Zero, [IntPtr]::Zero)
            if ($selectedBefore -ne $selectedAfter) { throw 'Opening the native combo changed its selection.' }
            $comboResults += @{ Index = $comboIndex; Opened = [bool]$opened; SelectionPreserved = $true }
            $comboIndex++
        }
        $comboResults | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutputDirectory 'combo-results.json') -Encoding UTF8
    }
    if ($CheckPropertyTabs) {
        $tabWindows = @([VbeNativeDarkProbe]::Snapshot($editor))
        $tabParents = @($tabWindows | Where-Object { $_.Class -eq 'wndclass_pbrs' -and $_.Visible })
        if ($tabParents.Count -ne 1) { throw 'Expected exactly one visible native Properties pane for the tab campaign.' }
        $tabs = @($tabWindows | Where-Object { $_.Class -eq 'SysTabControl32' -and $_.Visible -and $_.Parent -eq $tabParents[0].Hwnd })
        if ($tabs.Count -ne 1) { throw 'Expected exactly one native Properties tab control.' }
        $tabHandle = [IntPtr]$tabs[0].Hwnd
        $tabParentHandle = [IntPtr]$tabParents[0].Hwnd
        $tabResourceHandle = [IntPtr]::Zero
        if ($InProcessAddInExperiment) {
            if ([VbeNativeDarkProbe]::SendMessage($tabHandle, 0x856a, [IntPtr]2, [IntPtr]::Zero) -ne 1) { throw 'The direct Properties tab renderer is not attached.' }
            $tabResourceHandle = $tabHandle
        }
        $count = [VbeNativeDarkProbe]::SendMessage($tabHandle, 0x1304, [IntPtr]::Zero, [IntPtr]::Zero)
        if ($count -ne 2) { throw ('Expected two Properties tabs, found ' + $count) }
        $initial = [VbeNativeDarkProbe]::SendMessage($tabHandle, 0x130b, [IntPtr]::Zero, [IntPtr]::Zero)
        $initialTabVisible = [VbeNativeDarkProbe]::IsWindowVisible($tabHandle)
        $initialTabEnabled = [VbeNativeDarkProbe]::IsWindowEnabled($tabHandle)
        if ($initial -lt 0 -or $initial -ge $count -or -not $initialTabEnabled) { throw 'The initial Properties tab state is not suitable for the campaign.' }
        $tabStates = @()
        $tabResources = @()
        $tabStateRestored = $false
        $idlePaintVerified = $false
        $unrelatedToolbarRefreshSuppressed = $false
        try {
            foreach ($index in 0..1) {
                [void][VbeNativeDarkProbe]::SendMessage($tabHandle, 0x1330, [IntPtr]$index, [IntPtr]::Zero)
                Start-Sleep -Milliseconds 200
                $actual = [VbeNativeDarkProbe]::SendMessage($tabHandle, 0x130b, [IntPtr]::Zero, [IntPtr]::Zero)
                if ($actual -ne $index) { throw 'The native Properties tab did not switch.' }
                $tabStates += Export-PropertyTabState $tabHandle $tabParentHandle ('properties-tab-' + $index)
            }
            Start-Sleep -Milliseconds 500
            $tabResources += Get-ProbeGuiResourceSample 'warmed' $tabResourceHandle
            foreach ($iteration in 0..23) {
                $index = $iteration % 2
                [void][VbeNativeDarkProbe]::SendMessage($tabHandle, 0x1330, [IntPtr]$index, [IntPtr]::Zero)
                Start-Sleep -Milliseconds 40
                if ([VbeNativeDarkProbe]::SendMessage($tabHandle, 0x130b, [IntPtr]::Zero, [IntPtr]::Zero) -ne $index) { throw 'A repeated Properties tab switch failed.' }
            }
            Start-Sleep -Milliseconds 500
            $tabResources += Get-ProbeGuiResourceSample 'after-24-alternations' $tabResourceHandle
            if ($InProcessAddInExperiment -and $LocalRefreshExperiment) {
                $unrelatedToolbarRefreshSuppressed = $tabResources[0].DeferredMsoPassCount -eq $tabResources[1].DeferredMsoPassCount
                if (-not $unrelatedToolbarRefreshSuppressed) { throw 'Properties tab alternations triggered unrelated deferred Office toolbar refreshes.' }
            }
            if (-not [VbeNativeDarkProbe]::RedrawPartial($tabHandle)) { throw 'Partial Properties tab redraw failed.' }
            Start-Sleep -Milliseconds 150
            $tabStates += Export-PropertyTabState $tabHandle $tabParentHandle 'properties-tabs-partial-redraw'
            [void][VbeNativeDarkProbe]::ShowWindow($tabHandle, 0)
            if ([VbeNativeDarkProbe]::IsWindowVisible($tabHandle)) { throw 'The Properties tabs did not hide.' }
            Start-Sleep -Milliseconds 150
            Export-LiveWindowImage $tabParentHandle 'properties-tabs-hidden-pane-live'
            [void][VbeNativeDarkProbe]::ShowWindow($tabHandle, 4)
            if (-not [VbeNativeDarkProbe]::IsWindowVisible($tabHandle)) { throw 'The Properties tabs did not reappear.' }
            Start-Sleep -Milliseconds 200
            $tabStates += Export-PropertyTabState $tabHandle $tabParentHandle 'properties-tabs-shown'
            [void][VbeNativeDarkProbe]::EnableWindow($tabHandle, $false)
            if ([VbeNativeDarkProbe]::IsWindowEnabled($tabHandle)) { throw 'The Properties tabs did not become disabled.' }
            Start-Sleep -Milliseconds 150
            $tabStates += Export-PropertyTabState $tabHandle $tabParentHandle 'properties-tabs-disabled'
            [void][VbeNativeDarkProbe]::EnableWindow($tabHandle, $true)
            if (-not [VbeNativeDarkProbe]::IsWindowEnabled($tabHandle)) { throw 'The Properties tabs did not become enabled again.' }
            [void][VbeNativeDarkProbe]::SendMessage($tabHandle, 0x1330, [IntPtr]$initial, [IntPtr]::Zero)
            Start-Sleep -Milliseconds 200
            $tabStates += Export-PropertyTabState $tabHandle $tabParentHandle 'properties-tabs-restored'
            # Do not capture or send repaint messages between idle samples.
            foreach ($sample in 1..3) {
                Start-Sleep -Milliseconds 750
                $tabResources += Get-ProbeGuiResourceSample ('idle-' + $sample) $tabResourceHandle
            }
            if ($InProcessAddInExperiment) {
                $idleSamples = @($tabResources | Where-Object { $_.Phase -like 'idle-*' })
                $idlePaintVerified = $idleSamples[0].TabPaintCount -eq $idleSamples[1].TabPaintCount -and
                    $idleSamples[1].TabPaintCount -eq $idleSamples[2].TabPaintCount
                if (-not $idlePaintVerified) { throw 'The direct Properties tab renderer kept painting during the two measured idle intervals.' }
            }
        } finally {
            [void][VbeNativeDarkProbe]::EnableWindow($tabHandle, $initialTabEnabled)
            [void][VbeNativeDarkProbe]::ShowWindow($tabHandle, $(if ($initialTabVisible) { 4 } else { 0 }))
            [void][VbeNativeDarkProbe]::SendMessage($tabHandle, 0x1330, [IntPtr]$initial, [IntPtr]::Zero)
            $tabStateRestored = [VbeNativeDarkProbe]::IsWindowVisible($tabHandle) -eq $initialTabVisible -and
                [VbeNativeDarkProbe]::IsWindowEnabled($tabHandle) -eq $initialTabEnabled -and
                [VbeNativeDarkProbe]::SendMessage($tabHandle, 0x130b, [IntPtr]::Zero, [IntPtr]::Zero) -eq $initial
            @{ Count = $count; Initial = $initial; AlternationsRequested = 24; Restored = $tabStateRestored;
                States = $tabStates; Resources = $tabResources;
                IdlePaintActivityVerified = $idlePaintVerified; UnrelatedToolbarRefreshSuppressed = $unrelatedToolbarRefreshSuppressed;
                ResourceScope = 'GDI, USER and CPU: entire disposable Excel process. Tab counters: direct Properties renderer. Mso counters: all Office toolbar chrome passes in the VBE; counters alone do not prove absence of visible blinking.' } |
                ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'properties-tabs.json') -Encoding UTF8
        }
        if (-not $tabStateRestored) { throw 'The initial native Properties tab state was not restored.' }
    }
    if ($CheckPropertyRows) {
        $rowsWindows = @([VbeNativeDarkProbe]::Snapshot($editor))
        $propertiesPane = @($rowsWindows | Where-Object { $_.Class -eq 'wndclass_pbrs' -and $_.Visible })
        if ($propertiesPane.Count -ne 1) { throw 'Expected one native Properties pane.' }
        $rowsLists = @($rowsWindows | Where-Object { $_.Class -eq 'ListBox' -and $_.Visible -and $_.Parent -eq $propertiesPane[0].Hwnd })
        if ($rowsLists.Count -ne 1) { throw 'Expected one Properties list.' }
        $listHandle = [IntPtr]$rowsLists[0].Hwnd
        $parentHandle = [IntPtr]$propertiesPane[0].Hwnd
        $rowCount = [VbeNativeDarkProbe]::SendMessage($listHandle, 0x18b, [IntPtr]::Zero, [IntPtr]::Zero)
        if ($rowCount -lt 10) { throw ('Properties fixture has too few rows: ' + $rowCount) }
        $rowStates = @()
        foreach ($rowIndex in @(0, 3, ($rowCount - 1), 0)) {
            [void][VbeNativeDarkProbe]::SendMessage($listHandle, 0x186, [IntPtr]$rowIndex, [IntPtr]::Zero)
            $controlId = [VbeNativeDarkProbe]::GetDlgCtrlID($listHandle)
            [void][VbeNativeDarkProbe]::SendMessage($parentHandle, 0x111, [IntPtr]($controlId -bor 0x10000), $listHandle)
            [void][VbeNativeDarkProbe]::SendMessage($listHandle, 0x197, [IntPtr]$rowIndex, [IntPtr]::Zero)
            Start-Sleep -Milliseconds 250
            $selected = [VbeNativeDarkProbe]::SendMessage($listHandle, 0x188, [IntPtr]::Zero, [IntPtr]::Zero)
            $top = [VbeNativeDarkProbe]::SendMessage($listHandle, 0x18e, [IntPtr]::Zero, [IntPtr]::Zero)
            if ($selected -ne $rowIndex) { throw 'Properties selection did not update.' }
            Export-WindowImage $parentHandle ('property-rows-' + $rowStates.Count)
            $rowStates += @{ Selected = $selected; Top = $top }
        }
        @{ Count = $rowCount; States = $rowStates } | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $OutputDirectory 'property-row-states.json')
    }
    if ($DebugPreview) {
        if (-not $OpenCodeWindow) { throw 'DebugPreview requires OpenCodeWindow.' }
        $module.CodeModule.CodePane.SetSelection(4, 1, 4, 1)
        $run = $excel.VBE.CommandBars.FindControl(1, 186)
        if (-not $run.Enabled) { throw 'Native run command is unavailable.' }
        $run.Execute()
        Start-Sleep -Seconds 1
        if ($book.VBProject.Mode -ne 1) { throw 'The disposable macro did not enter break mode.' }
        Export-WindowImage $editor 'debug-break'
        Start-Sleep -Milliseconds 250
        Export-LiveWindowImage $editor 'debug-break-live'
        $run.Execute()
        Start-Sleep -Seconds 1
        if ($book.VBProject.Mode -ne 2) { throw 'The disposable macro did not return to design mode.' }
        Export-WindowImage $editor 'debug-finished'
        Start-Sleep -Milliseconds 250
        Export-LiveWindowImage $editor 'debug-finished-live'
    }
    if ($CycleCodeWindow) {
        if (-not $OpenCodeWindow) { throw 'CycleCodeWindow requires OpenCodeWindow.' }
        $module.CodeModule.CodePane.Window.Close()
        $module.CodeModule.CodePane.Show()
        Start-Sleep -Seconds 2
        Export-WindowImage $editor 'reopened'
        [VbeNativeDarkProbe]::Snapshot($editor) | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'windows-reopened.json') -Encoding UTF8
    }
    if ($SettingsThemeCycle) {
        Set-NativeThemeFromSettings $false
        if (Test-Path -LiteralPath $productionRecovery) { throw 'The original palette recovery file remains after verification.' }
        Start-Sleep -Milliseconds 300
        Export-LiveWindowImage $editor 'settings-disabled-live'
        'Settings checkbox disabled; preference saved; original palette verified.' | Set-Content (Join-Path $OutputDirectory 'settings-disabled.txt')
        if ($initialNativeTheme) {
            Set-NativeThemeFromSettings $true
            if ((Get-FileHash -LiteralPath $paletteSnapshot).Hash -ne (Get-FileHash -LiteralPath $productionRecovery).Hash) { throw 'The recreated palette recovery differs from the original.' }
            'Initial enabled preference and original palette recovery restored and verified.' | Set-Content (Join-Path $OutputDirectory 'settings-reenabled.txt')
        }
        $settingsCycleRestored = $true
    }
    $results | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'results.json') -Encoding UTF8
    $results | Group-Object Class | Sort-Object Count -Descending | Select-Object Count, Name
}
catch {
    $_.Exception.ToString() | Set-Content -LiteralPath (Join-Path $OutputDirectory 'probe-error.txt')
    throw
}
finally {
    if ($renderTrace) {
        try { Stop-VbeRenderTrace -Context $renderTrace }
        catch {
            $renderTraceStopError = $_
            $_.Exception.ToString() | Set-Content -LiteralPath (Join-Path $OutputDirectory 'render-trace-stop-error.txt') -Encoding UTF8 -ErrorAction Continue
            Write-Warning $_.Exception.Message
        }
        finally { $renderTrace = $null }
    }
    if ($SettingsThemeCycle -and -not $settingsCycleRestored -and $excel -and $productionRecovery -and
        (($initialNativeTheme -and $initialPalettePath -eq $productionRecovery) -or (-not $initialNativeTheme -and (Test-Path -LiteralPath $productionRecovery)))) {
        try {
            $paletteAssembly = [Reflection.Assembly]::LoadFrom((Join-Path $PSScriptRoot '../../bin/Debug/net48/VBAi.dll'))
            $paletteChange = $paletteAssembly.GetType('VBAi.VbeNativePalette').GetMethod('Change', [Reflection.BindingFlags]'NonPublic,Static')
            [void]$paletteChange.Invoke($null, [object[]]@($excel.VBE.PSObject.BaseObject, $initialNativeTheme, [string]$productionRecovery))
            if ($initialNativeTheme -and (Get-FileHash -LiteralPath $paletteSnapshot).Hash -ne (Get-FileHash -LiteralPath $productionRecovery).Hash) { throw 'The initial palette recovery was not restored exactly; the probe snapshot has been retained.' }
        } catch { $paletteRestoreError = $_; Write-Warning $_.Exception.Message }
    }
    if ($ProductionPalette -and $excel -and $productionRecovery -and (Test-Path -LiteralPath $productionRecovery)) {
        try {
            [void]$paletteChange.Invoke($null, [object[]]@($excel.VBE.PSObject.BaseObject, $false, [string]$productionRecovery))
            'Restored and reopened verification succeeded.' | Set-Content (Join-Path $OutputDirectory 'production-palette-restored.txt')
        } catch { $paletteRestoreError = $_; Write-Warning $_.Exception.Message }
    }
    if ($PreviewPalette -and $excel -and (Test-Path -LiteralPath (Join-Path $OutputDirectory 'palette-before.json'))) {
        try {
            [VbePaletteInspection]::BeginRestore([uint32]$excelProcessId)
            $excel.VBE.CommandBars.FindControl(1, 522).Execute()
            if (-not [VbePaletteInspection]::Wait()) { throw 'The palette restore worker did not finish.' }
            [VbePaletteInspection]::Result | Set-Content -LiteralPath (Join-Path $OutputDirectory 'palette-restore.json') -Encoding UTF8
            if (-not ([VbePaletteInspection]::Result | ConvertFrom-Json).Restored) { throw [VbePaletteInspection]::Result }
            [VbePaletteInspection]::BeginVerifyRestore([uint32]$excelProcessId)
            $excel.VBE.CommandBars.FindControl(1, 522).Execute()
            if (-not [VbePaletteInspection]::Wait()) { throw 'The palette verification worker did not finish.' }
            [VbePaletteInspection]::Result | Set-Content -LiteralPath (Join-Path $OutputDirectory 'palette-restore-verified.json') -Encoding UTF8
            if (-not ([VbePaletteInspection]::Result | ConvertFrom-Json).RestoreVerified) { throw [VbePaletteInspection]::Result }
        } catch { $paletteRestoreError = $_; Write-Warning $_.Exception.Message }
    }
    try { if ($excel) { $excel.VBE.MainWindow.Visible = $false } } catch { Write-Warning $_.Exception.Message }
    try { if ($book) { $book.Close($false) } } catch { Write-Warning $_.Exception.Message }
    $book = $null
    $module = $null
    $propertyForm = $null
    $run = $null
    $optionsCommand = $null
    $menuBar = $null
    $popup = $null
    $propertyCommands = $null
    $dialogCommands = $null
    $dialogCommand = $null
    try { if ($excel) { $excel.Quit() } } catch { Write-Warning $_.Exception.Message }
    $excel = $null
    [Environment]::SetEnvironmentVariable('VBAi_NATIVE_DARK_EXPERIMENT', $previousExperiment, 'Process')
    [Environment]::SetEnvironmentVariable('VBAi_NATIVE_LOCAL_REFRESH_EXPERIMENT', $previousLocalRefresh, 'Process')
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($SettingsThemeCycle -and $settingsBackup -and (Test-Path -LiteralPath $settingsBackup)) {
        [IO.File]::Copy($settingsBackup, $settingsPath, $true)
        if ((Get-FileHash -LiteralPath $settingsBackup).Hash -ne (Get-FileHash -LiteralPath $settingsPath).Hash) { throw 'The original settings file was not restored exactly.' }
        Remove-Item -LiteralPath $settingsBackup
        'Original settings file restored byte-for-byte.' | Set-Content (Join-Path $OutputDirectory 'settings-file-restored.txt')
    }
    if ($excelProcessId) {
        $remaining = Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
        if ($remaining -and -not $remaining.WaitForExit(30000)) {
            throw "Disposable Excel PID=$excelProcessId did not exit; no forced termination was attempted."
        }
    }
    if ($paletteRestoreError) { throw $paletteRestoreError }
    if ($renderTraceStopError) { throw $renderTraceStopError }
}

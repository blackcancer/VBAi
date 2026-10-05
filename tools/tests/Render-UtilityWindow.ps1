param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('About', 'CrashReport', 'Update')]
    [string]$Window,
    [string]$AssemblyPath,
    [string]$Culture = 'fr-FR',
    [ValidateSet('Light', 'Dark')][string]$Theme = 'Light',
    [string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'

$settings = @{
    About = @{ AssemblyPath = 'artifacts/about/build-final/VBAi/Debug/net48/VBAi.dll';
        OutputDirectory = 'artifacts/about/screens'; Type = 'VBAi.AboutWindow'; FilePrefix = 'about' }
    CrashReport = @{ AssemblyPath = 'artifacts/crash-report/build-final/VBAi/Debug/net48/VBAi.dll';
        OutputDirectory = 'artifacts/crash-report/screens'; Type = 'VBAi.CrashReportWindow'; FilePrefix = 'crash-report' }
    Update = @{ AssemblyPath = 'artifacts/updates/build-final/VBAi/Debug/net48/VBAi.dll';
        OutputDirectory = 'artifacts/updates/screens'; Type = 'VBAi.UpdateWindow'; FilePrefix = 'updates' }
}
$selected = $settings[$Window]
if (-not $AssemblyPath) { $AssemblyPath = $selected.AssemblyPath }
if (-not $OutputDirectory) { $OutputDirectory = $selected.OutputDirectory }

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$staticFlags = [Reflection.BindingFlags]'Static,Public,NonPublic'
$uiText = $assembly.GetType('VBAi.UiText')
$selectedCulture = $uiText.GetMethod('Supported', $staticFlags).Invoke($null, @([Globalization.CultureInfo]::GetCultureInfo($Culture)))
$uiText.GetProperty('Culture', $staticFlags).SetValue($null, $selectedCulture)
# Process-local appearance: do not write the user's theme preference.
$uiTheme = $assembly.GetType('VBAi.UiTheme')
$choice = [Enum]::Parse($assembly.GetType('VBAi.ThemeChoice'), $Theme)
$uiTheme.GetField('<Choice>k__BackingField', $staticFlags).SetValue($null, $choice)
if ($Window -eq 'Update') {
    $updatePaths = $assembly.GetType('VBAi.UpdatePaths')
    $updatePaths.GetProperty('Root', $staticFlags).SetValue($null, [IO.Path]::GetFullPath('artifacts/updates/probe-cache'))
}
$windowForm = [Activator]::CreateInstance($assembly.GetType($selected.Type), $true)
try {
    $screen = [Windows.Forms.Screen]::AllScreens | Where-Object { -not $_.Primary } | Select-Object -First 1
    if (-not $screen) { $screen = [Windows.Forms.Screen]::PrimaryScreen }
    $windowForm.StartPosition = [Windows.Forms.FormStartPosition]::Manual
    $windowForm.TopMost = $true
    $windowForm.Location = [Drawing.Point]::new($screen.WorkingArea.Left + [int](($screen.WorkingArea.Width - $windowForm.Width)/2),
        $screen.WorkingArea.Top + [int](($screen.WorkingArea.Height - $windowForm.Height)/2))
    $windowForm.Show(); $windowForm.Activate()
    [Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 250
    [Windows.Forms.Application]::DoEvents()
    $outside = @()
    function Find-ClippedControls($parent) {
        foreach ($child in $parent.Controls) {
            if (-not $child.Visible) { continue }
            if (-not $parent.ClientRectangle.Contains($child.Bounds)) { $script:outside += $child.Name }
            Find-ClippedControls $child
        }
    }
    Find-ClippedControls $windowForm
    $directory = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputDirectory))
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    $bitmap = [Drawing.Bitmap]::new($windowForm.Width, $windowForm.Height)
    try {
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try { $graphics.CopyFromScreen($windowForm.Location, [Drawing.Point]::Empty, $windowForm.Size) }
        finally { $graphics.Dispose() }
        $fileName = "$($selected.FilePrefix)-$Culture-$Theme"
        $path = Join-Path $directory "$fileName.png"
        $bitmap.Save($path, [Drawing.Imaging.ImageFormat]::Png)
        [pscustomobject]@{ Culture = $Culture; Theme = $Theme; Width = $windowForm.Width; Height = $windowForm.Height;
            ClippedControls = $outside; NativeHost = 'Standalone UI probe; no VBE or provider connection' } |
            ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory "$fileName.json") -Encoding UTF8
        if ($outside.Count) { throw "Clipped controls: $($outside -join ', ')" }
        Write-Output $path
    } finally { $bitmap.Dispose() }
} finally { $windowForm.Dispose() }

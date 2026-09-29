param(
    [string]$AssemblyPath = 'artifacts/updates/build-final/VBAi/Debug/net48/VBAi.dll',
    [string]$Culture = 'fr-FR',
    [ValidateSet('Light', 'Dark')][string]$Theme = 'Light',
    [string]$OutputDirectory = 'artifacts/updates/screens'
)
$ErrorActionPreference = 'Stop'
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
$updatePaths = $assembly.GetType('VBAi.UpdatePaths')
$updatePaths.GetProperty('Root', $staticFlags).SetValue($null, [IO.Path]::GetFullPath('artifacts/updates/probe-cache'))
$window = [Activator]::CreateInstance($assembly.GetType('VBAi.UpdateWindow'), $true)
try {
    $screen = [Windows.Forms.Screen]::AllScreens | Where-Object { -not $_.Primary } | Select-Object -First 1
    if (-not $screen) { $screen = [Windows.Forms.Screen]::PrimaryScreen }
    $window.StartPosition = [Windows.Forms.FormStartPosition]::Manual
    $window.TopMost = $true
    $window.Location = [Drawing.Point]::new($screen.WorkingArea.Left + [int](($screen.WorkingArea.Width - $window.Width)/2),
        $screen.WorkingArea.Top + [int](($screen.WorkingArea.Height - $window.Height)/2))
    $window.Show(); $window.Activate()
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
    Find-ClippedControls $window
    $directory = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputDirectory))
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    $bitmap = [Drawing.Bitmap]::new($window.Width, $window.Height)
    try {
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try { $graphics.CopyFromScreen($window.Location, [Drawing.Point]::Empty, $window.Size) }
        finally { $graphics.Dispose() }
        $path = Join-Path $directory "updates-$Culture-$Theme.png"
        $bitmap.Save($path, [Drawing.Imaging.ImageFormat]::Png)
        [pscustomobject]@{ Culture = $Culture; Theme = $Theme; Width = $window.Width; Height = $window.Height; ClippedControls = $outside; NativeHost = 'Standalone UI probe; no VBE or provider connection' } |
            ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory "updates-$Culture-$Theme.json") -Encoding UTF8
        if ($outside.Count) { throw "Clipped controls: $($outside -join ', ')" }
        Write-Output $path
    } finally { $bitmap.Dispose() }
} finally { $window.Dispose() }

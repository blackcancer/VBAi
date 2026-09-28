param([string]$AssemblyPath = 'bin/Debug/net48/CodexVBE.dll', [string]$OutputDirectory = 'artifacts/compact-ui/editor')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms,System.Drawing,System.Web.Extensions
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
$static = [Reflection.BindingFlags]'Static,NonPublic,Public'
$directory = [IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($directory) | Out-Null
$theme = $assembly.GetType('CodexVBE.UiTheme')
$window = [Activator]::CreateInstance($assembly.GetType('CodexVBE.ModernEditorWindow'),$true)
function Wait-Task($task) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while (-not $task.IsCompleted) {
        if ($watch.Elapsed.TotalSeconds -gt 30) { throw 'WebView task timed out' }
        [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 10
    }
    $task.GetAwaiter().GetResult()
}
function Script([string]$source) { Wait-Task ($script:browser.CoreWebView2.ExecuteScriptAsync($source)) }
try {
    $window.Size = [Drawing.Size]::new(900,650)
    $window.StartPosition = [Windows.Forms.FormStartPosition]::Manual
    $window.Location = [Drawing.Point]::new(-2300,-1500)
    $window.ShowInTaskbar = $false
    $window.Show()
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while (-not $window.GetType().GetProperty('Ready',$flags).GetValue($window)) {
        if ($watch.Elapsed.TotalSeconds -gt 40) { throw 'Monaco did not become ready' }
        [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 20
    }
    $script:browser = $window.GetType().GetProperty('Browser',$flags).GetValue($window)
    Script "window.vbai.open('appearance-probe', 'Option Explicit\nSub Example()\nEnd Sub')" | Out-Null
    $results = @()
    foreach ($choice in @('Light','Dark','Light')) {
        $theme.GetField('<Choice>k__BackingField',$static).SetValue($null,[Enum]::Parse($assembly.GetType('CodexVBE.ThemeChoice'),$choice))
        Wait-Task ($window.GetType().GetMethod('Theme',$flags).Invoke($window,@())) | Out-Null
        foreach ($mode in @('code','diff','code-after-diff')) {
            if ($mode -eq 'diff') { Script "window.vbai.compare('Option Explicit')" | Out-Null }
            elseif ($mode -eq 'code-after-diff') { Script 'window.vbai.hideDiff()' | Out-Null }
            $actual = Script "JSON.stringify({ info: window.vbai.testInfo(), background:getComputedStyle(document.body).backgroundColor, surfaces: [...document.querySelectorAll('.monaco-editor')].map(e => ({theme:e.className,background:getComputedStyle(e).backgroundColor})) })" | ConvertFrom-Json | ConvertFrom-Json
            $expected = if ($choice -eq 'Dark') { 'vbai-dark' } else { 'vbai-light' }
            if ($actual.info.themeName -ne $expected) { throw "Wrong Monaco theme: $choice / $mode" }
            $color = $theme.GetProperty('Surface',$static).GetValue($null)
            $rgb = 'rgb({0}, {1}, {2})' -f $color.R,$color.G,$color.B
            if ($actual.background -ne $rgb) { throw "Monaco body differs: $($actual.background) != $rgb" }
            foreach ($surface in $actual.surfaces) { if ($surface.background -ne $rgb) { throw "Editor surface differs: $($surface.background) != $rgb" } }
            $results += "PASS $choice / ${mode}: $expected, $rgb"
            Write-Output $results[-1]
        }
    }
    $results | Set-Content -LiteralPath (Join-Path $directory 'theme-validation.txt') -Encoding UTF8
} finally { $window.Dispose() }


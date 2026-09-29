param([string]$AssemblyPath = 'bin/Debug/net48/VBAi.dll',
    [ValidateSet('Light','Dark')][string]$Theme = 'Dark',
    [string]$OutputDirectory = 'artifacts/compact-ui')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms,System.Drawing,System.Design
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$flags = [Reflection.BindingFlags]'Static,NonPublic'
$themeType = $assembly.GetType('VBAi.UiTheme')
$themeType.GetField('<Choice>k__BackingField',$flags).SetValue($null,[Enum]::Parse($assembly.GetType('VBAi.ThemeChoice'),$Theme))
$directory = [IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($directory) | Out-Null
[ComponentModel.LicenseManager]::CurrentContext = New-Object ComponentModel.Design.DesigntimeLicenseContext
foreach ($name in @('GitWindow','LlmSettingsWindow','AboutWindow','CrashReportWindow','UpdateWindow','UpdateProgressWindow')) {
    [ComponentModel.LicenseManager]::CurrentContext = New-Object ComponentModel.Design.DesigntimeLicenseContext
    $form = [Activator]::CreateInstance($assembly.GetType('VBAi.'+$name),$true)
    try {
        $themeType.GetMethod('Apply',$flags).Invoke($null,@($form))
        $form.ShowInTaskbar = $false
        $form.StartPosition = [Windows.Forms.FormStartPosition]::Manual
        $form.Location = [Drawing.Point]::new(-2400,-1600)
        $form.Show()
        [Windows.Forms.Application]::DoEvents()
        # Runtime palette after OnShown has been suppressed; no providers or documents are opened.
        [ComponentModel.LicenseManager]::CurrentContext = New-Object ComponentModel.LicenseContext
        $themeType.GetMethod('Apply',$flags).Invoke($null,@($form))
        $tabs = @($form.Controls.Find('tabs',$true)) + @($form.Controls.Find('settingsTabs',$true)) | Select-Object -First 1
        $count = if ($tabs) { $tabs.TabCount } else { 1 }
        for ($index=0; $index -lt $count; $index++) {
            if ($tabs) { $tabs.SelectedIndex=$index }
            $form.PerformLayout(); [Windows.Forms.Application]::DoEvents()
            $bitmap = [Drawing.Bitmap]::new($form.Width,$form.Height)
            try {
                $form.DrawToBitmap($bitmap,[Drawing.Rectangle]::new(0,0,$bitmap.Width,$bitmap.Height))
                $path = Join-Path $directory "$name-$Theme-$index.png"
                $bitmap.Save($path,[Drawing.Imaging.ImageFormat]::Png)
                Write-Output $path
            } finally { $bitmap.Dispose() }
        }
    } finally { $form.Dispose() }
}

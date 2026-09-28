param([Parameter(Mandatory=$true)][string]$AssemblyPath)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms,PresentationFramework,PresentationCore,WindowsBase
$assembly=[Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$flags=[Reflection.BindingFlags]'Instance,Static,Public,NonPublic'
$theme=$assembly.GetType('CodexVBE.UiTheme').GetProperty('Choice',$flags)
$oldTheme=$theme.GetValue($null)
$directory=Join-Path (Get-Location) 'artifacts/vbe-completion/ui'
[IO.Directory]::CreateDirectory($directory) | Out-Null
try {
    foreach ($mode in @('Light','Dark')) {
        $theme.SetValue($null,[Enum]::Parse($assembly.GetType('CodexVBE.ThemeChoice'),$mode))
        $chat=[Activator]::CreateInstance($assembly.GetType('CodexVBE.ChatWindow'),$true)
        try {
            $chat.GetType().GetMethod('InitializeShell',$flags).Invoke($chat,@()) | Out-Null
            $chat.GetType().GetMethod('InitializeTranscript',$flags).Invoke($chat,@()) | Out-Null
            $panel=[Windows.Controls.StackPanel]::new(); $panel.Margin=[Windows.Thickness]::new(12)
            foreach ($state in @('Unavailable','Restored','Attempted')) {
                $change=[Activator]::CreateInstance($assembly.GetType('CodexVBE.FormCutChange'),$true)
                $change.Project='TestProject'; $change.Form='UserForm1'; $change.ParentPath='Controls/Frame1'; $change.ControlCount=2
                $change.Restored=$state -eq 'Restored'; $change.Attempted=$state -eq 'Attempted'
                $view=$chat.GetType().GetMethod('RenderFormCut',$flags).Invoke($chat,@($change))
                $button=$view.Child.Children[$view.Child.Children.Count-1]
                if ($button.IsEnabled) { throw "Inactive recovery $state was enabled." }
                if ([string]::IsNullOrWhiteSpace($button.ToolTip)) { throw 'Missing recovery tooltip.' }
                $panel.Children.Add($view) | Out-Null
            }
            $cardHost=[Windows.Controls.Border]::new(); $cardHost.Width=520; $cardHost.Child=$panel
            $cardHost.Background=[Windows.Media.BrushConverter]::new().ConvertFromString($(if ($mode -eq 'Dark') { '#151B26' } else { '#F8FAFC' }))
            $cardHost.Measure([Windows.Size]::new(520,1000)); $height=[Math]::Ceiling($cardHost.DesiredSize.Height)
            $cardHost.Arrange([Windows.Rect]::new(0,0,520,$height)); $cardHost.UpdateLayout()
            $bitmap=[Windows.Media.Imaging.RenderTargetBitmap]::new(520,[int]$height,96,96,[Windows.Media.PixelFormats]::Pbgra32)
            $bitmap.Render($cardHost)
            $encoder=[Windows.Media.Imaging.PngBitmapEncoder]::new(); $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
            $stream=[IO.File]::Create((Join-Path $directory ("form-recovery-"+$mode+".png")))
            try { $encoder.Save($stream) } finally { $stream.Dispose() }
            Write-Output "PASS $mode recovery cards: three disabled historical/completed states with tooltips."
        } finally { $chat.Dispose() }
    }
} finally { $theme.SetValue($null,$oldTheme) }

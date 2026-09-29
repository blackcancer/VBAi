param([string]$AssemblyPath = 'artifacts/multilingual-tests/VBAi.dll')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
Add-Type -TypeDefinition @'
public sealed class MultilingualMenu { public string Caption {get;set;} }
public sealed class MultilingualBar { public int Type = 1; public MultilingualMenu[] Controls {get;set;} }
public sealed class MultilingualVbe { public MultilingualBar[] CommandBars {get;set;} }
'@
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$flags = [Reflection.BindingFlags]'Static,NonPublic,Public'
$fields = [Reflection.BindingFlags]'Instance,NonPublic,Public'
$localization = $assembly.GetType('VBAi.UiText')
$initialize = $localization.GetMethod('Initialize',$flags)
$languages = $assembly.GetType('VBAi.UiLanguages').GetField('All',$flags).GetValue($null)
$output = Join-Path (Get-Location) 'artifacts/localization/multilingual-windows'
New-Item -ItemType Directory -Force $output | Out-Null
try {
    foreach ($language in $languages) {
        $culture = $language.GetType().GetField('CultureName',$fields).GetValue($language)
        $captions = @($language.GetType().GetField('View',$fields).GetValue($language)[0], $language.GetType().GetField('Tools',$fields).GetValue($language)[0])
        $bar = [MultilingualBar]::new()
        $bar.Controls = @($captions | ForEach-Object { $menu = [MultilingualMenu]::new(); $menu.Caption = $_; $menu })
        $vbe = [MultilingualVbe]::new(); $vbe.CommandBars = @($bar)
        $initialize.Invoke($null,@($vbe))
        foreach ($name in @('ChatWindow','LlmSettingsWindow','GitWindow','VbeApprovalDialog')) {
            $form = [Activator]::CreateInstance($assembly.GetType('VBAi.'+$name),$true)
            try {
                $form.Show(); [Windows.Forms.Application]::DoEvents()
                if ($form.RightToLeftLayout -ne ($culture -eq 'ar-SA')) { throw "$culture ${name}: incorrect direction" }
                $bitmap = [Drawing.Bitmap]::new($form.Width,$form.Height)
                try {
                    $form.DrawToBitmap($bitmap,[Drawing.Rectangle]::new(0,0,$form.Width,$form.Height))
                    $bitmap.Save((Join-Path $output "$culture-$name.png"))
                } finally { $bitmap.Dispose() }
            } finally { $form.Dispose() }
        }
        Write-Output "PASS $culture : four windows constructed, displayed and captured"
    }
} finally { $initialize.Invoke($null,@($null)) }

param([string]$AssemblyPath = 'bin/Debug/net48/VBAi.dll')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
Add-Type -TypeDefinition @'
public sealed class LanguageMenu { public string Caption {get;set;} }
public sealed class LanguageBar { public int Type = 1; public LanguageMenu[] Controls {get;set;} }
public sealed class LanguageVbe { public LanguageBar[] CommandBars {get;set;} }
'@
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$type = $assembly.GetType('VBAi.UiText', $true)
$static = [Reflection.BindingFlags]'Static,NonPublic,Public'
$instance = [Reflection.BindingFlags]'Instance,NonPublic,Public'
function Invoke-Text([string]$method, [object[]]$arguments) { $type.GetMethod($method,$static).Invoke($null,$arguments) }
function Assert($condition,[string]$message) { if (-not $condition) { throw $message } }
function Host-Menu([string]$caption) {
    $menu = [LanguageMenu]::new(); $menu.Caption = $caption
    $bar = [LanguageBar]::new(); $bar.Controls = @($menu)
    $hostModel = [LanguageVbe]::new(); $hostModel.CommandBars = @($bar)
    return $hostModel
}
function Field($form,[string]$name) { $form.GetType().GetField($name,$instance).GetValue($form) }
$originalUiCulture = [Threading.Thread]::CurrentThread.CurrentUICulture
$originalCulture = [Threading.Thread]::CurrentThread.CurrentCulture
try {
    $french = [Globalization.CultureInfo]::GetCultureInfo('fr-CA')
    $english = [Globalization.CultureInfo]::GetCultureInfo('en-US')
    Assert ((Invoke-Text Detect @((Host-Menu '&Affichage'),$english)).Name -eq 'fr-FR') 'VBE French must override English Windows.'
    Assert ((Invoke-Text Detect @((Host-Menu '&View'),$french)).Name -eq 'en-US') 'VBE English must override French Windows.'
    Assert ((Invoke-Text Detect @((Host-Menu '&Näytä'),$french)).Name -eq 'en-US') 'Unsupported IDE language must fall back to English.'
    Assert ((Invoke-Text Detect @($null,$french)).Name -eq 'fr-FR') 'Missing IDE must fall back to Windows display language.'
    Assert ((Invoke-Text Detect @($null,[Globalization.CultureInfo]::GetCultureInfo('fi-FI'))).Name -eq 'en-US') 'Unsupported Windows language must fall back to English.'

    $neutral = [xml](Get-Content src/VBAi/Localization/UiStrings.resx -Raw -Encoding UTF8)
    $translated = [xml](Get-Content src/VBAi/Localization/UiStringsFrench.resx -Raw -Encoding UTF8)
    $keys = @($neutral.root.data | ForEach-Object { $_.name })
    $translatedKeys = @($translated.root.data | ForEach-Object { $_.name })
    Assert ((Compare-Object $keys $translatedKeys).Count -eq 0) 'Translation resource keys differ.'
    Assert (($keys | Sort-Object -Unique).Count -eq $keys.Count) 'Duplicate resource keys.'
    Assert (@($translated.root.data | Where-Object { [string]::IsNullOrEmpty($_.value) }).Count -eq 0) 'Empty French translations.'
    New-Item -ItemType Directory -Force artifacts/localization/screenshots | Out-Null
    foreach ($language in @('en','fr')) {
        Invoke-Text Initialize @((Host-Menu $(if ($language -eq 'fr') { 'Outils' } else { 'Tools' })))
        Assert ([Threading.Thread]::CurrentThread.CurrentUICulture.Name -eq $originalUiCulture.Name) 'Host UI culture was changed.'
        Assert ([Threading.Thread]::CurrentThread.CurrentCulture.Name -eq $originalCulture.Name) 'Host number/date culture was changed.'
        Assert ((Invoke-Text Get @('New conversation')) -eq $(if ($language -eq 'fr') {'Nouvelle conversation'} else {'New conversation'})) 'Dynamic translation failed.'
        foreach ($window in @('ChatWindow','LlmSettingsWindow','GitWindow')) {
            $form = [Activator]::CreateInstance($assembly.GetType('VBAi.'+$window),$true)
            try {
                Assert ($form.Icon -ne $null) "Missing icon: $window"
                $fieldName = @{ChatWindow='send'; LlmSettingsWindow='saveButton'; GitWindow='connect'}[$window]
                $expected = if ($language -eq 'fr') { @{ChatWindow='Envoyer ↑';LlmSettingsWindow='Enregistrer';GitWindow='Lier le dépôt'}[$window] } else { @{ChatWindow='Send ↑';LlmSettingsWindow='Save';GitWindow='Link repository'}[$window] }
                Assert ((Field $form $fieldName).Text -eq $expected) "$language $window button translation failed."
                if ($window -eq 'ChatWindow') {
                    Assert ((Field $form configure).Text -eq $(if ($language -eq 'fr') {'Paramètres…'} else {'Settings…'})) 'Chat menu not localized.'
                    Assert ((Field $form toolTips).GetToolTip((Field $form send)) -eq (Invoke-Text Get @('Send the message and its context to the agent.'))) 'Tooltip not localized.'
                }
                $form.Show(); [Windows.Forms.Application]::DoEvents()
                $bitmap = [Drawing.Bitmap]::new($form.Width,$form.Height)
                try { $form.DrawToBitmap($bitmap,[Drawing.Rectangle]::new(0,0,$form.Width,$form.Height)); $bitmap.Save((Join-Path (Get-Location) "artifacts/localization/screenshots/$window-$language.png")) }
                finally { $bitmap.Dispose() }
                Write-Output "PASS $language $window captions, resources and icon"
            } finally { $form.Dispose() }
        }
        $approval = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeApprovalDialog'),$true)
        try {
            Assert ($approval.Text -eq (Invoke-Text Get @('VBAi — approve edit'))) 'Approval dialog title not localized.'
            Assert ((Field $approval approve).Text -eq (Invoke-Text Get @('Allow'))) 'Approval button not localized.'
            Assert ((Field $approval reject).Text -eq (Invoke-Text Get @('Deny'))) 'Deny button not localized.'
            Write-Output "PASS $language approval dialog captions"
        } finally { $approval.Dispose() }
        $sessionType = $assembly.GetType('VBAi.ChatSessionState')
        $session = [Activator]::CreateInstance($sessionType,$true)
        Assert ($session.Title -eq 'Nouvelle conversation') 'Persisted default session token changed.'
        Assert ($session.DisplayTitle -eq (Invoke-Text Get @('New conversation'))) 'Default session title not localized.'
        $session.Title = 'Mon projet / my project'
        Assert ($session.DisplayTitle -eq $session.Title) 'User title was translated.'
        $providerType = $assembly.GetType('VBAi.LlmProvider')
        $providers = $providerType.GetField('All',$static).GetValue($null)
        $custom = @($providers | Where-Object { $_.IsCustom })[0]
        Assert ($custom.Name -eq 'Personnalisé (OpenAI)') 'Custom provider storage key changed.'
        Assert ($custom.ToString() -eq (Invoke-Text Get @('Custom (OpenAI)'))) 'Custom provider display not localized.'
        $commands = $assembly.GetType('VBAi.ChatCommand')
        $expand = $commands.GetMethod('Expand',$static)
        Assert ($expand.Invoke($null,@('/explain example')) -eq $expand.Invoke($null,@('/expliquer example'))) 'English and French command aliases differ.'
    }
    Write-Output "PASS language precedence, English fallback, $($keys.Count) resource pairs, stable session/provider IDs and command aliases"
} finally { Invoke-Text Initialize @($null) }

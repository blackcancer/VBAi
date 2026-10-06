<#
.SYNOPSIS
Explicit standalone UI fixtures. Rendering may show windows on the test desktop.
Chat fixtures capture their exact HWND. UtilityWindow retains screen-capture
and clipping evidence. CompactUi retains design-time off-screen tab renders.
.DESCRIPTION
Select exactly one scenario. No aggregate scenario is provided. Assembly and
output defaults remain specific to each scenario; Help renders its complete ordered
interface bank without activation. Pass an explicit candidate
assembly when required. Shared helper import opens no window or host.
.EXAMPLE
.\Render-Ui.ps1 -Scenario ChatUx -Mode Reasoning -Theme Dark -AssemblyPath artifacts/candidate/VBAi.dll
.EXAMPLE
.\Render-Ui.ps1 -Scenario UtilityWindow -Window About -Culture fr-FR
#>
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('ChatDesignerViews','ChatUx','CompactUi','UtilityWindow','Help')][string]$Scenario,
    [string]$AssemblyPath, [string]$OutputDirectory,
    [ValidateSet('System','Light','Dark')][string]$Theme,
    [ValidateSet('Conversation','History','Reference','Command','Welcome','Queue','Reasoning')][string]$Mode,
    [int]$Width, [int]$Height, [switch]$ScrollToTop,
    [ValidateSet('About','CrashReport','Update')][string]$Window, [string]$Culture
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'UiProbe.psm1') -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'HelpCapture.psm1') -ErrorAction Stop

function Invoke-ChatDesignerViews {
    param(
        [string]$AssemblyPath = 'artifacts/chat-designer/build/VBAi/Debug/net48/VBAi.dll',
        [string]$OutputDirectory = 'artifacts/chat-designer/render',
        [ValidateSet('Light','Dark')][string]$Theme = 'Light',
        [int]$Width = 720
    )
    Add-Type -AssemblyName System.Windows.Forms,System.Drawing,PresentationFramework
    $assembly = Open-UiAssembly $AssemblyPath
    $flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
    $themeType = $assembly.GetType('VBAi.UiTheme')
    $themeType.GetField('<Choice>k__BackingField',[Reflection.BindingFlags]'Static,NonPublic').SetValue($null,[Enum]::Parse($assembly.GetType('VBAi.ThemeChoice'),$Theme))
    $window = New-UiObject $assembly ChatWindow
    $frame = [Windows.Forms.Form]::new()
    $flow = [Windows.Forms.FlowLayoutPanel]::new()
    $flow.FlowDirection = [Windows.Forms.FlowDirection]::TopDown
    $flow.WrapContents = $false; $flow.AutoScroll = $true; $flow.Dock = [Windows.Forms.DockStyle]::Fill
    $frame.Controls.Add($flow); $frame.Size = [Drawing.Size]::new($Width,1000)
    $frame.Text = 'VBAi - Designer cards'
    $themeType.GetMethod('Apply',[Reflection.BindingFlags]'Static,NonPublic').Invoke($null,@($frame))
    $hosts = [Collections.Generic.List[object]]::new()
    try {
        Invoke-UiMethod $window InitializeShell @()
        Invoke-UiMethod $window InitializeComposer @($null)
        Invoke-UiMethod $window InitializeTranscript @()
        $reference = New-UiObject $assembly VbeChatReference
        $reference.Project = 'Budget'; $reference.Module = 'ModuleCalcul'; $reference.Name = 'CalculerTotal'; $reference.Kind = 'Function'
        $user = New-UiObject $assembly ChatEntry; $user.Speaker='Vous'; $user.Text='Explain @Budget.ModuleCalcul.CalculerTotal'
        $assistant = New-UiObject $assembly ChatEntry; $assistant.Speaker='Assistant'; $assistant.Text="## Result`n`nThe **last row** is now included in the total. Review @Budget.ModuleCalcul.CalculerTotal.`n`n- Code synchronized`n- Change available for rollback`n`n[Documentation](https://learn.microsoft.com)"
        $refs = [Array]::CreateInstance($reference.GetType(),1); $refs.SetValue($reference,0); $assistant.References=$refs
        $change = New-UiObject $assembly CodeChange; $change.Project='Budget'; $change.Module='ModuleCalcul'; $change.Before="For i = 1 To lastRow - 1`n    total = total + Cells(i, 2).Value`nNext i"; $change.After="For i = 1 To lastRow`n    total = total + Cells(i, 2).Value`nNext i"
        $changed = New-UiObject $assembly ChatEntry; $changed.Speaker='Code'; $changed.Change=$change
        foreach ($entry in @($user,$assistant,$changed)) {
            Invoke-UiMethod $window AddEntry @($entry)
            $hosted = Invoke-UiMethod $window RenderEntry @($entry)
            $view = $hosted.Child; $hosts.Add($hosted)
            $view.MinimumSize = [Drawing.Size]::Empty
            $view.MaximumSize = [Drawing.Size]::new($Width-50,0); $view.MinimumSize=$view.MaximumSize; $view.Width=$Width-50
            $hosted.Child = $null
            $flow.Controls.Add($view); $view.Visible=$true
            $view.Height = $view.GetPreferredSize([Drawing.Size]::new($Width-50,0)).Height
        }
        $frame.Show(); [Windows.Forms.Application]::DoEvents()
        $frame.Refresh(); [Windows.Forms.Application]::DoEvents()
        $directory = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputDirectory))
        [IO.Directory]::CreateDirectory($directory) | Out-Null
        $path = Join-Path $directory "designer-cards-$Theme-$Width.png"
        Save-UiNativeWindow $frame $path
        Write-Output $path
    } finally { $frame.Dispose(); foreach ($hosted in $hosts) { $hosted.Dispose() }; $window.Dispose() }
}

function Invoke-ChatUx {
    param(
        [ValidateSet('Conversation', 'History', 'Reference', 'Command', 'Welcome', 'Queue', 'Reasoning')][string]$Mode = 'Conversation',
        [int]$Width = 720,
        [int]$Height = 950,
        [switch]$ScrollToTop,
        [ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System',
        [string]$OutputDirectory = 'artifacts/chat-build',
        [string]$AssemblyPath = 'artifacts/chat-build/VBAi/Debug/net48/VBAi.dll'
    )
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing, PresentationFramework
    Add-Type -TypeDefinition @'
public class RenderVbe { public object[] VBProjects { get { return new object[0]; } } }
'@
    $assembly = Open-UiAssembly $AssemblyPath
    $flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
    # Process-local theme only: leave the user's stored preference unchanged.
    $themeType = $assembly.GetType('VBAi.UiTheme')
    $choice = [Enum]::Parse($assembly.GetType('VBAi.ThemeChoice'), $Theme)
    $themeType.GetField('<Choice>k__BackingField', [Reflection.BindingFlags]'Static,NonPublic').SetValue($null, $choice)
    $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession'), $flags, $null, @([RenderVbe]::new()), $null)
    $window = New-UiObject $assembly ChatWindow
    try {
        Invoke-UiMethod $window InitializeShell @()
        Invoke-UiMethod $window InitializeComposer @($session)
        Invoke-UiMethod $window InitializeTranscript @()
        foreach ($pair in @(@('providerPicker','Codex'), @('modelPicker','GPT-6 Sol'), @('effortPicker','Moyen'), @('scopePicker','SuiviBudget - Budget.xlsm'))) {
            $picker = Get-UiField $window $pair[0]
            $picker.Items.Add($pair[1]) | Out-Null
            $picker.SelectedIndex = 0
        }
        Invoke-UiMethod $window RefreshModelSummary @()
        (Get-UiField $window status).Text = 'Prêt'
        (Get-UiField $window effortPicker).Enabled = $true
        (Get-UiField $window sessionTitle).Text = 'Fiabiliser le calcul du total'
        $reference = New-UiObject $assembly VbeChatReference
        $reference.Project = 'SuiviBudget'; $reference.Module = 'ModuleCalcul'; $reference.Name = 'CalculerTotal'; $reference.Kind = 'Function'
        if ($Mode -eq 'Welcome') { Invoke-UiMethod $window ShowWelcome @() }
        elseif ($Mode -eq 'Reasoning') {
            Invoke-UiMethod $window AddTranscriptMessage @('Vous', 'Vérifie le calcul du total et les bornes de la boucle.')
            Invoke-UiMethod $window SetBusy @($true)
            $entry = New-UiObject $assembly ChatEntry
            $entry.Speaker = 'Réflexion'
            $activity = New-UiObject $assembly CodexAgentActivity
            $activity.Kind = 'reasoning'; $activity.Status = 'inProgress'
            $activity.Detail = 'Je vérifie la dernière ligne parcourue et le traitement des cellules vides avant de proposer une correction.'
            $entry.Activity = $activity
            Invoke-UiMethod $window AddEntry @($entry)
        }
        else {
            $entry = New-UiObject $assembly ChatEntry
            $entry.Speaker = 'Vous'; $entry.Text = 'Corrige @SuiviBudget.ModuleCalcul.CalculerTotal'
            $array = [Array]::CreateInstance($reference.GetType(), 1); $array.SetValue($reference, 0); $entry.References = $array
            Invoke-UiMethod $window AddEntry @($entry)
            Invoke-UiMethod $window ReceiveChatUpdate @('summary', 'demo-summary', 'Je parcours la boucle et ses bornes pour identifier les lignes omises dans le total.', $false)
            Invoke-UiMethod $window ReceiveChatUpdate @('summary', 'demo-summary', 'La boucle omet la dernière ligne. Je corrige sa borne supérieure et conserve le reste de la procédure.', $true)
            $change = New-UiObject $assembly CodeChange
            $change.Project = 'SuiviBudget'; $change.Module = 'ModuleCalcul'
            $change.TurnId = 'demo-turn'
            $change.Before = "For i = 1 To derniereLigne - 1`r`n    total = total + Cells(i, 2).Value`r`nNext i"
            $change.After = "For i = 1 To derniereLigne`r`n    total = total + Cells(i, 2).Value`r`nNext i"
            $change.BeforeSha256 = '1234567890'; $change.AfterSha256 = '0987654321'; $change.AfterLineCount = 3
            (Get-UiField $window codeChanges).Add($change)
            Invoke-UiMethod $window AddCodeChangeCard @($change)
            Invoke-UiMethod $window AddTranscriptMessage @('Assistant', 'La dernière ligne est maintenant incluse dans **CalculerTotal**. Le changement est appliqué et reste annulable.')
            (Get-UiField $window changes).Text = 'Modifications · 1'
            (Get-UiField $window changes).Enabled = $true
        }
        if ($Mode -eq 'History') {
            foreach ($title in @('Fiabiliser le calcul du total', 'Expliquer la macro de consolidation', 'Ajouter un export CSV')) {
                $chat = New-UiObject $assembly ChatSessionState; $chat.Title = $title
                (Get-UiField $window scopeSessions).Add($chat)
            }
            Invoke-UiMethod $window RefreshHistory @()
            (Get-UiField $window historyPanel).Visible = $true
            (Get-UiField $window historyPanel).BringToFront()
            (Get-UiField $window chatTitleEditor).Text = 'Fiabiliser le calcul du total'
        }
        if ($Mode -eq 'Queue') {
            $chat = New-UiObject $assembly ChatSessionState
            $window.GetType().GetField('currentSession', $flags).SetValue($window, $chat)
            foreach ($text in @('Ajoute les contrôles de validation dans le module.', 'Explique ensuite le résultat et les changements appliqués.')) {
                $queued = New-UiObject $assembly QueuedChatMessage; $queued.Text = $text; $chat.PendingMessages.Add($queued)
            }
            Invoke-UiMethod $window SetBusy @($true)
            (Get-UiField $window prompt).Text = ''
            Invoke-UiMethod $window RefreshPendingMessages @()
        }
        $screen = [Windows.Forms.Screen]::AllScreens | Where-Object { -not $_.Primary } | Select-Object -First 1
        if (-not $screen) { $screen = [Windows.Forms.Screen]::PrimaryScreen }
        $window.Width = $Width; $window.Height = [Math]::Min($Height, $screen.WorkingArea.Height)
        $window.StartPosition = [Windows.Forms.FormStartPosition]::Manual
        $window.Location = [Drawing.Point]::new($screen.WorkingArea.Left + [int](($screen.WorkingArea.Width - $window.Width)/2),
            $screen.WorkingArea.Top + [int](($screen.WorkingArea.Height - $window.Height)/2))
        $window.TopMost = $true
        $window.Show(); $window.Activate()
        [Windows.Forms.Application]::DoEvents()
        if ($Mode -eq 'History') {
            (Get-UiField $window historyPanel).Visible = $true
            (Get-UiField $window historyPanel).BringToFront()
        }
        if ($Mode -eq 'Reference') {
            $prompt = Get-UiField $window prompt
            $prompt.Text = 'Explique @'
            $prompt.CaretIndex = $prompt.Text.Length
            (Get-UiField $window referenceTimer).Stop()
            (Get-UiField $window referenceIndex).Entries.Add($reference)
            $window.GetType().GetField('referenceIndexReady', $flags).SetValue($window, $true)
            $prompt.Focus() | Out-Null
            Invoke-UiMethod $window UpdateReferences @()
        }
        if ($Mode -eq 'Command') {
            $prompt = Get-UiField $window prompt
            $prompt.Text = '/'; $prompt.CaretIndex = 1; $prompt.Focus() | Out-Null
            Invoke-UiMethod $window UpdateReferences @()
        }
        [Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 300
        [Windows.Forms.Application]::DoEvents()
        if ($ScrollToTop) {
            $window.GetType().GetField('followConversation', $flags).SetValue($window, $false)
            $scroll = Get-UiField $window conversationScroll
            if ($scroll) { $scroll.ScrollToTop() }
            [Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 150
            [Windows.Forms.Application]::DoEvents()
        }
        $directory = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputDirectory))
        [IO.Directory]::CreateDirectory($directory) | Out-Null
        $path = Join-Path $directory "modern-$Mode-$Theme-$Width.png"
        Save-UiNativeWindow $window $path
        if ($Mode -eq 'Queue') {
            Save-UiControlBitmap (Get-UiField $window pendingMessagesPanel) (Join-Path $directory "queue-panel-$Theme-$Width.png")
        }
        Write-Output $path
    }
    finally { $window.Dispose() }
}

function Invoke-CompactUi {
    param([string]$AssemblyPath = 'bin/Debug/net48/VBAi.dll',
        [ValidateSet('Light','Dark')][string]$Theme = 'Dark',
        [string]$OutputDirectory = 'artifacts/compact-ui')
    Add-Type -AssemblyName System.Windows.Forms,System.Drawing,System.Design
    $assembly = Open-UiAssembly $AssemblyPath
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
                $path = Join-Path $directory "$name-$Theme-$index.png"
                Save-UiControlBitmap $form $path
                Write-Output $path
            }
        } finally { $form.Dispose() }
    }
}

function Invoke-UtilityWindow {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('About', 'CrashReport', 'Update')]
        [string]$Window,
        [string]$AssemblyPath,
        [string]$Culture = 'fr-FR',
        [ValidateSet('Light', 'Dark')][string]$Theme = 'Light',
        [string]$OutputDirectory
    )
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
    $assembly = Open-UiAssembly $AssemblyPath
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
        $outside = [Collections.Generic.List[string]]::new()
        function Find-ClippedControls($parent) {
            foreach ($child in $parent.Controls) {
                if (-not $child.Visible) { continue }
                if (-not $parent.ClientRectangle.Contains($child.Bounds)) { $outside.Add($child.Name) }
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
                ClippedControls = @($outside); NativeHost = 'Standalone UI probe; no VBE or provider connection' } |
            ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory "$fileName.json") -Encoding UTF8
            if ($outside.Count) { throw "Clipped controls: $($outside -join ', ')" }
            Write-Output $path
        } finally { $bitmap.Dispose() }
    } finally { $windowForm.Dispose() }
}

$commands = @{
    'ChatDesignerViews' = 'Invoke-ChatDesignerViews'
    'ChatUx' = 'Invoke-ChatUx'
    'CompactUi' = 'Invoke-CompactUi'
    'UtilityWindow' = 'Invoke-UtilityWindow'
    'Help' = 'Invoke-HelpCapture'
}
$command = Get-Command $commands[$Scenario] -CommandType Function
$arguments = Get-UiScenarioParameters $command $PSBoundParameters
& $command @arguments

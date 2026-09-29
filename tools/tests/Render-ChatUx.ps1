param(
    [ValidateSet('Conversation', 'History', 'Reference', 'Command', 'Welcome', 'Queue', 'Reasoning')][string]$Mode = 'Conversation',
    [int]$Width = 720,
    [int]$Height = 950,
    [switch]$ScrollToTop,
    [ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System',
    [string]$OutputDirectory = 'artifacts/chat-build',
    [string]$AssemblyPath = 'artifacts/chat-build/VBAi/Debug/net48/VBAi.dll'
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, PresentationFramework
Add-Type -TypeDefinition @'
public static class RenderWindowCapture {
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool PrintWindow(System.IntPtr handle, System.IntPtr dc, uint flags);
}
public class RenderVbe { public object[] VBProjects { get { return new object[0]; } } }
'@
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
# Process-local theme only: leave the user's stored preference unchanged.
$themeType = $assembly.GetType('VBAi.UiTheme')
$choice = [Enum]::Parse($assembly.GetType('VBAi.ThemeChoice'), $Theme)
$themeType.GetField('<Choice>k__BackingField', [Reflection.BindingFlags]'Static,NonPublic').SetValue($null, $choice)
function New-Internal([string]$Name) { [Activator]::CreateInstance($assembly.GetType("VBAi.$Name"), $true) }
function Field($object, [string]$name) { ,$object.GetType().GetField($name, $flags).GetValue($object) }
function Call($object, [string]$name, [object[]]$arguments) { $object.GetType().GetMethod($name, $flags).Invoke($object, $arguments) }
$session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession'), $flags, $null, @([RenderVbe]::new()), $null)
$window = New-Internal ChatWindow
try {
    Call $window InitializeShell @()
    Call $window InitializeComposer @($session)
    Call $window InitializeTranscript @()
    foreach ($pair in @(@('providerPicker','Codex'), @('modelPicker','GPT-6 Sol'), @('effortPicker','Moyen'), @('scopePicker','SuiviBudget - Budget.xlsm'))) {
        $picker = Field $window $pair[0]
        $picker.Items.Add($pair[1]) | Out-Null
        $picker.SelectedIndex = 0
    }
    Call $window RefreshModelSummary @()
    (Field $window status).Text = 'Prêt'
    (Field $window effortPicker).Enabled = $true
    (Field $window sessionTitle).Text = 'Fiabiliser le calcul du total'
    $reference = New-Internal VbeChatReference
    $reference.Project = 'SuiviBudget'; $reference.Module = 'ModuleCalcul'; $reference.Name = 'CalculerTotal'; $reference.Kind = 'Function'
    if ($Mode -eq 'Welcome') { Call $window ShowWelcome @() }
    elseif ($Mode -eq 'Reasoning') {
        Call $window AddTranscriptMessage @('Vous', 'Vérifie le calcul du total et les bornes de la boucle.')
        Call $window SetBusy @($true)
        $entry = New-Internal ChatEntry
        $entry.Speaker = 'Réflexion'
        $activity = New-Internal CodexAgentActivity
        $activity.Kind = 'reasoning'; $activity.Status = 'inProgress'
        $activity.Detail = 'Je vérifie la dernière ligne parcourue et le traitement des cellules vides avant de proposer une correction.'
        $entry.Activity = $activity
        Call $window AddEntry @($entry)
    }
    else {
        $entry = New-Internal ChatEntry
        $entry.Speaker = 'Vous'; $entry.Text = 'Corrige @SuiviBudget.ModuleCalcul.CalculerTotal'
        $array = [Array]::CreateInstance($reference.GetType(), 1); $array.SetValue($reference, 0); $entry.References = $array
        Call $window AddEntry @($entry)
        Call $window ReceiveChatUpdate @('summary', 'demo-summary', 'Je parcours la boucle et ses bornes pour identifier les lignes omises dans le total.', $false)
        Call $window ReceiveChatUpdate @('summary', 'demo-summary', 'La boucle omet la dernière ligne. Je corrige sa borne supérieure et conserve le reste de la procédure.', $true)
        $change = New-Internal CodeChange
        $change.Project = 'SuiviBudget'; $change.Module = 'ModuleCalcul'
        $change.TurnId = 'demo-turn'
        $change.Before = "For i = 1 To derniereLigne - 1`r`n    total = total + Cells(i, 2).Value`r`nNext i"
        $change.After = "For i = 1 To derniereLigne`r`n    total = total + Cells(i, 2).Value`r`nNext i"
        $change.BeforeSha256 = '1234567890'; $change.AfterSha256 = '0987654321'; $change.AfterLineCount = 3
        (Field $window codeChanges).Add($change)
        Call $window AddCodeChangeCard @($change)
        Call $window AddTranscriptMessage @('Assistant', 'La dernière ligne est maintenant incluse dans **CalculerTotal**. Le changement est appliqué et reste annulable.')
        (Field $window changes).Text = 'Modifications · 1'
        (Field $window changes).Enabled = $true
    }
    if ($Mode -eq 'History') {
        foreach ($title in @('Fiabiliser le calcul du total', 'Expliquer la macro de consolidation', 'Ajouter un export CSV')) {
            $chat = New-Internal ChatSessionState; $chat.Title = $title
            (Field $window scopeSessions).Add($chat)
        }
        Call $window RefreshHistory @()
        (Field $window historyPanel).Visible = $true
        (Field $window historyPanel).BringToFront()
        (Field $window chatTitleEditor).Text = 'Fiabiliser le calcul du total'
    }
    if ($Mode -eq 'Queue') {
        $chat = New-Internal ChatSessionState
        $window.GetType().GetField('currentSession', $flags).SetValue($window, $chat)
        foreach ($text in @('Ajoute les contrôles de validation dans le module.', 'Explique ensuite le résultat et les changements appliqués.')) {
            $queued = New-Internal QueuedChatMessage; $queued.Text = $text; $chat.PendingMessages.Add($queued)
        }
        Call $window SetBusy @($true)
        (Field $window prompt).Text = ''
        Call $window RefreshPendingMessages @()
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
        (Field $window historyPanel).Visible = $true
        (Field $window historyPanel).BringToFront()
    }
    if ($Mode -eq 'Reference') {
        $prompt = Field $window prompt
        $prompt.Text = 'Explique @'
        $prompt.CaretIndex = $prompt.Text.Length
        (Field $window referenceTimer).Stop()
        (Field $window referenceIndex).Entries.Add($reference)
        $window.GetType().GetField('referenceIndexReady', $flags).SetValue($window, $true)
        $prompt.Focus() | Out-Null
        Call $window UpdateReferences @()
    }
    if ($Mode -eq 'Command') {
        $prompt = Field $window prompt
        $prompt.Text = '/'; $prompt.CaretIndex = 1; $prompt.Focus() | Out-Null
        Call $window UpdateReferences @()
    }
    [Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 300
    [Windows.Forms.Application]::DoEvents()
    if ($ScrollToTop) {
        $window.GetType().GetField('followConversation', $flags).SetValue($window, $false)
        $scroll = Field $window conversationScroll
        if ($scroll) { $scroll.ScrollToTop() }
        [Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 150
        [Windows.Forms.Application]::DoEvents()
    }
    $bitmap = [Drawing.Bitmap]::new($window.Width, $window.Height)
    try {
        # Capture the actual window handle, avoiding unrelated foreground windows.
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $dc = $graphics.GetHdc()
            try { if (-not [RenderWindowCapture]::PrintWindow($window.Handle, $dc, 2)) { throw 'PrintWindow failed' } }
            finally { $graphics.ReleaseHdc($dc) }
        } finally { $graphics.Dispose() }
        $directory = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputDirectory))
        [IO.Directory]::CreateDirectory($directory) | Out-Null
        $path = Join-Path $directory "modern-$Mode-$Theme-$Width.png"
        $bitmap.Save($path, [Drawing.Imaging.ImageFormat]::Png)
        if ($Mode -eq 'Queue') {
            $panel = Field $window pendingMessagesPanel
            $queueBitmap = [Drawing.Bitmap]::new($panel.Width, $panel.Height)
            try {
                $panel.DrawToBitmap($queueBitmap, [Drawing.Rectangle]::new(0, 0, $panel.Width, $panel.Height))
                $queueBitmap.Save((Join-Path $directory "queue-panel-$Theme-$Width.png"), [Drawing.Imaging.ImageFormat]::Png)
            } finally { $queueBitmap.Dispose() }
        }
        Write-Output $path
    }
    finally { $bitmap.Dispose() }
}
finally { $window.Dispose() }

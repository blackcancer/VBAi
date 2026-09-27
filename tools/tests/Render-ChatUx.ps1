param(
    [ValidateSet('Conversation', 'History', 'Reference', 'Welcome')][string]$Mode = 'Conversation',
    [int]$Width = 720,
    [int]$Height = 950,
    [string]$AssemblyPath = 'artifacts/chat-build/final/CodexVBE.dll'
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, PresentationFramework
Add-Type -TypeDefinition @'
public class RenderVbe { public object[] VBProjects { get { return new object[0]; } } }
'@
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
function New-Internal([string]$Name) { [Activator]::CreateInstance($assembly.GetType("CodexVBE.$Name"), $true) }
function Field($object, [string]$name) { ,$object.GetType().GetField($name, $flags).GetValue($object) }
function Call($object, [string]$name, [object[]]$arguments) { $object.GetType().GetMethod($name, $flags).Invoke($object, $arguments) }
$session = [Activator]::CreateInstance($assembly.GetType('CodexVBE.VbeSession'), $flags, $null, @([RenderVbe]::new()), $null)
$window = New-Internal ChatWindow
try {
    Call $window InitializeComposer @($session)
    Call $window InitializeTranscript @()
    foreach ($pair in @(@('providerPicker','Codex'), @('modelPicker','GPT-6 Sol'), @('effortPicker','Moyen'), @('scopePicker','SuiviBudget - Budget.xlsm'))) {
        $picker = Field $window $pair[0]
        $picker.Items.Add($pair[1]) | Out-Null
        $picker.SelectedIndex = 0
    }
    (Field $window status).Text = 'Prêt'
    (Field $window effortPicker).IsEnabled = $true
    (Field $window sessionTitle).Text = 'Fiabiliser le calcul du total'
    $reference = New-Internal VbeChatReference
    $reference.Project = 'SuiviBudget'; $reference.Module = 'ModuleCalcul'; $reference.Name = 'CalculerTotal'; $reference.Kind = 'Function'
    if ($Mode -eq 'Welcome') { Call $window ShowWelcome @() }
    else {
        $entry = New-Internal ChatEntry
        $entry.Speaker = 'Vous'; $entry.Text = 'Corrige @SuiviBudget.ModuleCalcul.CalculerTotal'
        $array = [Array]::CreateInstance($reference.GetType(), 1); $array.SetValue($reference, 0); $entry.References = $array
        Call $window AddEntry @($entry)
        Call $window ReceiveChatUpdate @('summary', 'demo-summary', 'Je parcours la boucle et ses bornes pour identifier les lignes omises dans le total.', $false)
        Call $window ReceiveChatUpdate @('summary', 'demo-summary', 'La boucle omet la dernière ligne. Je corrige sa borne supérieure et conserve le reste de la procédure.', $true)
        $change = New-Internal CodeChange
        $change.Project = 'SuiviBudget'; $change.Module = 'ModuleCalcul'
        $change.Before = "For i = 1 To derniereLigne - 1`r`n    total = total + Cells(i, 2).Value`r`nNext i"
        $change.After = "For i = 1 To derniereLigne`r`n    total = total + Cells(i, 2).Value`r`nNext i"
        $change.BeforeSha256 = '1234567890'; $change.AfterSha256 = '0987654321'; $change.AfterLineCount = 3
        (Field $window codeChanges).Add($change)
        Call $window AddCodeChangeCard @($change)
        Call $window AddTranscriptMessage @('Assistant', 'La dernière ligne est maintenant incluse dans **CalculerTotal**. Le changement est appliqué et reste annulable.')
        (Field $window changes).Content = 'Modifications · 1'
        (Field $window changes).IsEnabled = $true
    }
    if ($Mode -eq 'History') {
        foreach ($title in @('Fiabiliser le calcul du total', 'Expliquer la macro de consolidation', 'Ajouter un export CSV')) {
            $chat = New-Internal ChatSessionState; $chat.Title = $title
            (Field $window scopeSessions).Add($chat)
        }
        Call $window RefreshHistory @()
        (Field $window historyPanel).Visibility = [Windows.Visibility]::Visible
        (Field $window chatTitleEditor).Text = 'Fiabiliser le calcul du total'
    }
    $screen = [Windows.Forms.Screen]::AllScreens | Where-Object { -not $_.Primary } | Select-Object -First 1
    if (-not $screen) { $screen = [Windows.Forms.Screen]::PrimaryScreen }
    $window.Width = $Width; $window.Height = [Math]::Min($Height, $screen.WorkingArea.Height)
    $window.StartPosition = [Windows.Forms.FormStartPosition]::Manual
    $window.Location = [Drawing.Point]::new($screen.WorkingArea.Left + [int](($screen.WorkingArea.Width - $window.Width)/2),
        $screen.WorkingArea.Top + [int](($screen.WorkingArea.Height - $window.Height)/2))
    $window.Show(); $window.Activate()
    [Windows.Forms.Application]::DoEvents()
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
    [Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 300
    [Windows.Forms.Application]::DoEvents()
    $bitmap = [Drawing.Bitmap]::new($window.Width, $window.Height)
    try {
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try { $graphics.CopyFromScreen($window.Location, [Drawing.Point]::Empty, $window.Size) }
        finally { $graphics.Dispose() }
        $path = Join-Path (Get-Location) "artifacts/chat-build/modern-$Mode-$Width.png"
        $bitmap.Save($path, [Drawing.Imaging.ImageFormat]::Png)
        Write-Output $path
    }
    finally { $bitmap.Dispose() }
}
finally { $window.Dispose() }

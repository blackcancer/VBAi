param(
    [ValidateSet('Conversation', 'History', 'Reference', 'Command', 'Welcome')][string]$Mode = 'Conversation',
    [int]$Width = 720,
    [int]$Height = 950,
    [string]$AssemblyPath = 'artifacts/chat-build/CodexVBE/Debug/net48/CodexVBE.dll'
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
    Call $window InitializeShell @()
    Call $window InitializeComposer @($session)
    Call $window InitializeTranscript @()
    foreach ($pair in @(@('providerPicker','Codex'), @('modelPicker','GPT-6 Sol'), @('effortPicker','Moyen'), @('scopePicker','SuiviBudget - Budget.xlsm'))) {
        $picker = Field $window $pair[0]
        $picker.Items.Add($pair[1]) | Out-Null
        $picker.SelectedIndex = 0
    }
    (Field $window status).Text = 'Prêt'
    (Field $window effortPicker).Enabled = $true
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
    $screen = [Windows.Forms.Screen]::AllScreens | Where-Object { -not $_.Primary } | Select-Object -First 1
    if (-not $screen) { $screen = [Windows.Forms.Screen]::PrimaryScreen }
    $window.Width = $Width; $window.Height = [Math]::Min($Height, $screen.WorkingArea.Height)
    $window.StartPosition = [Windows.Forms.FormStartPosition]::Manual
    $window.Location = [Drawing.Point]::new($screen.WorkingArea.Left + [int](($screen.WorkingArea.Width - $window.Width)/2),
        $screen.WorkingArea.Top + [int](($screen.WorkingArea.Height - $window.Height)/2))
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
    $bitmap = [Drawing.Bitmap]::new($window.Width, $window.Height)
    try {
        $window.DrawToBitmap($bitmap, [Drawing.Rectangle]::new(0, 0, $window.Width, $window.Height))
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $drawer = Field $window historyPanel
            if ($drawer.Visible) {
                $drawerLocation = $drawer.PointToScreen([Drawing.Point]::Empty)
                $graphics.ExcludeClip([Drawing.Rectangle]::new($drawerLocation.X - $window.Left, $drawerLocation.Y - $window.Top, $drawer.Width, $drawer.Height))
            }
            foreach ($hostName in @('transcriptHost','promptHost')) {
                $hostControl = Field $window $hostName
                $surface = $hostControl.Child
                if (-not $surface -or -not $hostControl.Visible -or $surface.ActualWidth -le 0 -or $surface.ActualHeight -le 0) { continue }
                $target = [Windows.Media.Imaging.RenderTargetBitmap]::new([int][Math]::Ceiling($surface.ActualWidth), [int][Math]::Ceiling($surface.ActualHeight), 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
                $target.Render($surface)
                $encoder = [Windows.Media.Imaging.PngBitmapEncoder]::new()
                $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($target))
                $stream = [IO.MemoryStream]::new()
                try {
                    $encoder.Save($stream); $stream.Position = 0
                    $island = [Drawing.Image]::FromStream($stream)
                    try {
                        $location = $hostControl.PointToScreen([Drawing.Point]::Empty)
                        $graphics.DrawImage($island, [Drawing.Rectangle]::new($location.X - $window.Left, $location.Y - $window.Top, $hostControl.Width, $hostControl.Height))
                    } finally { $island.Dispose() }
                } finally { $stream.Dispose() }
            }
            if ($Mode -eq 'History') {
                $graphics.ResetClip()
                $drawerBitmap = [Drawing.Bitmap]::new($drawer.Width, $drawer.Height)
                try {
                    $drawer.DrawToBitmap($drawerBitmap, [Drawing.Rectangle]::new(0, 0, $drawer.Width, $drawer.Height))
                    $drawerLocation = $drawer.PointToScreen([Drawing.Point]::Empty)
                    $graphics.DrawImageUnscaled($drawerBitmap, $drawerLocation.X - $window.Left, $drawerLocation.Y - $window.Top)
                } finally { $drawerBitmap.Dispose() }
            }
        } finally { $graphics.Dispose() }
        $path = Join-Path (Get-Location) "artifacts/chat-build/modern-$Mode-$Width.png"
        $bitmap.Save($path, [Drawing.Imaging.ImageFormat]::Png)
        Write-Output $path
    }
    finally { $bitmap.Dispose() }
}
finally { $window.Dispose() }

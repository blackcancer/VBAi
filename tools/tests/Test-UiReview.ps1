param([string]$AssemblyPath = 'artifacts/ui-product/CodexVBE.dll')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, PresentationFramework
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$flags = [Reflection.BindingFlags]'Instance,Static,Public,NonPublic'
function Field($target, $name) { ,$target.GetType().GetField($name,$flags).GetValue($target) }
function Call($target,$name,[object[]]$arguments) { $target.GetType().GetMethod($name,$flags).Invoke($target,$arguments) }
function New-Internal($name) { [Activator]::CreateInstance($assembly.GetType('CodexVBE.'+$name),$true) }
function Pump { 1..10 | ForEach-Object { [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 20 } }
function Assert($value,$message) { if (-not $value) { throw $message } }
$output = Join-Path (Get-Location) 'artifacts/ui-review/screenshots'
New-Item -ItemType Directory -Force $output | Out-Null
$chat = New-Internal ChatWindow
try {
    Call $chat InitializeShell @()
    Call $chat InitializeTranscript @()
    for ($i=0;$i -lt 300;$i++) {
        $entry = New-Internal ChatEntry
        $entry.Speaker = 'Assistant'; $entry.Text = "Message $i`n`n**Bold** and _italic_.`n`n| Module | State |`n|---|---|`n| A | Ready |"
        Call $chat AddEntry @($entry)
    }
    Call $chat RefreshTranscriptWindow @([int]220)
    Assert ((Field $chat visibleEntries).Count -eq 81) 'Recent message window was not bounded.'
    $chat.Show(); Pump
    $items = Field $chat conversationItems
    $items.ScrollIntoView((Field $chat visibleEntries)[80]); Pump
    $realized = (Field $chat entryViews).Count
    Assert ($realized -gt 0 -and $realized -lt 30) "Message virtualization failed: $realized views."
    Call $chat ReceiveChatUpdate @('final','review-stream','A streamed answer',$false)
    Pump
    Call $chat ReceiveChatUpdate @('final','review-stream','A **completed** answer',$true)
    Pump
    Assert ((Field $chat transcriptEntries)[300].Text -eq 'A **completed** answer') 'Stream completion lost content.'
    $markdown = $assembly.GetType('CodexVBE.ChatMarkdown').GetMethod('Render',$flags)
    $view = $markdown.Invoke($null,@("# Heading`n`n| A | B |`n|---|---|`n| 1 | 2 |`n`n[Link](https://github.com)",$null,$null,[Action[string]]{param($message) throw $message}))
    Assert (@($view.Document.Blocks | Where-Object { $_ -is [Windows.Documents.Table] }).Count -eq 1) 'Markdown table was not rendered as a table.'
    Write-Output "PASS Markdown table, streaming and virtualized transcript ($realized realized views / 300 messages)"
} finally { $chat.Dispose() }
$theme = $assembly.GetType('CodexVBE.UiTheme')
$choice = $theme.GetProperty('Choice',$flags)
$original = $choice.GetValue($null)
try {
    foreach ($mode in @('Light','Dark')) {
        $choice.SetValue($null,[Enum]::Parse($assembly.GetType('CodexVBE.ThemeChoice'),$mode))
        $git = New-Internal GitWindow
        try {
            $git.Show(); Pump
            $tabs = Field $git tabs
            $tabs.SelectedTab = Field $git githubTab
            $pane = Field $git githubPane
            foreach ($page in @('repositoriesPage','pullsPage')) {
                (Field $pane pages).SelectedTab = Field $pane $page
                $git.Refresh(); Pump
                $bitmap = [Drawing.Bitmap]::new($git.Width,$git.Height)
                try { $git.DrawToBitmap($bitmap,[Drawing.Rectangle]::new(0,0,$git.Width,$git.Height)); $bitmap.Save((Join-Path $output "$mode-$page.png")) } finally { $bitmap.Dispose() }
            }
            $tabs.SelectedTab = Field $git changesTab
            $diff = Field $git diff
            $diff.ShowDiff("Sub Test()`n  value = 1`nEnd Sub", "Sub Test()`n  value = 2`nEnd Sub")
            $git.Refresh(); Pump
            Assert ((Field $diff grid).RowCount -ge 3) 'Diff is empty.'
            $diff.ShowDiff('','')
            Assert ((Field $diff grid).RowCount -eq 0) 'Refreshing an empty diff failed.'
            foreach ($name in @('changesTab','historyTab','branchesTab','checkpointsTab','conflictsTab','importTab')) {
                $page = Field $git $name; $page.Enabled = $true; $tabs.SelectedTab = $page
                if ($name -eq 'conflictsTab') {
                    (Field $git baseContent).Text = "Sub Test()`r`n    value = 1`r`nEnd Sub"
                    (Field $git resolutionText).Text = "Sub Test()`r`n    value = 3`r`nEnd Sub"
                    $conflicts = Field $git conflictDiff
                    $row = $conflicts.Rows.Add('value = 2','value = 3')
                    $conflicts.Rows[$row].Cells[0].Tag = 'vba-removed'; $conflicts.Rows[$row].Cells[1].Tag = 'vba-added'
                }
                Pump
                $bitmap = [Drawing.Bitmap]::new($git.Width,$git.Height)
                try { $git.DrawToBitmap($bitmap,[Drawing.Rectangle]::new(0,0,$git.Width,$git.Height)); $bitmap.Save((Join-Path $output "$mode-$name.png")) } finally { $bitmap.Dispose() }
            }
            $tabs.SelectedTab = Field $git changesTab
            $positions = @((Field $git actions).Controls | ForEach-Object { $_.Bounds.ToString() }) -join ';'
            (Field $git operationProgress).Visible = $true
            (Field $git cancelOperation).Visible = $true
            (Field $git cancelOperation).Enabled = $true
            Pump
            Assert ($positions -eq (@((Field $git actions).Controls | ForEach-Object { $_.Bounds.ToString() }) -join ';')) 'Progress moved the action buttons.'
            $bitmap = [Drawing.Bitmap]::new($git.Width,$git.Height)
            try { $git.DrawToBitmap($bitmap,[Drawing.Rectangle]::new(0,0,$git.Width,$git.Height)); $bitmap.Save((Join-Path $output "$mode-operation.png")) } finally { $bitmap.Dispose() }
            $git.Size = [Drawing.Size]::new(700,580); Pump
            $bitmap = [Drawing.Bitmap]::new($git.Width,$git.Height)
            try { $git.DrawToBitmap($bitmap,[Drawing.Rectangle]::new(0,0,$git.Width,$git.Height)); $bitmap.Save((Join-Path $output "$mode-narrow.png")) } finally { $bitmap.Dispose() }
            Write-Output "PASS $mode GitHub panels and diff refresh"
        } finally { $git.Dispose() }
    }
} finally { $choice.SetValue($null,$original) }

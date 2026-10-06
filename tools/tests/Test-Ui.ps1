<#
.SYNOPSIS
Standalone UI oracles. These scenarios may show windows, render controls or
initialize a WebView. Run on an explicitly selected test desktop. No Office
or SOLIDWORKS host is connected.
.DESCRIPTION
Select exactly one scenario. No aggregate scenario is provided. Assembly and
output defaults remain specific to each scenario; pass an explicit candidate
assembly when required. Shared helper import opens no window or host.
.EXAMPLE
.\Test-Ui.ps1 -Scenario ChatUx -AssemblyPath artifacts/candidate/VBAi.dll
.EXAMPLE
.\Test-Ui.ps1 -Scenario DisplayProfiles -AssemblyPath artifacts/candidate/VBAi.dll -OutputDirectory artifacts/displays
#>
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('ChatUx','EditorAppearance','FormRecoveryCard','MultilingualWindows','UiReview','DisplayProfiles')][string]$Scenario,
    [string]$AssemblyPath, [string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'UiProbe.psm1') -ErrorAction Stop

function Invoke-ChatUx {
    param([string]$AssemblyPath = "artifacts/chat-build/VBAi/Debug/net48/VBAi.dll")
    Add-Type -AssemblyName System.Windows.Forms, PresentationFramework
    Add-Type -TypeDefinition @'
using System;
using System.Linq;
public class ChatTestLines {
    public string this[int start, int count] { get { return string.Join("\r\n", new[] { "Sub Hello()", "End Sub" }.Skip(start - 1).Take(count)); } }
}
public class ChatTestBody { public int this[string name, int kind] { get { return 1; } } }
public class ChatTestPane {
    public int Line, Start, End;
    public void Show() { }
    public void SetSelection(int line, int start, int endLine, int end) { Line = line; Start = start; End = end; }
    public void GetSelection(ref int line, ref int start, ref int endLine, ref int end) { line = endLine = Line; start = Start; end = End; }
}
public class ChatTestModule {
    public int CountOfLines { get { return 2; } }
    public ChatTestLines Lines = new ChatTestLines();
    public ChatTestPane CodePane = new ChatTestPane();
    public ChatTestBody ProcBodyLine = new ChatTestBody();
}
public class ChatTestComponent { public string Name = "Module1"; public int Type = 1; public ChatTestModule CodeModule = new ChatTestModule(); }
public class ChatTestProject {
    public string Name = "Projet"; public string FileName = "C:\\Tests\\Macro.xlsm"; public int Mode = 2;
    public ChatTestComponent[] VBComponents = new[] { new ChatTestComponent() };
}
public class ChatTestVbe { public ChatTestProject[] VBProjects = new[] { new ChatTestProject() }; }
'@
    $assembly = Open-UiAssembly $AssemblyPath
    $flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
    function Call($object, [string]$name, [object[]]$arguments) {
        for ($i = 0; $i -lt $arguments.Length; $i++) {
            if ($null -ne $arguments[$i]) { $arguments[$i] = $arguments[$i].PSObject.BaseObject }
        }
        try { $object.GetType().GetMethod($name, $flags).Invoke($object, $arguments) }
        catch { throw ("$name failed: " + $_.Exception.ToString()) }
    }

    $vbe = [ChatTestVbe]::new()
    $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession'), $flags, $null, @($vbe), $null)
    $index = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeChatReferences'), $flags, $null, @($session), $null)
    $entries = $index.Entries
    for ($i = 0; $i -lt 50; $i++) {
        $module = New-UiObject $assembly VbeChatReference
        $module.Project = 'Projet'; $module.Module = "Module$i"; $module.Kind = 'Module'
        $entries.Add($module)
    }
    $method = New-UiObject $assembly VbeChatReference
    $method.Project = 'Projet'; $method.Module = 'Module1'; $method.Name = 'Hello'; $method.Kind = 'Sub'
    $entries.Add($method)
    $matches = @(Call $index MatchPrefix @('', [char]'@'))
    Assert-Ui ($matches.Count -eq 1 -and $matches[0].Token -eq '@Projet.Module1.Hello') '@ filtering must precede the 40-result limit.'
    $nav = Call $index Navigate @($method)
    Assert-Ui ($nav.Ok -and $vbe.VBProjects[0].VBComponents[0].CodeModule.CodePane.Line -eq 1) 'Procedure navigation failed.'
    Write-Output 'PASS @ function filtering and VBE navigation'

    # Feed real protocol notifications to the client without starting an authenticated process.
    $settings = New-UiObject $assembly LlmSettings
    $tools = [Activator]::CreateInstance($assembly.GetType('VBAi.LlmVbeTools'), $flags, $null, @($session, $null, $settings), $null)
    $ui = [Windows.Forms.WindowsFormsSynchronizationContext]::new()
    $client = [Activator]::CreateInstance($assembly.GetType('VBAi.CodexAppServerClient'), $flags, $null,
        @($ui, $tools, $null, $settings, [string]'test-thread'), $null)
    $events = [Collections.Generic.List[string]]::new()
    $handler = [Action[string,string,string,bool]] { param($kind, $id, $text, $complete) $events.Add("$kind|$id|$text|$complete") }
    $client.GetType().GetEvent('ChatUpdate').AddEventHandler($client, $handler)
    $completion = [Threading.Tasks.TaskCompletionSource[string]]::new()
    Set-UiField $client turnDone $completion
    try {
        Call $client OnLine @('{"method":"item/reasoning/textDelta","params":{"threadId":"test-thread","itemId":"r","delta":"not a summary"}}')
        Call $client OnLine @('{"method":"item/agentMessage/delta","params":{"threadId":"another-thread","itemId":"a","delta":"foreign"}}')
        Call $client OnLine @('{"method":"item/reasoning/summaryPartAdded","params":{"threadId":"test-thread","itemId":"r","summaryIndex":0}}')
        Call $client OnLine @('{"method":"item/reasoning/summaryTextDelta","params":{"threadId":"test-thread","itemId":"r","delta":"summary"}}')
        Call $client OnLine @('{"method":"item/completed","params":{"threadId":"test-thread","item":{"id":"r","type":"reasoning","summary":[{"type":"summary_text","text":"Final summary."}]}}}')
        Call $client OnLine @('{"method":"item/completed","params":{"threadId":"test-thread","item":{"id":"empty","type":"reasoning","summary":[]}}}')
        Call $client OnLine @('{"method":"item/completed","params":{"threadId":"test-thread","item":{"id":"a","type":"agentMessage","phase":"final_answer","text":"final"}}}')
        Call $client OnLine @('{"method":"turn/completed","params":{"threadId":"test-thread","turn":{"status":"completed"}}}')
        [Windows.Forms.Application]::DoEvents()
        Assert-Ui ($events.Count -eq 4 -and $events[0] -eq 'summary|r|summary|False' -and
            $events[1] -eq 'summary|r|Final summary.|True' -and $events[2] -eq 'summary|empty||True' -and
            $events[3] -eq 'final|a|final|True') 'Protocol did not decode the reasoning summary correctly.'
        Assert-Ui ($completion.Task.IsCompleted -and $completion.Task.Result -eq 'final') 'Protocol turn completion failed.'
        Write-Output 'PASS protocol summaries, final answer and foreign-thread isolation'
    }
    finally { $client.Dispose() }

    $directory = Join-Path ([IO.Path]::GetTempPath()) ('VBAi-Chat-' + [guid]::NewGuid().ToString('N'))
    $dbPath = Join-Path $directory 'chat.db'
    $storeType = $assembly.GetType('VBAi.ChatSessionStore')
    $store = [Activator]::CreateInstance($storeType, $flags, $null, @([string]$dbPath), $null)
    $window = New-UiObject $assembly ChatWindow
    try {
        Call $window InitializeShell @()
        Call $window InitializeComposer @($session)
        Call $window InitializeTranscript @()
        Set-UiField $window settings (New-UiObject $assembly LlmSettings)
        Set-UiField $window sessionStore $store
        $providers = $assembly.GetType('VBAi.LlmProvider').GetField('All', [Reflection.BindingFlags]'Public,Static').GetValue($null)
        foreach ($provider in $providers) { (Get-UiField $window providerPicker).Items.Add($provider) | Out-Null }
        $a = New-UiObject $assembly ChatSessionState
        $a.Scope = 'C:\Macro-A.xlsm'; $a.Provider = 'Claude'; $a.Title = 'Macro A'
        $b = New-UiObject $assembly ChatSessionState
        $b.Scope = 'C:\Macro-B.xlsm'; $b.Provider = 'Claude'; $b.Title = 'Macro B'
        Call $window ActivateSession @($a, $false)
        Call $window AddTranscriptMessage @('Vous', 'Analyser @Projet.Module1.Hello')
        Call $window ReceiveChatUpdate @('summary', 'summary-1', 'Verification de la procedure.', $false)
        Call $window ReceiveChatUpdate @('summary', 'summary-1', 'Verification terminee.', $true)
        Call $window ReceiveChatUpdate @('summary', 'empty-summary', $null, $true)
        Call $window ReceiveChatUpdate @('message', 'answer-1', 'Reponse ', $false)
        Call $window ReceiveChatUpdate @('final', 'answer-1', 'Reponse complete.', $true)
        Call $window CompleteAssistantResponse @('Reponse complete.')
        Assert-Ui ((Get-UiField $window transcriptEntries).Count -eq 3) 'Streaming duplicated the final answer.'
        $summaryEntry = @((Get-UiField $window transcriptEntries) | Where-Object { $_.Speaker -eq 'Réflexion' })
        Assert-Ui ($summaryEntry.Count -eq 1 -and $summaryEntry[0].Text -eq 'Verification terminee.') 'The summary was empty or its completed text was lost.'
        $change = New-UiObject $assembly CodeChange
        $change.Project = 'Projet'; $change.Module = 'Module1'
        $change.Before = "Sub Hello()\r\nEnd Sub"; $change.After = "Sub Hello()\r\nDebug.Print 1\r\nEnd Sub"
        $change.BeforeSha256 = '1234567890'; $change.AfterSha256 = '0987654321'; $change.AfterLineCount = 3
        Call $window AddCodeChangeCard @($change)
        (Get-UiField $window selectedReferences).Add($method)
        (Get-UiField $window prompt).Text = 'Brouillon @Projet.Module1.Hello'
        Call $window SaveCurrentSession @()
        $scopeType = $window.GetType().GetNestedType('MacroScope', [Reflection.BindingFlags]'NonPublic')
        $scope = [Activator]::CreateInstance($scopeType, $true)
        $scope.Key = $a.Scope; $scope.Label = 'Macro A'; $scope.Project = 'Projet'
        (Get-UiField $window scopePicker).Items.Add($scope) | Out-Null
        (Get-UiField $window scopePicker).SelectedIndex = 0
        (Get-UiField $window scopeSessions).Add($a)
        Call $window NewSession @($null)
        $secondChat = Get-UiField $window currentSession
        Assert-Ui ($secondChat.Id -ne $a.Id -and $secondChat.Scope -eq $a.Scope) 'New chat was not linked to the same macro.'
        Call $window AddTranscriptMessage @('Vous', 'Deuxieme conversation')
        Call $window SaveCurrentSession @()
        Call $window ActivateSession @($a, $true)
        Assert-Ui ((Get-UiField $window prompt).Text -eq 'Brouillon @Projet.Module1.Hello') 'Switching between chats lost the draft.'
        (Get-UiField $window chatTitleEditor).Text = 'Macro A renommee'
        Call $window RenameCurrentChat @()
        Call $store SaveMemory @($a.Scope, 'Convention locale A')
        Assert-Ui ((Call $store ReadMemory @($b.Scope)) -eq '') 'Project memory leaked across macros.'
        Assert-Ui ((Get-UiField $window attachMemory).Checked -eq $false) 'Memory transmission must be opt-in.'
        Call $window ActivateSession @($b, $true)
        Assert-Ui ((Get-UiField $window transcriptEntries).Count -eq 0) 'Conversation leaked between macro documents.'
        Assert-Ui ((Get-UiField $window prompt).Text -eq '') 'Draft leaked between macro documents.'
        Call $window AddTranscriptMessage @('Vous', 'Question B')
        Call $window SaveCurrentSession @()
        $store.Dispose()
        $store = [Activator]::CreateInstance($storeType, $flags, $null, @([string]$dbPath), $null)
        Set-UiField $window sessionStore $store
        $loadedA = @(Call $store List @($a.Scope))
        $loadedB = @(Call $store List @($b.Scope))
        Assert-Ui ($loadedA.Count -eq 2 -and $loadedB.Count -eq 1) 'SQLite scope isolation or multichat failed.'
        $restoredA = $loadedA | Where-Object { $_.Id -eq $a.Id }
        Assert-Ui ($restoredA.Title -eq 'Macro A renommee') 'Chat rename was not persisted.'
        Assert-Ui ($restoredA.Entries.Count -eq 4) 'SQLite transcript did not survive reopening.'
        Assert-Ui ($restoredA.Entries[3].Change.Before -eq $change.Before) 'Rollback snapshot did not survive reopening.'
        Assert-Ui ($restoredA.DraftReferences[0].Token -eq $method.Token) 'Reference identity did not survive reopening.'
        Assert-Ui ((Call $store ReadMemory @($a.Scope)) -eq 'Convention locale A') 'Project memory did not survive reopening.'
        Call $window ActivateSession @($restoredA, $true)
        Assert-Ui ((Get-UiField $window prompt).Text -eq 'Brouillon @Projet.Module1.Hello') 'Draft was not restored.'
        $window.Show(); [Windows.Forms.Application]::DoEvents()
        (Get-UiField $window conversationItems).UpdateLayout()
        [Windows.Forms.Application]::DoEvents()
        Assert-Ui ((Get-UiField $window rollbackButtons).Count -eq 1) 'Inline rollback card was not restored.'
        Assert-Ui ((Get-UiField $window prompt).SpellCheck.IsEnabled) 'Spell checking regressed.'
        $interrupted = New-UiObject $assembly ChatSessionState
        $interrupted.Scope = $a.Scope; $interrupted.Provider = 'Claude'
        $interrupted.MessagesJson = '[{"role":"user","content":"Fix"},{"role":"assistant","tool_calls":[{"id":"unfinished","function":{"name":"read_module","arguments":"{}"}}]}]'
        Call $window ActivateSession @($interrupted, $true)
        Assert-Ui ((Get-UiField $window messages).Count -eq 2) 'Interrupted tool history was not repaired.'
        Assert-Ui (-not ((Get-UiField $window json).Serialize((Get-UiField $window messages))).Contains('unfinished')) 'Orphaned tool call survived restart.'

        Call $window ActivateSession @($restoredA, $true)
        $restoredA.Pinned = $true
        (Get-UiField $window modePicker).SelectedItem = [Enum]::Parse($assembly.GetType('VBAi.ChatMode'), 'Plan')
        $attachment = New-UiObject $assembly ChatAttachment
        $attachment.Label = 'Captured selection'; $attachment.Text = 'Debug.Print 42'; $attachment.Sha256 = 'snapshot'
        (Get-UiField $window draftAttachments).Add($attachment)
        Call $window SaveCurrentSession @()
        $saved = @(Call $store List @($a.Scope)) | Where-Object { $_.Id -eq $a.Id }
        Assert-Ui ($saved.Pinned -and $saved.Mode.ToString() -eq 'Plan' -and $saved.DraftAttachments[0].Text -eq 'Debug.Print 42') 'Workflow state was not persisted.'
        $last = (Get-UiField $window transcriptEntries)[2]
        Call $window ForkChat @($last)
        $fork = Get-UiField $window currentSession
        Assert-Ui ($fork.Id -ne $a.Id -and $fork.Scope -eq $a.Scope -and -not $fork.CodexThreadId) 'Fork reused a provider thread or crossed a macro.'
        Assert-Ui ($fork.Entries.Count -eq 3 -and $fork.ResumeContext.Contains('Reponse complete.')) 'Fork lost the selected conversation prefix.'
        Assert-Ui (-not $fork.MessagesJson.Contains('unfinished')) 'Fork copied orphaned tool history.'
        Write-Output 'PASS pinned mode/attachment persistence and provider-independent conversation fork'

        Write-Output 'PASS SQLite reopen, document isolation, draft, references, streamed response, inline diff and rollback history'
    }
    finally {
        $window.Dispose()
        if (Test-Path -LiteralPath $dbPath) { Remove-Item -LiteralPath $dbPath }
        if (Test-Path -LiteralPath $directory) { Remove-Item -LiteralPath $directory }
    }
}

function Invoke-EditorAppearance {
    param([string]$AssemblyPath = 'bin/Debug/net48/VBAi.dll', [string]$OutputDirectory = 'artifacts/compact-ui/editor')
    Add-Type -AssemblyName System.Windows.Forms,System.Drawing,System.Web.Extensions
    $assembly = Open-UiAssembly $AssemblyPath
    $flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
    $static = [Reflection.BindingFlags]'Static,NonPublic,Public'
    $directory = [IO.Path]::GetFullPath($OutputDirectory)
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    $theme = $assembly.GetType('VBAi.UiTheme')
    $window = [Activator]::CreateInstance($assembly.GetType('VBAi.ModernEditorWindow'),$true)
    function Wait-Task($task) {
        $watch = [Diagnostics.Stopwatch]::StartNew()
        while (-not $task.IsCompleted) {
            if ($watch.Elapsed.TotalSeconds -gt 30) { throw 'WebView task timed out' }
            [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 10
        }
        $task.GetAwaiter().GetResult()
    }
    function Script([string]$source) { Wait-Task ($browser.CoreWebView2.ExecuteScriptAsync($source)) }
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
        $browser = $window.GetType().GetProperty('Browser',$flags).GetValue($window)
        Script "window.vbai.open('appearance-probe', 'Option Explicit\nSub Example()\nEnd Sub')" | Out-Null
        $results = @()
        foreach ($choice in @('Light','Dark','Light')) {
            $theme.GetField('<Choice>k__BackingField',$static).SetValue($null,[Enum]::Parse($assembly.GetType('VBAi.ThemeChoice'),$choice))
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
}

function Invoke-FormRecoveryCard {
    param([Parameter(Mandatory=$true)][string]$AssemblyPath)
    Add-Type -AssemblyName System.Windows.Forms,PresentationFramework,PresentationCore,WindowsBase
    $assembly=Open-UiAssembly $AssemblyPath
    $flags=[Reflection.BindingFlags]'Instance,Static,Public,NonPublic'
    $theme=$assembly.GetType('VBAi.UiTheme').GetProperty('Choice',$flags)
    $oldTheme=$theme.GetValue($null)
    $directory=Join-Path (Get-Location) 'artifacts/vbe-completion/ui'
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    try {
        foreach ($mode in @('Light','Dark')) {
            $theme.SetValue($null,[Enum]::Parse($assembly.GetType('VBAi.ThemeChoice'),$mode))
            $chat=[Activator]::CreateInstance($assembly.GetType('VBAi.ChatWindow'),$true)
            try {
                $chat.GetType().GetMethod('InitializeShell',$flags).Invoke($chat,@()) | Out-Null
                $chat.GetType().GetMethod('InitializeTranscript',$flags).Invoke($chat,@()) | Out-Null
                $panel=[Windows.Controls.StackPanel]::new(); $panel.Margin=[Windows.Thickness]::new(12)
                foreach ($state in @('Unavailable','Restored','Attempted')) {
                    $change=[Activator]::CreateInstance($assembly.GetType('VBAi.FormCutChange'),$true)
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
}

function Invoke-MultilingualWindows {
    param([string]$AssemblyPath = 'artifacts/multilingual-tests/VBAi.dll')
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing
    [Windows.Forms.Application]::EnableVisualStyles()
    Add-Type -TypeDefinition @'
public sealed class MultilingualMenu { public string Caption {get;set;} }
public sealed class MultilingualBar { public int Type = 1; public MultilingualMenu[] Controls {get;set;} }
public sealed class MultilingualVbe { public MultilingualBar[] CommandBars {get;set;} }
'@
    $assembly = Open-UiAssembly $AssemblyPath
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
                    Save-UiControlBitmap $form (Join-Path $output "$culture-$name.png")
                } finally { $form.Dispose() }
            }
            Write-Output "PASS $culture : four windows constructed, displayed and captured"
        }
    } finally { $initialize.Invoke($null,@($null)) }
}

function Invoke-UiReview {
    param([string]$AssemblyPath = 'artifacts/ui-product/VBAi.dll')
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing, PresentationFramework
    $assembly = Open-UiAssembly $AssemblyPath
    $flags = [Reflection.BindingFlags]'Instance,Static,Public,NonPublic'
    function Pump { 1..10 | ForEach-Object { [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 20 } }
    $output = Join-Path (Get-Location) 'artifacts/ui-review/screenshots'
    New-Item -ItemType Directory -Force $output | Out-Null
    $fixtureRoot = Join-Path (Get-Location) ('artifacts/ui-review/fixtures/' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force $fixtureRoot | Out-Null
    $settingsPath = $assembly.GetType('VBAi.LlmSettings').GetField('StoragePathOverride',$flags)
    $historyPath = $assembly.GetType('VBAi.ChatWindow').GetField('HistoryPath',$flags)
    $themePath = $assembly.GetType('VBAi.UiTheme').GetField('FileName',$flags)
    $originalSettingsPath = $settingsPath.GetValue($null)
    $originalHistoryPath = $historyPath.GetValue($null)
    $originalThemePath = $themePath.GetValue($null)
    try {
        $settingsPath.SetValue($null,(Join-Path $fixtureRoot 'settings.json'))
        $isolatedHistory = Join-Path $fixtureRoot 'chat.db'
        $historyPath.SetValue($null,[Func[string]]({ $isolatedHistory }.GetNewClosure()))
        $themePath.SetValue($null,(Join-Path $fixtureRoot 'theme.txt'))
        $chat = New-UiObject $assembly ChatWindow
        try {
            Invoke-UiMethod $chat InitializeShell @()
            Invoke-UiMethod $chat InitializeTranscript @()
            for ($i=0;$i -lt 300;$i++) {
                $entry = New-UiObject $assembly ChatEntry
                $entry.Speaker = 'Assistant'; $entry.Text = "Message $i`n`n**Bold** and _italic_.`n`n| Module | State |`n|---|---|`n| A | Ready |"
                Invoke-UiMethod $chat AddEntry @($entry)
            }
            Invoke-UiMethod $chat RefreshTranscriptWindow @([int]220)
            Assert-Ui ((Get-UiField $chat visibleEntries).Count -eq 81) 'Recent message window was not bounded.'
            $chat.Show(); Pump
            $items = Get-UiField $chat conversationItems
            $items.ScrollIntoView((Get-UiField $chat visibleEntries)[80]); Pump
            $realized = (Get-UiField $chat entryViews).Count
            Assert-Ui ($realized -gt 0 -and $realized -lt 30) "Message virtualization failed: $realized views."
            Invoke-UiMethod $chat ReceiveChatUpdate @('final','review-stream','A streamed answer',$false)
            Pump
            Invoke-UiMethod $chat ReceiveChatUpdate @('final','review-stream','A **completed** answer',$true)
            Pump
            Assert-Ui ((Get-UiField $chat transcriptEntries)[300].Text -eq 'A **completed** answer') 'Stream completion lost content.'
            $view = New-UiObject $assembly ChatTextContentView
            try {
                $view.CreateControl()
                Invoke-UiMethod $view ShowMarkdown @("# Heading`n`n| A | B |`n|---|---|`n| 1 | 2 |`n`n[Link](https://github.com)",$null,$null,[Action[string]]{param($message) throw $message})
                $content = Get-UiField $view content
                Assert-Ui ($content.ReadOnly) 'Markdown transcript must remain read-only.'
                Assert-Ui ($content.Text.Contains("A`tB`t`n") -and $content.Text.Contains("1`t2`t`n")) 'Markdown table lost its native column or row separators.'
                $content.Select($content.Text.IndexOf("A`tB"),1)
                Assert-Ui ($null -ne $content.SelectionFont -and $content.SelectionFont.Bold) 'Markdown table header lost its bold formatting.'
                $content.Select($content.Text.IndexOf("1`t2"),1)
                Assert-Ui ($null -ne $content.SelectionFont -and -not $content.SelectionFont.Bold) 'Markdown table body incorrectly inherited header formatting.'
                $linkCount = 0
                foreach ($action in (Get-UiField $view actions)) { if ($null -ne (Get-UiField $action Invoke)) { $linkCount++ } }
                Assert-Ui ($linkCount -eq 1) 'Markdown link action was not preserved.'
            } finally { $view.Dispose() }
            Write-Output "PASS Markdown table, streaming and virtualized transcript ($realized realized views / 300 messages)"
        } finally { $chat.Dispose() }
        $theme = $assembly.GetType('VBAi.UiTheme')
        $choice = $theme.GetProperty('Choice',$flags)
        $original = $choice.GetValue($null)
        try {
            foreach ($mode in @('Light','Dark')) {
                $choice.SetValue($null,[Enum]::Parse($assembly.GetType('VBAi.ThemeChoice'),$mode))
                $git = New-UiObject $assembly GitWindow
                try {
                    $git.Show(); Pump
                    $tabs = Get-UiField $git tabs
                    $tabs.SelectedTab = Get-UiField $git githubTab
                    $pane = Get-UiField $git githubPane
                    Assert-Ui ($pane.Controls.Contains((Get-UiField $pane pages))) 'GitHub tabs are detached from their view.'
                    foreach ($page in @('repositoriesPage','pullsPage')) {
                        (Get-UiField $pane pages).SelectedTab = Get-UiField $pane $page
                        $git.Refresh(); Pump
                        Save-UiControlBitmap $git (Join-Path $output "$mode-$page.png")
                    }
                    $tabs.SelectedTab = Get-UiField $git changesTab
                    $diff = Get-UiField $git diff
                    $diff.ShowDiff("Sub Test()`n  value = 1`nEnd Sub", "Sub Test()`n  value = 2`nEnd Sub")
                    $git.Refresh(); Pump
                    Assert-Ui ((Get-UiField $diff grid).RowCount -ge 3) 'Diff is empty.'
                    $diff.ShowDiff('','')
                    Assert-Ui ((Get-UiField $diff grid).RowCount -eq 0) 'Refreshing an empty diff failed.'
                    foreach ($name in @('connectionTab','changesTab','historyTab','branchesTab','checkpointsTab','conflictsTab','importTab')) {
                        $page = Get-UiField $git $name; $page.Enabled = $true; $tabs.SelectedTab = $page
                        if ($name -eq 'conflictsTab') {
                            (Get-UiField $git baseContent).Text = "Sub Test()`r`n    value = 1`r`nEnd Sub"
                            (Get-UiField $git resolutionText).Text = "Sub Test()`r`n    value = 3`r`nEnd Sub"
                            $conflicts = Get-UiField $git conflictDiff
                            $row = $conflicts.Rows.Add('value = 2','value = 3')
                            $conflicts.Rows[$row].Cells[0].Tag = 'vba-removed'; $conflicts.Rows[$row].Cells[1].Tag = 'vba-added'
                        }
                        Pump
                        Save-UiControlBitmap $git (Join-Path $output "$mode-$name.png")
                    }
                    $tabs.SelectedTab = Get-UiField $git changesTab
                    $positions = @((Get-UiField $git actions).Controls | ForEach-Object { $_.Bounds.ToString() }) -join ';'
                    (Get-UiField $git operationProgress).Visible = $true
                    (Get-UiField $git cancelOperation).Visible = $true
                    (Get-UiField $git cancelOperation).Enabled = $true
                    Pump
                    Assert-Ui ($positions -eq (@((Get-UiField $git actions).Controls | ForEach-Object { $_.Bounds.ToString() }) -join ';')) 'Progress moved the action buttons.'
                    Save-UiControlBitmap $git (Join-Path $output "$mode-operation.png")
                    $git.Size = [Drawing.Size]::new(700,580); Pump
                    Save-UiControlBitmap $git (Join-Path $output "$mode-narrow.png")
                    Write-Output "PASS $mode GitHub panels and diff refresh"
                } finally { $git.Dispose() }
                # Populate the real settings controls with synthetic local-provider data.
                # Suppress only first-show account discovery; no credentials or user history are needed for these UI checks.
                $syntheticSettings = New-UiObject $assembly LlmSettings
                $syntheticSettings.ProviderName = 'Ollama'
                $settings = [Activator]::CreateInstance($assembly.GetType('VBAi.LlmSettingsWindow'),$flags,$null,@($syntheticSettings),$null)
                $settings.GetType().GetField('githubLoaded',$flags).SetValue($settings,$true)
                try {
                    $settings.Show(); Pump
                    $settingsTabs = Get-UiField $settings settingsTabs
                    foreach ($page in $settingsTabs.TabPages) {
                        $settingsTabs.SelectedTab = $page; Pump
                        Save-UiControlBitmap $settings (Join-Path $output "$mode-$($page.Name).png")
                    }
                } finally { $settings.Dispose() }

            }
        } finally { $choice.SetValue($null,$original) }
    } finally {
        $settingsPath.SetValue($null,$originalSettingsPath)
        $historyPath.SetValue($null,$originalHistoryPath)
        $themePath.SetValue($null,$originalThemePath)
    }
}

function Invoke-DisplayProfiles {
    param([Parameter(Mandatory=$true)][string]$AssemblyPath,[Parameter(Mandatory=$true)][string]$OutputDirectory)
    Add-Type -AssemblyName System.Windows.Forms,System.Drawing
    [Windows.Forms.Application]::EnableVisualStyles()
    $assembly=Open-UiAssembly $AssemblyPath
    $output=[IO.Path]::GetFullPath($OutputDirectory);[IO.Directory]::CreateDirectory($output)|Out-Null
    $rows=@()
    foreach($screen in [Windows.Forms.Screen]::AllScreens){
        foreach($name in @('ChatWindow','LlmSettingsWindow','GitWindow','VbeApprovalDialog')){
            $form=[Activator]::CreateInstance($assembly.GetType('VBAi.'+$name),$true)
            try{
                $form.StartPosition=[Windows.Forms.FormStartPosition]::Manual
                $form.Location=[Drawing.Point]::new($screen.WorkingArea.Left+20,$screen.WorkingArea.Top+20)
                $form.Show();[Windows.Forms.Application]::DoEvents()
                $actual=[Windows.Forms.Screen]::FromHandle($form.Handle)
                if($actual.DeviceName -ne $screen.DeviceName){throw "$name did not reach $($screen.DeviceName)."}
                $graphics=[Drawing.Graphics]::FromHwnd($form.Handle)
                try{$dpiX=$graphics.DpiX;$dpiY=$graphics.DpiY}finally{$graphics.Dispose()}
                $file=($screen.DeviceName -replace '[^a-zA-Z0-9]','')+'-'+$name+'.png'
                Save-UiControlBitmap $form (Join-Path $output $file)
                $rows+=[pscustomobject]@{Display=$screen.DeviceName;Primary=$screen.Primary;WorkingArea=$screen.WorkingArea.ToString();Window=$name;Bounds=$form.Bounds.ToString();DeviceDpi=$form.DeviceDpi;GraphicsDpiX=$dpiX;GraphicsDpiY=$dpiY;CompletelyInsideWorkingArea=$screen.WorkingArea.Contains($form.Bounds);Screenshot=$file;Qualification='Actual detached WinForms display; not native VBE docking or a simulated DPI profile'}
            }finally{$form.Dispose()}
        }
    }
    $rows|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $output 'display-profiles.json') -Encoding UTF8
    $rows|Format-Table Display,Window,DeviceDpi,CompletelyInsideWorkingArea
}

$commands = @{
    'ChatUx' = 'Invoke-ChatUx'
    'EditorAppearance' = 'Invoke-EditorAppearance'
    'FormRecoveryCard' = 'Invoke-FormRecoveryCard'
    'MultilingualWindows' = 'Invoke-MultilingualWindows'
    'UiReview' = 'Invoke-UiReview'
    'DisplayProfiles' = 'Invoke-DisplayProfiles'
}
$command = Get-Command $commands[$Scenario] -CommandType Function
$arguments = Get-UiScenarioParameters $command $PSBoundParameters
& $command @arguments

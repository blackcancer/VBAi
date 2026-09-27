param([string]$AssemblyPath = "artifacts/chat-build/final/CodexVBE.dll")
$ErrorActionPreference = 'Stop'
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
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
function New-Internal([string]$Name) { [Activator]::CreateInstance($assembly.GetType("CodexVBE.$Name"), $true) }
function Field($object, [string]$name) { ,$object.GetType().GetField($name, $flags).GetValue($object) }
function Set-Field($object, [string]$name, $value) { $object.GetType().GetField($name, $flags).SetValue($object, $value) }
function Call($object, [string]$name, [object[]]$arguments) {
    for ($i = 0; $i -lt $arguments.Length; $i++) {
        if ($null -ne $arguments[$i]) { $arguments[$i] = $arguments[$i].PSObject.BaseObject }
    }
    $object.GetType().GetMethod($name, $flags).Invoke($object, $arguments)
}
function Assert($condition, [string]$message) { if (-not $condition) { throw $message } }

$vbe = [ChatTestVbe]::new()
$session = [Activator]::CreateInstance($assembly.GetType('CodexVBE.VbeSession'), $flags, $null, @($vbe), $null)
$index = [Activator]::CreateInstance($assembly.GetType('CodexVBE.VbeChatReferences'), $flags, $null, @($session), $null)
$entries = $index.Entries
for ($i = 0; $i -lt 50; $i++) {
    $module = New-Internal VbeChatReference
    $module.Project = 'Projet'; $module.Module = "Module$i"; $module.Kind = 'Module'
    $entries.Add($module)
}
$method = New-Internal VbeChatReference
$method.Project = 'Projet'; $method.Module = 'Module1'; $method.Name = 'Hello'; $method.Kind = 'Sub'
$entries.Add($method)
$matches = @(Call $index MatchPrefix @('', [char]'@'))
Assert ($matches.Count -eq 1 -and $matches[0].Token -eq '@Projet.Module1.Hello') '@ filtering must precede the 40-result limit.'
$nav = Call $index Navigate @($method)
Assert ($nav.Ok -and $vbe.VBProjects[0].VBComponents[0].CodeModule.CodePane.Line -eq 1) 'Procedure navigation failed.'
Write-Output 'PASS @ function filtering and VBE navigation'

# Feed real protocol notifications to the client without starting an authenticated process.
$settings = New-Internal LlmSettings
$tools = [Activator]::CreateInstance($assembly.GetType('CodexVBE.LlmVbeTools'), $flags, $null, @($session, $null, $settings), $null)
$ui = [Windows.Forms.WindowsFormsSynchronizationContext]::new()
$client = [Activator]::CreateInstance($assembly.GetType('CodexVBE.CodexAppServerClient'), $flags, $null,
    @($ui, $tools, $null, $settings, [string]'test-thread'), $null)
$events = [Collections.Generic.List[string]]::new()
$handler = [Action[string,string,string,bool]] { param($kind, $id, $text, $complete) $events.Add("$kind|$id|$text|$complete") }
$client.GetType().GetEvent('ChatUpdate').AddEventHandler($client, $handler)
$completion = [Threading.Tasks.TaskCompletionSource[string]]::new()
Set-Field $client turnDone $completion
try {
    Call $client OnLine @('{"method":"item/reasoning/textDelta","params":{"threadId":"test-thread","itemId":"r","delta":"not a summary"}}')
    Call $client OnLine @('{"method":"item/agentMessage/delta","params":{"threadId":"another-thread","itemId":"a","delta":"foreign"}}')
    Call $client OnLine @('{"method":"item/reasoning/summaryTextDelta","params":{"threadId":"test-thread","itemId":"r","delta":"summary"}}')
    Call $client OnLine @('{"method":"item/completed","params":{"threadId":"test-thread","item":{"id":"a","type":"agentMessage","phase":"final_answer","text":"final"}}}')
    Call $client OnLine @('{"method":"turn/completed","params":{"threadId":"test-thread","turn":{"status":"completed"}}}')
    [Windows.Forms.Application]::DoEvents()
    Assert ($events.Count -eq 2 -and $events[0].StartsWith('summary|') -and $events[1].StartsWith('final|')) 'Protocol updates leaked or lost an event.'
    Assert ($completion.Task.IsCompleted -and $completion.Task.Result -eq 'final') 'Protocol turn completion failed.'
    Write-Output 'PASS protocol summaries, final answer and foreign-thread isolation'
}
finally { $client.Dispose() }

$directory = Join-Path ([IO.Path]::GetTempPath()) ('CodexVBE-Chat-' + [guid]::NewGuid().ToString('N'))
$dbPath = Join-Path $directory 'chat.db'
$storeType = $assembly.GetType('CodexVBE.ChatSessionStore')
$store = [Activator]::CreateInstance($storeType, $flags, $null, @([string]$dbPath), $null)
$window = New-Internal ChatWindow
try {
    Call $window InitializeComposer @($session)
    Call $window InitializeTranscript @()
    Set-Field $window settings (New-Internal LlmSettings)
    Set-Field $window sessionStore $store
    $providers = $assembly.GetType('CodexVBE.LlmProvider').GetField('All', [Reflection.BindingFlags]'Public,Static').GetValue($null)
    foreach ($provider in $providers) { (Field $window providerPicker).Items.Add($provider) | Out-Null }
    $a = New-Internal ChatSessionState
    $a.Scope = 'C:\Macro-A.xlsm'; $a.Provider = 'Claude'; $a.Title = 'Macro A'
    $b = New-Internal ChatSessionState
    $b.Scope = 'C:\Macro-B.xlsm'; $b.Provider = 'Claude'; $b.Title = 'Macro B'
    Call $window ActivateSession @($a, $false)
    Call $window AddTranscriptMessage @('Vous', 'Analyser @Projet.Module1.Hello')
    Call $window ReceiveChatUpdate @('summary', 'summary-1', 'Verification de la procedure.', $false)
    Call $window ReceiveChatUpdate @('summary', 'summary-1', 'Verification terminee.', $true)
    Call $window ReceiveChatUpdate @('message', 'answer-1', 'Reponse ', $false)
    Call $window ReceiveChatUpdate @('final', 'answer-1', 'Reponse complete.', $true)
    Call $window CompleteAssistantResponse @('Reponse complete.')
    Assert ((Field $window transcriptEntries).Count -eq 3) 'Streaming duplicated the final answer.'
    $change = New-Internal CodeChange
    $change.Project = 'Projet'; $change.Module = 'Module1'
    $change.Before = "Sub Hello()\r\nEnd Sub"; $change.After = "Sub Hello()\r\nDebug.Print 1\r\nEnd Sub"
    $change.BeforeSha256 = '1234567890'; $change.AfterSha256 = '0987654321'; $change.AfterLineCount = 3
    Call $window AddCodeChangeCard @($change)
    (Field $window selectedReferences).Add($method)
    (Field $window prompt).Text = 'Brouillon @Projet.Module1.Hello'
    Call $window SaveCurrentSession @()
    $scopeType = $window.GetType().GetNestedType('MacroScope', [Reflection.BindingFlags]'NonPublic')
    $scope = [Activator]::CreateInstance($scopeType, $true)
    $scope.Key = $a.Scope; $scope.Label = 'Macro A'; $scope.Project = 'Projet'
    (Field $window scopePicker).Items.Add($scope) | Out-Null
    (Field $window scopePicker).SelectedIndex = 0
    (Field $window scopeSessions).Add($a)
    Call $window NewSession @($null)
    $secondChat = Field $window currentSession
    Assert ($secondChat.Id -ne $a.Id -and $secondChat.Scope -eq $a.Scope) 'New chat was not linked to the same macro.'
    Call $window AddTranscriptMessage @('Vous', 'Deuxieme conversation')
    Call $window SaveCurrentSession @()
    Call $window ActivateSession @($a, $true)
    Assert ((Field $window prompt).Text -eq 'Brouillon @Projet.Module1.Hello') 'Switching between chats lost the draft.'
    (Field $window chatTitleEditor).Text = 'Macro A renommee'
    Call $window RenameCurrentChat @()
    Call $store SaveMemory @($a.Scope, 'Convention locale A')
    Assert ((Call $store ReadMemory @($b.Scope)) -eq '') 'Project memory leaked across macros.'
    Assert ((Field $window attachMemory).IsChecked -eq $false) 'Memory transmission must be opt-in.'
    Call $window ActivateSession @($b, $true)
    Assert ((Field $window transcriptEntries).Count -eq 0) 'Conversation leaked between macro documents.'
    Assert ((Field $window prompt).Text -eq '') 'Draft leaked between macro documents.'
    Call $window AddTranscriptMessage @('Vous', 'Question B')
    Call $window SaveCurrentSession @()
    $store.Dispose()
    $store = [Activator]::CreateInstance($storeType, $flags, $null, @([string]$dbPath), $null)
    Set-Field $window sessionStore $store
    $loadedA = @(Call $store List @($a.Scope))
    $loadedB = @(Call $store List @($b.Scope))
    Assert ($loadedA.Count -eq 2 -and $loadedB.Count -eq 1) 'SQLite scope isolation or multichat failed.'
    $restoredA = $loadedA | Where-Object { $_.Id -eq $a.Id }
    Assert ($restoredA.Title -eq 'Macro A renommee') 'Chat rename was not persisted.'
    Assert ($restoredA.Entries.Count -eq 4) 'SQLite transcript did not survive reopening.'
    Assert ($restoredA.Entries[3].Change.Before -eq $change.Before) 'Rollback snapshot did not survive reopening.'
    Assert ($restoredA.DraftReferences[0].Token -eq $method.Token) 'Reference identity did not survive reopening.'
    Assert ((Call $store ReadMemory @($a.Scope)) -eq 'Convention locale A') 'Project memory did not survive reopening.'
    Call $window ActivateSession @($restoredA, $true)
    Assert ((Field $window prompt).Text -eq 'Brouillon @Projet.Module1.Hello') 'Draft was not restored.'
    Assert ((Field $window rollbackButtons).Count -eq 1) 'Inline rollback card was not restored.'
    Assert ((Field $window prompt).SpellCheck.IsEnabled) 'Spell checking regressed.'
    $interrupted = New-Internal ChatSessionState
    $interrupted.Scope = $a.Scope; $interrupted.Provider = 'Claude'
    $interrupted.MessagesJson = '[{"role":"user","content":"Fix"},{"role":"assistant","tool_calls":[{"id":"unfinished","function":{"name":"read_module","arguments":"{}"}}]}]'
    Call $window ActivateSession @($interrupted, $true)
    Assert ((Field $window messages).Count -eq 2) 'Interrupted tool history was not repaired.'
    Assert (-not ((Field $window json).Serialize((Field $window messages))).Contains('unfinished')) 'Orphaned tool call survived restart.'
    Write-Output 'PASS SQLite reopen, document isolation, draft, references, streamed response, inline diff and rollback history'
}
finally {
    $window.Dispose()
    if (Test-Path -LiteralPath $dbPath) { Remove-Item -LiteralPath $dbPath }
    if (Test-Path -LiteralPath $directory) { Remove-Item -LiteralPath $directory }
}

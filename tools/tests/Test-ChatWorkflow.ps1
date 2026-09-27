$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Linq;
public class FakeCodeLines {
    private readonly FakeCodeModule module;
    public FakeCodeLines(FakeCodeModule module) { this.module = module; }
    public string this[int start, int count] { get { return string.Join("\r\n", module.Source.Skip(start - 1).Take(count)); } }
}
public class FakeCodeModule {
    public List<string> Source = new List<string> { "Sub Hello()", "    Debug.Print 1", "End Sub" };
    public int CountOfLines { get { return Source.Count; } }
    public FakeCodeLines Lines { get { return new FakeCodeLines(this); } }
    public void DeleteLines(int start, int count) { Source.RemoveRange(start - 1, count); }
    public void InsertLines(int start, string text) { Source.InsertRange(start - 1,
        text.Replace("\r\n", "\n").Replace('\r', '\n').Split(new[]{'\n'}, StringSplitOptions.RemoveEmptyEntries)); }
}
public class FakeComponent { public string Name { get { return "Module1"; } } public int Type { get { return 1; } } public FakeCodeModule CodeModule = new FakeCodeModule(); }
public class FakeProject { public string Name { get { return "Projet"; } } public string FileName { get { return "probe.xlsm"; } } public int Mode { get { return 2; } } public FakeComponent[] VBComponents = new[] { new FakeComponent() }; }
public class FakeVbe { public FakeProject[] VBProjects = new[] { new FakeProject() }; }
'@
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path 'artifacts\chat-build\final\CodexVBE.dll'))
$flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
$vbe = [FakeVbe]::new()
$sessionType = $assembly.GetType('CodexVBE.VbeSession')
$session = [Activator]::CreateInstance($sessionType, $flags, $null, @($vbe), $null)
$requestType = $assembly.GetType('CodexVBE.Request')
$read = [Activator]::CreateInstance($requestType)
$read.Command = 'read_module'; $read.Project = 'Projet'; $read.Module = 'Module1'
$before = $sessionType.GetMethod('Execute').Invoke($session, @($read))
if (-not $before.Ok) { throw $before.Error }
$settingsType = $assembly.GetType('CodexVBE.LlmSettings')
$settings = [Activator]::CreateInstance($settingsType, $true)
$settings.VbeEditApproval = 'AskEachTime'
$toolsType = $assembly.GetType('CodexVBE.LlmVbeTools')
$tools = [Activator]::CreateInstance($toolsType, $flags, $null, @($session, $null, $settings), $null)
$requestJson = [string](@{ Project = 'Projet'; Module = 'Module1'; ExpectedSha256 = $before.Data.Sha256; StartLine = 2; Count = 1; Text = '    Debug.Print 2' } | ConvertTo-Json -Compress)
$result = $toolsType.GetMethod('Invoke').Invoke($tools, @([string]'replace_lines', [string]$requestJson)) | ConvertFrom-Json
if (-not $result.Ok -or $vbe.VBProjects[0].VBComponents[0].CodeModule.Source[1] -ne '    Debug.Print 2') { throw 'replace_lines failed.' }
$changeType = $assembly.GetType('CodexVBE.CodeChange')
$afterText = "Sub Hello()`r`n    Debug.Print 2`r`nEnd Sub"
$change = [Activator]::CreateInstance($changeType, $flags, $null, @([string]'Projet', [string]'Module1', [string]$before.Data.Code, [string]$before.Data.Sha256, [string]$afterText, [string]$result.Data.Sha256, [int]3), $null)
$rows = $changeType.GetProperty('Rows').GetValue($change)
if (@($rows | Where-Object { $_.Kind.ToString() -eq 'Removed' }).Count -ne 1 -or @($rows | Where-Object { $_.Kind.ToString() -eq 'Added' }).Count -ne 1) { throw 'Structured diff failed.' }
$vbe.VBProjects[0].VBComponents[0].CodeModule.Source[1] = '    Debug.Print 3'
$stale = $toolsType.GetMethod('RestoreCodeChange').Invoke($tools, @($change))
if ($stale.Ok -or $vbe.VBProjects[0].VBComponents[0].CodeModule.Source[1] -ne '    Debug.Print 3') { throw 'Stale restore was not refused.' }
$vbe.VBProjects[0].VBComponents[0].CodeModule.Source[1] = '    Debug.Print 2'
$restore = $toolsType.GetMethod('RestoreCodeChange').Invoke($tools, @($change))
if (-not $restore.Ok -or $vbe.VBProjects[0].VBComponents[0].CodeModule.Source[1] -ne '    Debug.Print 1') { throw 'Safe restore failed.' }
Write-Output 'PASS edit without approval in AskEachTime, structured diff, stale refusal, restore'

function Assert($condition, [string]$message) { if (-not $condition) { throw $message } }
function New-Internal([string]$name) { [Activator]::CreateInstance($assembly.GetType("CodexVBE.$name"), $true) }
function Static([string]$type, [string]$method, [object[]]$arguments) { $assembly.GetType("CodexVBE.$type").GetMethod($method).Invoke($null, $arguments) }
$modeType = $assembly.GetType('CodexVBE.ChatMode')
foreach ($mode in @('Discussion', 'Plan')) {
    $tools.Mode = [Enum]::Parse($modeType, $mode)
    $denied = $toolsType.GetMethod('Invoke').Invoke($tools, @('replace_lines', $requestJson)) | ConvertFrom-Json
    Assert (-not $denied.Ok) "$mode allowed a code mutation."
    foreach ($tool in @('immediate_execute', 'respond_debug_dialog', 'edit_watch', 'remove_watch', 'quick_watch', 'add_watch', 'rename_component')) {
        $task = $toolsType.GetMethod('InvokeAsync').Invoke($tools, @($tool, '{}'))
        $reply = $task.GetAwaiter().GetResult() | ConvertFrom-Json
        Assert (-not $reply.Ok -and $reply.Error.Contains($mode)) "$mode bypass through $tool."
    }
    $readReply = $toolsType.GetMethod('Invoke').Invoke($tools, @('read_module', '{"Project":"Projet","Module":"Module1"}')) | ConvertFrom-Json
    Assert $readReply.Ok "$mode blocked inspection."
}
$tools.Mode = [Enum]::Parse($modeType, 'Agent')
$tools.BoundProject = 'OtherProject'
$denied = $toolsType.GetMethod('Invoke').Invoke($tools, @('replace_lines', $requestJson)) | ConvertFrom-Json
Assert (-not $denied.Ok -and $denied.Error.Contains('autre projet')) 'Cross-document mutation was not blocked.'
$tools.BoundProject = $null
Write-Output 'PASS Discussion/Plan gates every sync/async mutation path; scoped edits'

$multi = New-Internal CodeChange
$multi.Before = "A`r`nold-one`r`nC`r`nD`r`nE`r`nF`r`nG`r`nold-two`r`nI"
$multi.After = "A`r`nnew-one`r`nC`r`nD`r`nE`r`nF`r`nG`r`nnew-two`r`nI"
$hunks = @(Static CodeRollback Hunks @($multi.Before, $multi.After))
Assert ($hunks.Count -eq 2) 'Separated edits should produce two rollback blocks.'
$current = "external-prefix`r`n" + $multi.After + "`r`nexternal-suffix"
$partial = Static CodeRollback Apply @($multi, $current, [Nullable[int]]0)
Assert ($partial.Contains('old-one') -and $partial.Contains('new-two') -and $partial.StartsWith('external-prefix')) 'Partial rollback damaged another edit.'
$multi.RestoredHunks.Add(0)
$complete = Static CodeRollback Apply @($multi, $partial, $null)
Assert ($complete -eq ("external-prefix`r`n" + $multi.Before + "`r`nexternal-suffix")) 'Second rollback lost independent changes.'
$conflict = $false
try { Static CodeRollback Apply @($multi, $partial.Replace('new-two','user-edit'), $null) | Out-Null } catch { $conflict = $true }
Assert $conflict 'Overlapping user edit should cause a conflict.'
$adjacent = New-Internal CodeChange
$adjacent.Before = "a`r`nb`r`nc`r`nd"; $adjacent.After = "A`r`nb`r`nC`r`nd"
$part = Static CodeRollback Apply @($adjacent, $adjacent.After, [Nullable[int]]0)
$adjacent.RestoredHunks.Add(0)
Assert ((Static CodeRollback Apply @($adjacent, $part, $null)) -eq $adjacent.Before) 'Nearby partial hunks could not be fully restored.'
Write-Output 'PASS multi-block rollback, shifted code preservation, overlapping conflict and partial continuation'

# Run the actual multi-edit restore against a mutable VBIDE fixture.
$module = $vbe.VBProjects[0].VBComponents[0].CodeModule
$module.Source.Clear(); $module.Source.AddRange([string[]]@('a','b','c'))
$first = New-Internal CodeChange
$first.Project = 'Projet'; $first.Module = 'Module1'; $first.Before = "a`r`nb`r`nc"; $first.After = "a`r`nB`r`nc"
$second = New-Internal CodeChange
$second.Project = 'Projet'; $second.Module = 'Module1'; $second.Before = $first.After; $second.After = "a`r`nB`r`nC"
$module.Source[1] = 'B'; $module.Source[2] = 'C'
$targets = [Array]::CreateInstance($changeType, 2); $targets.SetValue($first,0); $targets.SetValue($second,1)
$reply = $toolsType.GetMethod('RestoreChanges').Invoke($tools, @($targets,$null))
Assert ($reply.Ok -and $first.Restored -and $second.Restored -and ($module.Source -join ',') -eq 'a,b,c') 'Grouped rollback did not restore a chain of edits.'
Write-Output 'PASS intervention rollback applies reverse chain with live SHA checks'

Assert ((Static ChatCommand Expand @('/expliquer la boucle')).Contains('la boucle')) 'Slash command lost its arguments.'
$sessionState = New-Internal ChatSessionState
$sessionState.Title = 'History'; $sessionState.Scope = 'MacroA'
$entry = New-Internal ChatEntry; $entry.Speaker = 'Vous'; $entry.Text = 'unique-message-key'; $sessionState.Entries.Add($entry)
Assert (Static ChatHistory Matches @($sessionState,'unique-message-key')) 'Message search only inspected titles.'
$export = Static ChatHistory Export @($sessionState)
Assert ($export.Contains('unique-message-key') -and $export.Contains('MacroA')) 'Markdown export lost conversation content.'
Write-Output 'PASS slash arguments, message search and Markdown export'

# A failing preflight must not mutate another module first.
$first.Restored = $false; $first.RestoredHunks.Clear()
$second.Restored = $false; $second.RestoredHunks.Clear(); $second.Module = 'MissingModule'
$module.Source[1] = 'B'
$beforePreflight = $module.Source -join ','
$preflight = $toolsType.GetMethod('RestoreChanges').Invoke($tools, @($targets,$null))
Assert (-not $preflight.Ok -and ($module.Source -join ',') -eq $beforePreflight) 'Failed preflight partially edited another module.'
foreach ($pair in @(@('', 'one'), @('one', ''), @("a`r`nb", "a`r`n`r`nb"), @("a`r`nb`r`nc", "c`r`nb`r`na"))) {
    $edge = New-Internal CodeChange; $edge.Before = $pair[0]; $edge.After = $pair[1]
    Assert ((Static CodeRollback Apply @($edge,$edge.After,$null)) -eq $edge.Before) 'Insertion/deletion/blank-line rollback failed.'
}
Write-Output 'PASS preflight prevents partial writes; insertion, deletion and blank-line restoration'

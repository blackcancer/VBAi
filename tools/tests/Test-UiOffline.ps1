<#
.SYNOPSIS
Host-free workflow, compiled Designer or project-metadata checks. Designer
scenarios construct design surfaces without showing windows. ProjectMetadata
evaluates MSBuild items without loading VBAi.
.DESCRIPTION
Select exactly one scenario. No aggregate scenario is provided. Assembly and
output defaults remain specific to each scenario; pass an explicit candidate
assembly when required. Shared helper import opens no window or host.
.EXAMPLE
.\Test-UiOffline.ps1 -Scenario ChatWorkflow -AssemblyPath artifacts/candidate/VBAi.dll
.EXAMPLE
.\Test-UiOffline.ps1 -Scenario ProjectMetadata
#>
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('ChatWorkflow','ChatDesigner','WinFormsDesigners','ProjectMetadata')][string]$Scenario,
    [string]$AssemblyPath, [string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'UiProbe.psm1') -ErrorAction Stop

function Invoke-ChatWorkflow {
    param([string]$AssemblyPath = 'artifacts/chat-build/VBAi/Debug/net48/VBAi.dll')
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
    $assembly = Open-UiAssembly $AssemblyPath
    $flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
    $vbe = [FakeVbe]::new()
    $sessionType = $assembly.GetType('VBAi.VbeSession')
    $session = [Activator]::CreateInstance($sessionType, $flags, $null, @($vbe), $null)
    $requestType = $assembly.GetType('VBAi.Request')
    $read = [Activator]::CreateInstance($requestType)
    $read.Command = 'read_module'; $read.Project = 'Projet'; $read.Module = 'Module1'
    $before = $sessionType.GetMethod('Execute').Invoke($session, @($read))
    if (-not $before.Ok) { throw $before.Error }
    $settingsType = $assembly.GetType('VBAi.LlmSettings')
    $settings = [Activator]::CreateInstance($settingsType, $true)
    $settings.VbeEditApproval = 'AskEachTime'
    $toolsType = $assembly.GetType('VBAi.LlmVbeTools')
    $tools = [Activator]::CreateInstance($toolsType, $flags, $null, @($session, $null, $settings), $null)
    $requestJson = [string](@{ Project = 'Projet'; Module = 'Module1'; ExpectedSha256 = $before.Data.Sha256; StartLine = 2; Count = 1; Text = '    Debug.Print 2' } | ConvertTo-Json -Compress)
    $result = $toolsType.GetMethod('Invoke').Invoke($tools, @([string]'replace_lines', [string]$requestJson)) | ConvertFrom-Json
    if (-not $result.Ok -or $vbe.VBProjects[0].VBComponents[0].CodeModule.Source[1] -ne '    Debug.Print 2') { throw 'replace_lines failed.' }
    $changeType = $assembly.GetType('VBAi.CodeChange')
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

    function Static([string]$type, [string]$method, [object[]]$arguments) { $assembly.GetType("VBAi.$type").GetMethod($method).Invoke($null, $arguments) }
    $modeType = $assembly.GetType('VBAi.ChatMode')
    foreach ($mode in @('Discussion', 'Plan')) {
        $tools.Mode = [Enum]::Parse($modeType, $mode)
        $denied = $toolsType.GetMethod('Invoke').Invoke($tools, @('replace_lines', $requestJson)) | ConvertFrom-Json
        Assert-Ui (-not $denied.Ok) "$mode allowed a code mutation."
        foreach ($tool in @('immediate_execute', 'respond_debug_dialog', 'edit_watch', 'remove_watch', 'quick_watch', 'add_watch', 'rename_component')) {
            $task = $toolsType.GetMethod('InvokeAsync').Invoke($tools, @($tool, '{}'))
            $reply = $task.GetAwaiter().GetResult() | ConvertFrom-Json
            Assert-Ui (-not $reply.Ok -and $reply.Error.Contains($mode)) "$mode bypass through $tool."
        }
        $readReply = $toolsType.GetMethod('Invoke').Invoke($tools, @('read_module', '{"Project":"Projet","Module":"Module1"}')) | ConvertFrom-Json
        Assert-Ui $readReply.Ok "$mode blocked inspection."
    }
    $tools.Mode = [Enum]::Parse($modeType, 'Agent')
    $tools.BoundProject = 'OtherProject'
    $denied = $toolsType.GetMethod('Invoke').Invoke($tools, @('replace_lines', $requestJson)) | ConvertFrom-Json
    Assert-Ui (-not $denied.Ok -and $denied.Error.Contains('autre projet')) 'Cross-document mutation was not blocked.'
    $tools.BoundProject = $null
    Write-Output 'PASS Discussion/Plan gates every sync/async mutation path; scoped edits'

    $multi = New-UiObject $assembly CodeChange
    $multi.Before = "A`r`nold-one`r`nC`r`nD`r`nE`r`nF`r`nG`r`nold-two`r`nI"
    $multi.After = "A`r`nnew-one`r`nC`r`nD`r`nE`r`nF`r`nG`r`nnew-two`r`nI"
    $hunks = @(Static CodeRollback Hunks @($multi.Before, $multi.After))
    Assert-Ui ($hunks.Count -eq 2) 'Separated edits should produce two rollback blocks.'
    $current = "external-prefix`r`n" + $multi.After + "`r`nexternal-suffix"
    $partial = Static CodeRollback Apply @($multi, $current, [Nullable[int]]0)
    Assert-Ui ($partial.Contains('old-one') -and $partial.Contains('new-two') -and $partial.StartsWith('external-prefix')) 'Partial rollback damaged another edit.'
    $multi.RestoredHunks.Add(0)
    $complete = Static CodeRollback Apply @($multi, $partial, $null)
    Assert-Ui ($complete -eq ("external-prefix`r`n" + $multi.Before + "`r`nexternal-suffix")) 'Second rollback lost independent changes.'
    $conflict = $false
    try { Static CodeRollback Apply @($multi, $partial.Replace('new-two','user-edit'), $null) | Out-Null } catch { $conflict = $true }
    Assert-Ui $conflict 'Overlapping user edit should cause a conflict.'
    $adjacent = New-UiObject $assembly CodeChange
    $adjacent.Before = "a`r`nb`r`nc`r`nd"; $adjacent.After = "A`r`nb`r`nC`r`nd"
    $part = Static CodeRollback Apply @($adjacent, $adjacent.After, [Nullable[int]]0)
    $adjacent.RestoredHunks.Add(0)
    Assert-Ui ((Static CodeRollback Apply @($adjacent, $part, $null)) -eq $adjacent.Before) 'Nearby partial hunks could not be fully restored.'
    Write-Output 'PASS multi-block rollback, shifted code preservation, overlapping conflict and partial continuation'

    # Run the actual multi-edit restore against a mutable VBIDE fixture.
    $module = $vbe.VBProjects[0].VBComponents[0].CodeModule
    $module.Source.Clear(); $module.Source.AddRange([string[]]@('a','b','c'))
    $first = New-UiObject $assembly CodeChange
    $first.Project = 'Projet'; $first.Module = 'Module1'; $first.Before = "a`r`nb`r`nc"; $first.After = "a`r`nB`r`nc"
    $second = New-UiObject $assembly CodeChange
    $second.Project = 'Projet'; $second.Module = 'Module1'; $second.Before = $first.After; $second.After = "a`r`nB`r`nC"
    $module.Source[1] = 'B'; $module.Source[2] = 'C'
    $targets = [Array]::CreateInstance($changeType, 2); $targets.SetValue($first,0); $targets.SetValue($second,1)
    $reply = $toolsType.GetMethod('RestoreChanges').Invoke($tools, @($targets,$null))
    Assert-Ui ($reply.Ok -and $first.Restored -and $second.Restored -and ($module.Source -join ',') -eq 'a,b,c') 'Grouped rollback did not restore a chain of edits.'
    Write-Output 'PASS intervention rollback applies reverse chain with live SHA checks'

    Assert-Ui ((Static ChatCommand Expand @('/expliquer la boucle')).Contains('la boucle')) 'Slash command lost its arguments.'
    $sessionState = New-UiObject $assembly ChatSessionState
    $sessionState.Title = 'History'; $sessionState.Scope = 'MacroA'
    $entry = New-UiObject $assembly ChatEntry; $entry.Speaker = 'Vous'; $entry.Text = 'unique-message-key'; $sessionState.Entries.Add($entry)
    Assert-Ui (Static ChatHistory Matches @($sessionState,'unique-message-key')) 'Message search only inspected titles.'
    $export = Static ChatHistory Export @($sessionState)
    Assert-Ui ($export.Contains('unique-message-key') -and $export.Contains('MacroA')) 'Markdown export lost conversation content.'
    Write-Output 'PASS slash arguments, message search and Markdown export'

    # A failing preflight must not mutate another module first.
    $first.Restored = $false; $first.RestoredHunks.Clear()
    $second.Restored = $false; $second.RestoredHunks.Clear(); $second.Module = 'MissingModule'
    $module.Source[1] = 'B'
    $beforePreflight = $module.Source -join ','
    $preflight = $toolsType.GetMethod('RestoreChanges').Invoke($tools, @($targets,$null))
    Assert-Ui (-not $preflight.Ok -and ($module.Source -join ',') -eq $beforePreflight) 'Failed preflight partially edited another module.'
    foreach ($pair in @(@('', 'one'), @('one', ''), @("a`r`nb", "a`r`n`r`nb"), @("a`r`nb`r`nc", "c`r`nb`r`na"))) {
        $edge = New-UiObject $assembly CodeChange; $edge.Before = $pair[0]; $edge.After = $pair[1]
        Assert-Ui ((Static CodeRollback Apply @($edge,$edge.After,$null)) -eq $edge.Before) 'Insertion/deletion/blank-line rollback failed.'
    }
    Write-Output 'PASS preflight prevents partial writes; insertion, deletion and blank-line restoration'
}

function Invoke-ChatDesigner {
    param([string]$AssemblyPath = 'artifacts/chat-build/VBAi/Debug/net48/VBAi.dll')
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing, System.Design, WindowsFormsIntegration
    Add-Type -ReferencedAssemblies System.dll, System.Design.dll -TypeDefinition @'
using System.ComponentModel.Design;
public static class ChatDesignerSelection {
    public static void Select(object service, object component) { ((ISelectionService)service).SetSelectedComponents(new object[] { component }, SelectionTypes.Replace); }
    public static object Primary(object service) { return ((ISelectionService)service).PrimarySelection; }
}
'@
    $assembly = Open-UiAssembly $AssemblyPath
    $type = $assembly.GetType('VBAi.ChatWindow', $true)
    $flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
    $form = [Activator]::CreateInstance($type, $true)
    try {
        foreach ($name in @('send','newChat','history','selection','compile','modules','methods','contextToggle')) {
            Assert-Ui ((Get-UiField $form $name) -is [Windows.Forms.Button]) "$name is not a designer-editable WinForms button."
        }
        foreach ($name in @('scopePicker','providerPicker','modelPicker','effortPicker','modePicker')) {
            Assert-Ui ((Get-UiField $form $name) -is [Windows.Forms.ComboBox]) "$name is not a designer-editable WinForms picker."
        }
        Assert-Ui ((Get-UiField $form transcriptHost).Child -eq $null) 'Parameterless constructor started the transcript engine.'
        $input = Get-UiField $form promptHost
        Assert-Ui ($input.GetType().Name -eq 'ChatInputView' -and $input.Controls.Find('previewEditor', $true).Count -eq 1) 'Designer input representation is missing.'
        $editorField = $input.GetType().GetField('editor', $flags)
        Assert-Ui ($editorField.GetValue($input) -eq $null) 'Designer constructor started the input engine.'
        Assert-Ui ((Get-UiField $form settings) -eq $null -and (Get-UiField $form tools) -eq $null -and (Get-UiField $form sessionStore) -eq $null) 'Designer constructor started runtime services.'
        $button = Get-UiField $form send
        Assert-Ui ($button.Parent -ne $null -and (Get-UiField $form historyPanel).Parent -ne $null) 'Designer control hierarchy is incomplete.'
        $text = [ComponentModel.TypeDescriptor]::GetProperties($button)['Text']
        Assert-Ui (-not $text.IsReadOnly) 'Button caption is read-only in the property grid.'
        $before = $text.GetValue($button)
        $text.SetValue($button, 'Designer edit probe')
        Assert-Ui ($button.Text -eq 'Designer edit probe') 'Designer property edit was not applied.'
        $text.SetValue($button, $before)
        $source = [IO.File]::ReadAllText((Join-Path (Get-Location) 'src/VBAi/Llm/Chat/ChatWindow.Designer.cs'))
        Assert-Ui (-not $source.Contains('InitializeShell') -and -not $source.Contains('BuildWorkflowControls')) 'InitializeComponent calls a runtime UI factory.'
        Assert-Ui (-not $source.Contains('Maj+Entrée') -and -not $source.Contains('Entrée : envoyer')) 'Unwanted keyboard annotation remains.'
        Write-Output 'PASS complete WinForms hierarchy, inert designer constructor and removed keyboard annotation'
    } finally { $form.Dispose() }
    $surface = [ComponentModel.Design.DesignSurface]::new()
    try {
        $surface.BeginLoad($type)
        if (-not $surface.IsLoaded -or $surface.LoadErrors.Count -gt 0) { throw ($surface.LoadErrors -join "`n") }
        $designerHost = $surface.GetService([ComponentModel.Design.IDesignerHost])
        Assert-Ui ($designerHost.RootComponent -ne $null) 'Designer has no root component.'
        Assert-Ui ($designerHost.GetDesigner($designerHost.RootComponent) -ne $null) 'WinForms root designer was not loaded.'
        $view = $surface.View
        Assert-Ui ($view -is [Windows.Forms.Control]) 'WinForms designer did not expose its editable surface.'
        $button = Get-UiField $designerHost.RootComponent send
        $selection = $surface.GetService([ComponentModel.Design.ISelectionService])
        [ChatDesignerSelection]::Select($selection, $button)
        Assert-Ui ([ChatDesignerSelection]::Primary($selection) -eq $button) 'Designer cannot select the send button.'
        Write-Output 'PASS WinForms DesignSurface load and control selection (compiled inherited view)'
    } finally { $surface.Dispose() }
    $settingsSurface = [ComponentModel.Design.DesignSurface]::new()
    try {
        $settingsType = $assembly.GetType('VBAi.LlmSettingsWindow', $true)
        $settingsSurface.BeginLoad($settingsType)
        Assert-Ui ($settingsSurface.IsLoaded -and $settingsSurface.LoadErrors.Count -eq 0) 'Settings designer failed to load.'
        $settingsHost = $settingsSurface.GetService([ComponentModel.Design.IDesignerHost])
        foreach ($name in @('manualModels','customName','azureEntra','githubAccount','githubLogin','githubRefresh','githubStatus')) {
            $control = $settingsType.GetField($name, $flags).GetValue($settingsHost.RootComponent)
            Assert-Ui ($control -is [Windows.Forms.Control] -and $control.Parent -ne $null) "Missing provider designer control: $name"
        }
        Write-Output 'PASS provider settings DesignSurface and fixed provider controls'
    } finally { $settingsSurface.Dispose() }
    $approvalSurface = [ComponentModel.Design.DesignSurface]::new()
    try {
        $approvalType = $assembly.GetType('VBAi.VbeApprovalDialog', $true)
        $approvalSurface.BeginLoad($approvalType)
        Assert-Ui ($approvalSurface.IsLoaded -and $approvalSurface.LoadErrors.Count -eq 0) 'Approval dialog designer failed to load.'
        $approvalHost = $approvalSurface.GetService([ComponentModel.Design.IDesignerHost])
        foreach ($name in @('details','actions','approve','reject')) {
            $control = $approvalType.GetField($name, $flags).GetValue($approvalHost.RootComponent)
            Assert-Ui ($control -is [Windows.Forms.Control] -and $control.Parent -ne $null) "Missing approval designer control: $name"
        }
        Write-Output 'PASS approval dialog DesignSurface and fixed controls'
    } finally { $approvalSurface.Dispose() }
    $gitSurface = [ComponentModel.Design.DesignSurface]::new()
    try {
        $gitType = $assembly.GetType('VBAi.GitWindow', $true)
        $gitSurface.BeginLoad($gitType)
        Assert-Ui ($gitSurface.IsLoaded -and $gitSurface.LoadErrors.Count -eq 0) 'Git designer failed to load.'
        $gitHost = $gitSurface.GetService([ComponentModel.Design.IDesignerHost])
        Assert-Ui ($gitHost.RootComponent -ne $null -and $gitSurface.View -is [Windows.Forms.Control]) 'Git designer has no editable form.'
        Write-Output 'PASS Git DesignSurface load'
    } finally { $gitSurface.Dispose() }
}

function Invoke-WinFormsDesigners {
    param([string]$AssemblyPath = 'artifacts/designer-build/VBAi.dll', [string]$OutputDirectory = 'artifacts/designer-validation')
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing, System.Design
    $assembly = Open-UiAssembly $AssemblyPath
    [IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($OutputDirectory)) | Out-Null
    $sourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../src/VBAi'))
    $designerSources = @{}
    foreach ($file in Get-ChildItem -LiteralPath $sourceRoot -Filter '*.Designer.cs' -Recurse) {
        $source = [IO.File]::ReadAllText($file.FullName)
        if ($source -match 'partial class (\w+)') { $designerSources[$Matches[1]] = $source }
    }
    $proof = @()
    $types = $assembly.GetTypes() | Where-Object {
        -not $_.IsAbstract -and $_.Namespace -eq 'VBAi' -and
        ([Windows.Forms.Form].IsAssignableFrom($_) -or [Windows.Forms.UserControl].IsAssignableFrom($_))
    }
    foreach ($type in $types) {
        $surface = [ComponentModel.Design.DesignSurface]::new()
        try {
            # Model source-owned controls on a fresh Form/UserControl root. Loading the
            # compiled subclass directly models inherited UI and locks private fields.
            $context = New-Object System.ComponentModel.Design.DesigntimeLicenseContext
            $subject = [ComponentModel.LicenseManager]::CreateWithContext($type, $context)
            $surface.BeginLoad($type.BaseType)
            if (-not $surface.IsLoaded -or $surface.LoadErrors.Count) { throw "$($type.Name): $($surface.LoadErrors -join '; ')" }
            $designerHost = $surface.GetService([ComponentModel.Design.IDesignerHost])
            $root = $designerHost.RootComponent
            $root.Font = $subject.Font
            $root.Size = $subject.Size
            $root.MinimumSize = $subject.MinimumSize
            $root.Padding = $subject.Padding
            $root.AutoSize = $subject.AutoSize
            $root.Name = $type.Name
            foreach ($control in @($subject.Controls)) { $root.Controls.Add($control) }
            if ($null -eq $designerHost.GetDesigner($root) -or $surface.View -isnot [Windows.Forms.Control]) { throw "$($type.Name): no editable designer" }
            $property = [ComponentModel.TypeDescriptor]::GetProperties($root)['Size']
            $autoSize = [ComponentModel.TypeDescriptor]::GetProperties($root)['AutoSize']
            $wasAutoSize = $null -ne $autoSize -and $autoSize.GetValue($root)
            if ($wasAutoSize) { $autoSize.SetValue($root, $false) }
            $original = $property.GetValue($root)
            $property.SetValue($root, [Drawing.Size]::new($original.Width + 10, $original.Height + 10))
            if ($root.Width -ne $original.Width + 10) { throw "$($type.Name): runtime code overrides designer sizing" }
            $property.SetValue($root, $original)
            if ($wasAutoSize) { $autoSize.SetValue($root, $true) }
            $children = 0
            $designerSource = $designerSources[$type.Name]
            if (-not $designerSources.ContainsKey($type.Name)) {
                # This layout base owns no fixed components; its concrete views have Designers.
                $declaredControls = @($type.GetFields([Reflection.BindingFlags]'DeclaredOnly,Instance,NonPublic,Public') | Where-Object { [Windows.Forms.Control].IsAssignableFrom($_.FieldType) })
                if ($type.FullName -ne 'VBAi.ChatDesignerView' -or $subject.Controls.Count -ne 0 -or $declaredControls.Count -ne 0) { throw "$($type.Name): no Designer source" }
                $designerSource = ''
            }
            foreach ($match in [regex]::Matches($designerSource, 'this\.(\w+) = new [\w.]+\(')) {
                $field = $type.GetField($match.Groups[1].Value, [Reflection.BindingFlags]'Instance,NonPublic,Public')
                if ($null -eq $field) { continue }
                $child = $field.GetValue($subject)
                if ($child -isnot [Windows.Forms.Control]) { continue }
                if ([string]::IsNullOrEmpty($child.Name)) { throw "$($type.Name)/$($field.Name): unnamed Designer control" }
                # Register the fixed components declared by InitializeComponent in the design host.
                if ($null -eq $child.Site -or $child.Site.Container -ne $designerHost.Container) {
                    if ($null -ne $child.Site) { $child.Site.Container.Remove($child) }
                    $designerHost.Container.Add($child, $field.Name)
                }
                if ($null -eq $designerHost.GetDesigner($child)) { throw "$($type.Name)/$($field.Name): no child designer" }
                $description = [ComponentModel.TypeDescriptor]::GetProperties($child)['AccessibleDescription']
                if ($description.IsReadOnly) { throw "$($type.Name)/$($field.Name): editable property is read-only" }
                $prior = $description.GetValue($child)
                $description.SetValue($child, 'Designer edit verification')
                if ($description.GetValue($child) -ne 'Designer edit verification') { throw "$($type.Name)/$($field.Name): property edit failed" }
                $serialization = [ComponentModel.Design.Serialization.CodeDomComponentSerializationService]::new($designerHost)
                $store = $serialization.CreateStore()
                try {
                    $serialization.SerializeAbsolute($store, $child)
                    $store.Close()
                    if ($store.Errors.Count) { throw "$($type.Name)/$($field.Name): serialization errors: $($store.Errors -join '; ')" }
                    $copyContainer = New-Object System.ComponentModel.Container
                    try {
                        $copies = @($serialization.Deserialize($store, $copyContainer))
                        $copy = @($copies | Where-Object { $_ -is [Windows.Forms.Control] -and $_.Name -eq $child.Name })
                        if ($copy.Count -ne 1 -or $copy[0].AccessibleDescription -ne 'Designer edit verification') { throw "$($type.Name)/$($field.Name): edited property did not roundtrip" }
                    } finally { $copyContainer.Dispose() }
                } finally { $store.Dispose() }
                $description.SetValue($child, $prior)
                $children++
            }
            $proof += [pscustomobject]@{ View = $type.Name; Load = 'PASS'; Resize = 'PASS'; DesignerChildren = $children; ChildPropertyEdits = 'PASS'; SerializationRoundtrip = 'PASS'; AutoSizeRestored = $wasAutoSize }
            Write-Output "PASS $($type.Name): DesignSurface load, resize and $children editable child components"
        } finally { $surface.Dispose(); if ($null -ne $subject) { $subject.Dispose(); $subject = $null } }
    }
    $proof | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'designers.json') -Encoding UTF8
    Write-Output "PASS $($types.Count) WinForms designers"
}

function Invoke-ProjectMetadata {
    param([string]$OutputDirectory = 'artifacts/designer-compatibility/metadata')
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
    $proof = @()
    function Read-Items($project) {
        $json = & dotnet msbuild $project -nologo -p:DesignTimeBuild=true '-getItem:Compile,EmbeddedResource'
        if ($LASTEXITCODE -ne 0) { throw "Cannot evaluate $project" }
        return ($json -join "`n" | ConvertFrom-Json).Items
    }
    function Require-Item($items, $identity, $subtype, $parent) {
        $entry = @($items | Where-Object { $_.Identity -eq $identity })
        if ($entry.Count -ne 1) { throw "Missing or duplicate item $identity" }
        if ($subtype -and $entry[0].SubType -ne $subtype) { throw "$identity must have SubType=$subtype" }
        if ($parent -and $entry[0].DependentUpon -ne $parent) { throw "$identity must depend on $parent" }
        return [pscustomobject]@{ File = $identity; SubType = $entry[0].SubType; Parent = $entry[0].DependentUpon; Result = 'PASS' }
    }
    $items = Read-Items (Join-Path $repository 'src/VBAi/VBAi.csproj')
    foreach ($name in @('Editor\ModernEditorWindow','Updates\UpdateWindow','Updates\UpdateProgressWindow','Git\GitWindow','Llm\Settings\LlmSettingsWindow')) {
        $parent = ($name -split '\\')[-1] + '.cs'
        $proof += Require-Item $items.Compile ($name+'.cs') 'Form' $null
        $proof += Require-Item $items.Compile ($name+'.Designer.cs') 'Code' $parent
        $proof += Require-Item $items.EmbeddedResource ($name+'.resx') $null $parent
    }
    foreach ($entry in $items.Compile | Where-Object { $_.Identity -match '(ModernEditorWindow|GitWindow|LlmSettingsWindow)\.[^.]+\.cs$' }) {
        $parent = [regex]::Match($entry.Identity, '(ModernEditorWindow|GitWindow|LlmSettingsWindow)\.').Groups[1].Value + '.cs'
        $proof += Require-Item $items.Compile $entry.Identity 'Code' $parent
    }
    $updater = Read-Items (Join-Path $repository 'src/VBAi.Updater/VBAi.Updater.csproj')
    $proof += Require-Item $updater.Compile '../VBAi/Updates/UpdateProgressWindow.cs' 'Code' $null
    $proof += Require-Item $updater.Compile '../VBAi/Updates/UpdateProgressWindow.Designer.cs' 'Code' 'UpdateProgressWindow.cs'
    $directory = [IO.Path]::GetFullPath((Join-Path $repository $OutputDirectory))
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    $proof | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'project-items.json') -Encoding UTF8
    Write-Output "PASS $($proof.Count) evaluated WinForms project items"
}

$commands = @{
    'ChatWorkflow' = 'Invoke-ChatWorkflow'
    'ChatDesigner' = 'Invoke-ChatDesigner'
    'WinFormsDesigners' = 'Invoke-WinFormsDesigners'
    'ProjectMetadata' = 'Invoke-ProjectMetadata'
}
$command = Get-Command $commands[$Scenario] -CommandType Function
$arguments = Get-UiScenarioParameters $command $PSBoundParameters
& $command @arguments

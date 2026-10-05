param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Label', 'TextBox', 'ComboBox', 'CheckBox', 'OptionButton', 'ToggleButton', 'CommandButton')]
    [string] $ControlType,
    [Parameter(Mandatory = $true)] [ValidateRange(1, [int]::MaxValue)] [int] $HostProcessId,
    [ValidateNotNullOrEmpty()] [string] $Project = 'VBAProject',
    [string] $Form,
    [switch] $OwnedDisposableWorkbook
)

$ErrorActionPreference = 'Stop'
if (-not $OwnedDisposableWorkbook) {
    throw 'Confirm that the selected Excel process contains an owned disposable workbook with -OwnedDisposableWorkbook.'
}

# Each entry preserves the command, seed values, revision sequence, and extra
# response oracle from its original single-control probe.
$definitions = @{
    Label = @{ Form = 'CodexLabelCopySurvey'; ProgId = 'Forms.Label.1'; Source = 'lblOriginal'; Copy = 'lblCopy';
        Caption = 'Label source'; Width = 120; Property = 'BackColor'; Value = '#FF8800';
        Duplicate = 'duplicate_form_label'; Error = 'DuplicateLabel did not return the expected partial copy and new tree revision.' }
    TextBox = @{ Form = 'CodexTextBoxCopySurvey'; ProgId = 'Forms.TextBox.1'; Source = 'txtOriginal'; Copy = 'txtCopy';
        Width = 120; Property = 'Value'; Value = 'Texte Codex';
        Duplicate = 'duplicate_form_textbox'; Error = 'DuplicateTextBox did not return the expected partial copy and new tree revision.' }
    ComboBox = @{ Form = 'CodexComboCopySurvey'; ProgId = 'Forms.ComboBox.1'; Source = 'cmbOriginal'; Copy = 'cmbCopy';
        Width = 120; Property = 'ListWidth'; Value = '72 pt'; ExtraFalse = @('ItemsCopied', 'BindingsCopied');
        Duplicate = 'duplicate_form_combobox'; Error = 'DuplicateComboBox did not return the expected partial copy and new tree revision.' }
    CheckBox = @{ Form = 'CodexCheckBoxCopySurvey'; ProgId = 'Forms.CheckBox.1'; Source = 'chkOriginal'; Copy = 'chkCopy';
        Caption = 'Choix Codex'; Width = 120; Property = 'Value'; Value = $true;
        Duplicate = 'duplicate_form_checkbox'; Error = 'DuplicateCheckBox did not return the expected partial copy and new tree revision.' }
    OptionButton = @{ Form = 'CodexOptionCopySurvey'; ProgId = 'Forms.OptionButton.1'; Source = 'optOriginal'; Copy = 'optCopy';
        Caption = 'Premier choix'; Width = 130; ExtraFalse = @('SelectionCopied', 'GroupCopied');
        Duplicate = 'duplicate_form_optionbutton'; Error = 'DuplicateOptionButton did not return the expected partial copy and new tree revision.' }
    ToggleButton = @{ Form = 'CodexToggleButtonCopySurvey'; ProgId = 'Forms.ToggleButton.1'; Source = 'tglOriginal'; Copy = 'tglCopy';
        Caption = 'Choix Codex'; Width = 120; RefuseValueWrite = $true;
        Duplicate = 'duplicate_form_togglebutton'; Error = 'DuplicateToggleButton did not return the expected partial copy and new tree revision.' }
    CommandButton = @{ Form = 'CodexButtonCopySurvey'; ProgId = 'Forms.CommandButton.1'; Source = 'cmdOriginal'; Copy = 'cmdCopy';
        Caption = 'Exécuter'; Width = 120; ExtraFalse = @('EventsCopied');
        Duplicate = 'duplicate_form_commandbutton'; Error = 'DuplicateCommandButton did not return the expected partial copy and new tree revision.' }
}
$definition = $definitions[$ControlType]
if (-not $Form) { $Form = $definition.Form }
if ([string]::IsNullOrWhiteSpace($Form)) { throw 'A nonempty disposable form name is required.' }

$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') {
    throw "The probe requires Excel; PID $HostProcessId is $($hostProcess.ProcessName)."
}

function Invoke-Vbe([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $reply = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $reply.Ok) { throw "$($Request.Command): $($reply.Error)" }
    return $reply.Data
}

function Read-Tree {
    Invoke-Vbe @{ Command = 'form_tree'; Project = $Project; Form = $Form }
}

$existing = @(Invoke-Vbe @{ Command = 'list_forms'; Project = $Project })
if (@($existing | Where-Object { $_.Name -eq $Form }).Count) {
    throw "Form $Form already exists; use a new name in a disposable workbook."
}

Invoke-Vbe @{ Command = 'create_form'; Project = $Project; Form = $Form } | Out-Null
$initial = Read-Tree
$add = @{ Command = 'add_form_control'; Project = $Project; Form = $Form;
    ControlType = $definition.ProgId; Control = $definition.Source;
    Left = 24; Top = 24; Width = $definition.Width; Height = 30;
    ExpectedFormVersion = $initial.FormVersion }
if ($definition.ContainsKey('Caption')) { $add.Caption = $definition.Caption }
Invoke-Vbe $add | Out-Null

if ($definition.ContainsKey('Property')) {
    $withControl = Read-Tree
    Invoke-Vbe @{ Command = 'set_form_node_property'; Project = $Project; Form = $Form;
        ControlPath = ('Controls/' + $definition.Source); Property = $definition.Property;
        Value = $definition.Value; ExpectedTreeVersion = $withControl.TreeVersion } | Out-Null
}
$before = Read-Tree

$refusedReply = $null
if ($definition.RefuseValueWrite) {
    $refused = @{ Command = 'set_form_node_property'; Project = $Project; Form = $Form;
        ControlPath = ('Controls/' + $definition.Source); Property = 'Value'; Value = $true;
        ExpectedTreeVersion = $before.TreeVersion }
    $refusedJson = ConvertTo-Json -InputObject $refused -Compress -Depth 8
    $refusedReply = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $refusedJson | ConvertFrom-Json
    $afterRefusal = Read-Tree
    if ($refusedReply.Ok -or $refusedReply.Error -notmatch 'temporarily disabled' -or
        $afterRefusal.TreeVersion -ne $before.TreeVersion) {
        throw 'ToggleButton.Value=true was not refused before mutation.'
    }
}

$copy = Invoke-Vbe @{ Command = $definition.Duplicate; Project = $Project; Form = $Form;
    ControlPath = ('Controls/' + $definition.Source); NewName = $definition.Copy;
    ExpectedTreeVersion = $before.TreeVersion }
$after = Read-Tree
$extraFailed = $false
$extraProperties = if ($definition.ContainsKey('ExtraFalse')) { @($definition.ExtraFalse) } else { @() }
foreach ($property in $extraProperties) {
    if ($copy.$property -ne $false) { $extraFailed = $true }
}
if ($copy.Completeness -ne 'Partial' -or $copy.NewPath -ne ('Controls/' + $definition.Copy) -or
    $extraFailed -or $after.NodeCount -ne ($before.NodeCount + 1) -or
    $after.TreeVersion -eq $before.TreeVersion) {
    throw $definition.Error
}

$result = [ordered]@{
    HostProcessId = $HostProcessId
    Form = $Form
    SourcePath = $copy.SourcePath
    NewPath = $copy.NewPath
    Completeness = $copy.Completeness
}
foreach ($property in $extraProperties) { $result[$property] = $copy.$property }
$result.CopiedProperties = ($copy.CopiedProperties -join ', ')
if ($definition.RefuseValueWrite) { $result.ValueWriteRefused = $refusedReply.Error }
if ($ControlType -ne 'OptionButton') {
    $result.BeforeTreeVersion = $before.TreeVersion
    $result.AfterTreeVersion = $after.TreeVersion
}
$result.BeforeNodeCount = $before.NodeCount
$result.AfterNodeCount = $after.NodeCount
$result.ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
[pscustomobject]$result | Format-List

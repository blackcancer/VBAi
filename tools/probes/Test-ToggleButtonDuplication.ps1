param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [string] $Project = 'VBAProject',
    [string] $Form = 'CodexToggleButtonCopySurvey'
)

$ErrorActionPreference = 'Stop'
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
Invoke-Vbe @{ Command = 'add_form_control'; Project = $Project; Form = $Form;
    ControlType = 'Forms.ToggleButton.1'; Control = 'tglOriginal'; Caption = 'Choix Codex';
    Left = 24; Top = 24; Width = 120; Height = 30;
    ExpectedFormVersion = $initial.FormVersion } | Out-Null
$before = Read-Tree

$refused = @{ Command = 'set_form_node_property'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/tglOriginal'; Property = 'Value'; Value = $true;
    ExpectedTreeVersion = $before.TreeVersion }
$refusedJson = ConvertTo-Json -InputObject $refused -Compress -Depth 8
$refusedReply = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $refusedJson | ConvertFrom-Json
$afterRefusal = Read-Tree
if ($refusedReply.Ok -or $refusedReply.Error -notmatch 'temporarily disabled' -or
    $afterRefusal.TreeVersion -ne $before.TreeVersion) {
    throw 'ToggleButton.Value=true was not refused before mutation.'
}

$copy = Invoke-Vbe @{ Command = 'duplicate_form_togglebutton'; Project = $Project; Form = $Form;
    ControlPath = 'Controls/tglOriginal'; NewName = 'tglCopy';
    ExpectedTreeVersion = $before.TreeVersion }
$after = Read-Tree
if ($copy.Completeness -ne 'Partial' -or $copy.NewPath -ne 'Controls/tglCopy' -or
    $after.NodeCount -ne ($before.NodeCount + 1) -or $after.TreeVersion -eq $before.TreeVersion) {
    throw 'DuplicateToggleButton did not return the expected partial copy and new tree revision.'
}

[pscustomobject]@{
    HostProcessId = $HostProcessId
    Form = $Form
    SourcePath = $copy.SourcePath
    NewPath = $copy.NewPath
    Completeness = $copy.Completeness
    CopiedProperties = ($copy.CopiedProperties -join ', ')
    ValueWriteRefused = $refusedReply.Error
    BeforeTreeVersion = $before.TreeVersion
    AfterTreeVersion = $after.TreeVersion
    BeforeNodeCount = $before.NodeCount
    AfterNodeCount = $after.NodeCount
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

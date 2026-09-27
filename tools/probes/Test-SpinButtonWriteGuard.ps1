param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [Parameter(Mandatory = $true)] [string] $ExpectedCodeBase,
    [string] $Project = 'VBAProject',
    [string] $Form = 'CodexSpinGuardSurvey'
)

$ErrorActionPreference = 'Stop'
$actualCodeBase = (Get-ItemProperty -LiteralPath 'Registry::HKEY_CURRENT_USER\Software\Classes\CLSID\{8E854243-087F-4D6C-9E0E-8622B0E50883}\InprocServer32').CodeBase
if ($actualCodeBase -ne $ExpectedCodeBase) {
    throw "Refusing any write probe: COM CodeBase differs from guarded build. Actual=$actualCodeBase"
}
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') {
    throw "The probe requires Excel; PID $HostProcessId is $($hostProcess.ProcessName)."
}

function Invoke-VbeRaw([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
}

function Invoke-Vbe([hashtable] $Request) {
    $reply = Invoke-VbeRaw $Request
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
    ControlType = 'Forms.SpinButton.1'; Control = 'spnProbe';
    Left = 24; Top = 24; Width = 32; Height = 80;
    ExpectedFormVersion = $initial.FormVersion } | Out-Null
$before = Read-Tree
$refused = @()
foreach ($change in @(
    @{ Name = 'Min'; Value = 2 },
    @{ Name = 'Max'; Value = 20 },
    @{ Name = 'Value'; Value = 7 },
    @{ Name = 'Delay'; Value = 100 },
    @{ Name = 'SmallChange'; Value = 2 }
)) {
    $reply = Invoke-VbeRaw @{ Command = 'set_form_node_property'; Project = $Project; Form = $Form;
        ControlPath = 'Controls/spnProbe'; Property = $change.Name; Value = $change.Value;
        ExpectedTreeVersion = $before.TreeVersion }
    if ($reply.Ok -or $reply.Error -notmatch 'temporarily disabled') {
        throw "SpinButton.$($change.Name) was not refused before COM mutation."
    }
    $after = Read-Tree
    if ($after.TreeVersion -ne $before.TreeVersion -or $after.NodeCount -ne $before.NodeCount) {
        throw "SpinButton.$($change.Name) refusal changed the form tree."
    }
    $refused += $change.Name
}

[pscustomobject]@{
    HostProcessId = $HostProcessId
    GuardedCodeBase = $actualCodeBase
    Refused = ($refused -join ', ')
    TreeUnchanged = $true
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
} | Format-List

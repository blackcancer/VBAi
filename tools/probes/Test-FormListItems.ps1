param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') { throw 'The probe requires Excel.' }

function Invoke-Vbe([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $reply = & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $reply.Ok) { throw "$($Request.Command): $($reply.Error)" }
    return $reply.Data
}

$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) {
    throw 'Expected exactly one disposable VBA project in design mode.'
}
$project = $projects[0].Name
$form = 'CodexListItemsProbe'
if (@(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } |
    Where-Object { $_.Name -eq $form }).Count) { throw "Form $form already exists." }
Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form } | Out-Null
$tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
foreach ($spec in @(@('cboProbe','Forms.ComboBox.1'), @('lstProbe','Forms.ListBox.1'))) {
    Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
        Control = $spec[0]; ControlType = $spec[1]; Left = 24; Top = 24;
        Width = 120; Height = 30; ExpectedFormVersion = $tree.FormVersion } | Out-Null
    $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
}

$results = foreach ($name in @('cboProbe', 'lstProbe')) {
    $data = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
        ControlPath = "Controls/$name"; Offset = 0; Limit = 10 }
    if ($data.TotalRows -ne 0 -or $data.ReturnedRows -ne 0 -or $data.HasMore -or
        $data.TreeVersion -ne $tree.TreeVersion) {
        throw "Unexpected empty-list result for $name."
    }
    [pscustomobject]@{ ControlPath = $data.ControlPath; ControlType = $data.ControlType;
        TotalRows = $data.TotalRows; ColumnCount = $data.ColumnCount;
        ReturnedRows = $data.ReturnedRows; TreeVersionStable = $true }
}

$invalid = @{ Command = 'form_list_items'; Project = $project; Form = $form;
    ControlPath = 'Controls/missing'; Offset = 0; Limit = 10 } |
    ConvertTo-Json -Compress
$invalidReply = & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $invalid | ConvertFrom-Json
if ($invalidReply.Ok -or $invalidReply.Error -notmatch 'canonical') {
    throw 'A noncanonical ControlPath was not refused.'
}
$stable = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
if ($stable.TreeVersion -ne $tree.TreeVersion) { throw 'Read-only requests changed TreeVersion.' }

[pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project; Form = $form;
    Results = $results; NoncanonicalPathRejected = $true; TreeVersionStable = $true;
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue) } | ConvertTo-Json -Depth 7

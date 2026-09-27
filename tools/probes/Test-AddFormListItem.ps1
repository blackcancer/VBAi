param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') { throw 'The probe requires Excel.' }
function Invoke-Reply([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    & (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
}
function Invoke-Vbe([hashtable] $Request) {
    $reply = Invoke-Reply $Request
    if (-not $reply.Ok) { throw "$($Request.Command): $($reply.Error)" }
    return $reply.Data
}
$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) { throw 'Expected one disposable design-mode project.' }
$project = $projects[0].Name
$form = 'CodexListAppendProbe'
if (@(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } | Where-Object { $_.Name -eq $form }).Count) {
    throw "Form $form already exists."
}
Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form } | Out-Null
$tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
foreach ($spec in @(@('cboProbe','Forms.ComboBox.1'), @('lstProbe','Forms.ListBox.1'))) {
    Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
        Control = $spec[0]; ControlType = $spec[1]; Left = 24; Top = 24;
        Width = 120; Height = 30; ExpectedFormVersion = $tree.FormVersion } | Out-Null
    $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
}
$results = foreach ($name in @('cboProbe','lstProbe')) {
    $path = "Controls/$name"
    $before = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
        ControlPath = $path; Offset = 0; Limit = 64 }
    if ($before.TotalRows -ne 0 -or -not $before.ListVersion -or $before.ColumnCount -ne 1) {
        throw "$name is not an empty versioned one-column list."
    }
    $first = Invoke-Vbe @{ Command = 'add_form_list_item'; Project = $project; Form = $form;
        ControlPath = $path; Text = 'Alpha'; ExpectedTreeVersion = $before.TreeVersion;
        ExpectedListVersion = $before.ListVersion }
    if (-not $first.Applied -or -not $first.Verified) { throw "$name Alpha append was not verified." }
    $one = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
        ControlPath = $path; Offset = 0; Limit = 64 }
    $stale = Invoke-Reply @{ Command = 'add_form_list_item'; Project = $project; Form = $form;
        ControlPath = $path; Text = 'Wrong'; ExpectedTreeVersion = $one.TreeVersion;
        ExpectedListVersion = $before.ListVersion }
    if ($stale.Ok -or $stale.Error -notmatch 'changed since it was read') {
        throw "$name stale ListVersion was not refused."
    }
    $afterStale = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
        ControlPath = $path; Offset = 0; Limit = 64 }
    if ($afterStale.ListVersion -ne $one.ListVersion -or $afterStale.TotalRows -ne 1) {
        throw "$name changed after the stale refusal."
    }
    $second = Invoke-Vbe @{ Command = 'add_form_list_item'; Project = $project; Form = $form;
        ControlPath = $path; Text = 'Beta'; ExpectedTreeVersion = $one.TreeVersion;
        ExpectedListVersion = $one.ListVersion }
    $final = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
        ControlPath = $path; Offset = 0; Limit = 64 }
    $values = @($final.Rows | ForEach-Object { $_.Cells[0].Value })
    if (-not $second.Applied -or -not $second.Verified -or $final.TotalRows -ne 2 -or
        $values.Count -ne 2 -or $values[0] -ne 'Alpha' -or $values[1] -ne 'Beta') {
        throw "$name second append/readback failed."
    }
    [pscustomobject]@{ Path = $path; Values = $values; StaleVersionRefused = $true;
        ListVersionChanged = ($final.ListVersion -ne $one.ListVersion) }
}
[pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project; Form = $form;
    Results = $results; ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue) } |
    ConvertTo-Json -Depth 6

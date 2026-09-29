param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') { throw 'The probe requires Excel.' }

function Invoke-Vbe([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $reply = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
    if (-not $reply.Ok) { throw "$($Request.Command): $($reply.Error)" }
    return $reply.Data
}

$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) {
    throw 'Expected exactly one disposable VBA project in design mode.'
}
$project = $projects[0].Name
$form = 'CodexListItemsProbe'
$existing = @(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } |
    Where-Object { $_.Name -eq $form })
if (-not $existing.Count) {
    Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form } | Out-Null
    $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
    foreach ($spec in @(@('cboProbe','Forms.ComboBox.1'), @('lstProbe','Forms.ListBox.1'))) {
        Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
            Control = $spec[0]; ControlType = $spec[1]; Left = 24; Top = 24;
            Width = 120; Height = 30; ExpectedFormVersion = $tree.FormVersion } | Out-Null
        $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
    }
}

$results = foreach ($name in @('cboProbe', 'lstProbe')) {
    $data = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
        ControlPath = "Controls/$name"; Offset = 0; Limit = 10 }
    $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
    if ($data.TreeVersion -ne $tree.TreeVersion -or
        ($data.TotalRows -ne 0 -and $data.TotalRows -ne 5)) {
        throw "Unexpected initial list result for $name."
    }
    if ($data.TotalRows -eq 0) {
        foreach ($value in @('Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon')) {
            $current = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
            $append = Invoke-Vbe @{ Command = 'probe_append_form_list_item'; Project = $project;
                Form = $form; ControlPath = "Controls/$name"; Text = $value;
                ExpectedTreeVersion = $current.TreeVersion }
            if (-not $append.Applied -or -not $append.Verified -or
                $append.CountAfter -ne ($append.CountBefore + 1)) {
                throw "AddItem was not verified for $name / $value."
            }
        }
    }
    $first = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
        ControlPath = "Controls/$name"; Offset = 0; Limit = 2 }
    $second = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
        ControlPath = "Controls/$name"; Offset = 2; Limit = 2 }
    $last = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
        ControlPath = "Controls/$name"; Offset = 4; Limit = 2 }
    $pages = @($first, $second, $last)
    $actual = @($pages | ForEach-Object { $_.Rows } | ForEach-Object { $_.Cells } |
        ForEach-Object { $_.Value })
    $expected = @('Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon')
    if ($actual.Count -ne $expected.Count -or
        (@(Compare-Object -ReferenceObject $expected -DifferenceObject $actual -SyncWindow 0).Count -ne 0) -or
        $first.TotalRows -ne 5 -or $first.ReturnedRows -ne 2 -or -not $first.HasMore -or
        $second.ReturnedRows -ne 2 -or -not $second.HasMore -or
        $last.ReturnedRows -ne 1 -or $last.HasMore -or
        @($pages | Where-Object { $_.TreeVersion -ne $first.TreeVersion }).Count) {
        throw "Paged List readback failed for $name."
    }
    [pscustomobject]@{ ControlPath = $data.ControlPath; ControlType = $data.ControlType;
        TotalRows = $first.TotalRows; ColumnCount = $first.ColumnCount;
        PageSizes = @($first.ReturnedRows, $second.ReturnedRows, $last.ReturnedRows);
        Values = $actual; TreeVersionStable = $true }
}

$invalid = @{ Command = 'form_list_items'; Project = $project; Form = $form;
    ControlPath = 'Controls/missing'; Offset = 0; Limit = 10 } |
    ConvertTo-Json -Compress
$invalidReply = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $invalid | ConvertFrom-Json
if ($invalidReply.Ok -or $invalidReply.Error -notmatch 'canonical') {
    throw 'A noncanonical ControlPath was not refused.'
}
$stable = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
$stableAgain = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
if ($stable.TreeVersion -ne $stableAgain.TreeVersion) { throw 'Read-only requests changed TreeVersion.' }

[pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project; Form = $form;
    Results = $results; NoncanonicalPathRejected = $true; TreeVersionStable = $true;
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue) } | ConvertTo-Json -Depth 7

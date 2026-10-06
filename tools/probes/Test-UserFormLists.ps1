param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Append', 'ReadPages', 'Remove', 'InitializeWithoutEvent')]
    [string] $Scenario,
    [Parameter(Mandatory = $true)] [ValidateRange(1, [int]::MaxValue)] [int] $HostProcessId,
    [switch] $OwnedDisposableWorkbook
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'VbeProbe.Common.ps1')
$scenarioParameters = @{
    'InitializeWithoutEvent' = @{ Allowed = @('HostProcessId','OwnedDisposableWorkbook'); Required = @('HostProcessId','OwnedDisposableWorkbook') }
    'ReadPages' = @{ Allowed = @('HostProcessId','OwnedDisposableWorkbook'); Required = @('HostProcessId','OwnedDisposableWorkbook') }
    'Remove' = @{ Allowed = @('HostProcessId','OwnedDisposableWorkbook'); Required = @('HostProcessId','OwnedDisposableWorkbook') }
    'Append' = @{ Allowed = @('HostProcessId','OwnedDisposableWorkbook'); Required = @('HostProcessId','OwnedDisposableWorkbook') }
}
$contract = $scenarioParameters[$Scenario]
Assert-VbeProbeParameters -Bound $PSBoundParameters -Required $contract.Required -Allowed $contract.Allowed
if (-not $OwnedDisposableWorkbook) {
    throw 'Confirm an owned disposable Excel workbook with -OwnedDisposableWorkbook.'
}
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') { throw 'The probe requires Excel.' }

$VbeProbeJsonDepth = 8
$VbeProbeRawError = $false
$VbeProbeTraceRequests = $false

switch ($Scenario) {
    'Append' {
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
    }
    'ReadPages' {
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
    }
    'Remove' {
        $projects = @(Invoke-Vbe @{ Command = 'list_projects' })
        if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) { throw 'Expected one disposable design-mode project.' }
        $project = $projects[0].Name
        $form = 'CodexListRemovalProbe'
        $control = 'lstProbe'
        if (@(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } | Where-Object { $_.Name -eq $form }).Count) {
            throw "Form $form already exists."
        }
        Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form } | Out-Null
        $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
        Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
            Control = $control; ControlType = 'Forms.ListBox.1'; Left = 24; Top = 24;
            Width = 140; Height = 80; ExpectedFormVersion = $tree.FormVersion } | Out-Null
        foreach ($value in @('Alpha', 'Beta', 'Gamma')) {
            $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
            $append = Invoke-Vbe @{ Command = 'probe_append_form_list_item'; Project = $project;
                Form = $form; ControlPath = "Controls/$control"; Text = $value;
                ExpectedTreeVersion = $tree.TreeVersion }
            if (-not $append.Verified) { throw "AddItem not verified for $value." }
        }
        $stale = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
            ControlPath = "Controls/$control"; Offset = 0; Limit = 64 }
        if ($stale.TotalRows -ne 3 -or -not $stale.ListVersion) { throw 'Full list version is unavailable.' }
        $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
        Invoke-Vbe @{ Command = 'probe_append_form_list_item'; Project = $project;
            Form = $form; ControlPath = "Controls/$control"; Text = 'Delta';
            ExpectedTreeVersion = $tree.TreeVersion } | Out-Null
        $current = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
            ControlPath = "Controls/$control"; Offset = 0; Limit = 64 }
        if ($current.TotalRows -ne 4 -or $current.ListVersion -eq $stale.ListVersion) {
            throw 'The content revision did not detect the append.'
        }
        $staleRequest = @{ Command = 'remove_form_list_item'; Project = $project;
            Form = $form; ControlPath = "Controls/$control"; RowIndex = 1;
            ExpectedTreeVersion = $current.TreeVersion; ExpectedListVersion = $stale.ListVersion }
        $staleReply = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson (ConvertTo-Json $staleRequest -Compress) | ConvertFrom-Json
        if ($staleReply.Ok -or $staleReply.Error -notmatch 'list changed') {
            throw 'A stale ListVersion was not refused before RemoveItem.'
        }
        $unchanged = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
            ControlPath = "Controls/$control"; Offset = 0; Limit = 64 }
        if ($unchanged.ListVersion -ne $current.ListVersion) { throw 'Stale refusal changed the list.' }
        $removed = Invoke-Vbe @{ Command = 'remove_form_list_item'; Project = $project;
            Form = $form; ControlPath = "Controls/$control"; RowIndex = 1;
            ExpectedTreeVersion = $current.TreeVersion; ExpectedListVersion = $current.ListVersion }
        if (-not $removed.Applied -or -not $removed.Verified -or $removed.RemovedValue -ne 'Beta') {
            throw 'RemoveItem was not verified.'
        }
        $after = Invoke-Vbe @{ Command = 'form_list_items'; Project = $project; Form = $form;
            ControlPath = "Controls/$control"; Offset = 0; Limit = 64 }
        $values = @($after.Rows | ForEach-Object { $_.Cells[0].Value })
        if ($after.TotalRows -ne 3 -or ($values -join '|') -ne 'Alpha|Gamma|Delta' -or
            $after.ListVersion -ne $removed.ListVersionAfter) {
            throw 'The requested item did not disappear exactly from the list.'
        }
        [pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project; Form = $form;
            StaleVersionRefused = $true; RemovedValue = $removed.RemovedValue;
            FinalValues = $values; FinalListVersion = $after.ListVersion;
            ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue) } | ConvertTo-Json -Depth 5
    }
    'InitializeWithoutEvent' {
        $projects = @(Invoke-Vbe @{ Command = 'list_projects' })
        if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) {
            throw 'Expected exactly one disposable project in design mode.'
        }
        $project = $projects[0].Name
        $form = 'CodexAutoInitializeProbe'
        if (@(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } | Where-Object { $_.Name -eq $form }).Count) {
            throw "Form $form already exists."
        }
        Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form } | Out-Null
        $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
        Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
            Control = 'cboProbe'; ControlType = 'Forms.ComboBox.1'; Left = 20; Top = 20;
            Width = 120; Height = 25; ExpectedFormVersion = $tree.FormVersion } | Out-Null
        $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
        $before = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $form }
        if ($before.Code -match 'UserForm_Initialize') { throw 'Expected no existing Initialize procedure.' }
        $request = @{ Command = 'set_form_list_initializer'; Project = $project; Form = $form;
            ControlPath = 'Controls/cboProbe'; Items = @('Alpha','Beta');
            ExpectedTreeVersion = $tree.TreeVersion; ExpectedSha256 = $before.Sha256 }
        $first = Invoke-Vbe $request
        $after = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $form }
        if (-not $first.Applied -or -not $first.Verified -or $first.VerificationPending -or
            $after.Code -notmatch 'Private Sub UserForm_Initialize' -or
            $after.Code -notmatch 'Me.cboProbe.AddItem "Alpha"' -or
            $after.Code -notmatch 'Me.cboProbe.AddItem "Beta"') {
            throw 'Automatic Initialize creation or generated block failed verification.'
        }
        $request.ExpectedSha256 = $after.Sha256
        $second = Invoke-Vbe $request
        $again = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $form }
        if ($second.Applied -or -not $second.Verified -or $again.Sha256 -ne $after.Sha256) {
            throw 'Repeating the same initializer was not idempotent.'
        }
        [pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project; Form = $form;
            CreatedInitialize = $true; ItemsGenerated = 2; Idempotent = $true;
            CodeSha256 = $again.Sha256; ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue) } |
        ConvertTo-Json -Depth 5
    }
}

<#
.SYNOPSIS
Runs one explicit SOLIDWORKS VBA scenario against the selected preloaded host.
.DESCRIPTION
The existing test.swp project, mode and source inventory are checked before each
scenario. The command never launches or closes SOLIDWORKS. Round trips and debug
scenarios retain their own verification and finally blocks. Run with Windows
PowerShell -STA when requesting native-pane inspection.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidateSet('Breakpoint','Debug','ClassFormRoundTrip','ModuleRoundTrip')][string]$Scenario,
    [ValidateRange(1,2147483647)][int]$HostProcessId,
    [ValidateSet('Class','Form')][string]$Kind,
    [switch]$InspectNativePanes
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../probes/VbeProbe.Common.ps1')
$VbeProbeJsonDepth=8
$VbeProbeRawError=$false
$required=@('HostProcessId');$allowed=@('HostProcessId')
if($Scenario -eq 'ClassFormRoundTrip'){$required+='Kind';$allowed+='Kind'}
if($Scenario -eq 'Breakpoint'){$allowed+='InspectNativePanes'}
Assert-VbeProbeParameters -Bound $PSBoundParameters -Required $required -Allowed $allowed
if ((Get-Process -Id $HostProcessId -ErrorAction Stop).ProcessName -ne 'SLDWORKS') {
    throw 'An already-open SOLIDWORKS process is required.'
}

$project = 'test'
$projectFile = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\test.swp'))
$projects = @(Invoke-Vbe @{ Command = 'list_projects' })


switch ($Scenario) {
    'Breakpoint' {
        $ErrorActionPreference = 'Stop'


        function Get-CommandById([int] $Id) {
            return @(Invoke-Vbe @{ Command = 'list_commands'; Query = '' }) |
            Where-Object { $_.Id -eq $Id -and $_.Enabled } | Select-Object -First 1
        }


        if (@($projects | Where-Object { $_.Name -eq $project -and $_.FileName -eq $projectFile -and $_.Mode -eq 2 }).Count -ne 1) {
            throw 'The disposable SOLIDWORKS VBA project is not uniquely available in design mode.'
        }
        $module = 'CodexSwBreakpointProbe'
        $baseline = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
        if (@($baseline | Where-Object { $_.Name -eq $module }).Count -ne 0) {
            throw 'The disposable module name already exists.'
        }
        $original = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
        $source = $null
        $outputPath = Join-Path $env:TEMP ("VBAi-sw-breakpoint-{0}.txt" -f $HostProcessId)
        if (Test-Path -LiteralPath $outputPath) { throw 'The disposable output path already exists.' }
        $breakpointSet = $false
        $breakpointRemovalInvoked = $false
        $breakLine = 5
        $stopped = $false
        $stepped = $false
        $continued = $false
        $localsProbe = $null
        try {
            $created = Invoke-Vbe @{ Command = 'create_module'; Project = $project; Module = $module; ExpectedMode = 2 }
            $code = "Option Explicit`r`nPublic Sub CodexSwBreakpointSmoke()`r`n    Dim probeValue As Long`r`n    probeValue = 1`r`n    probeValue = probeValue + 1`r`n    Open `"$outputPath`" For Output As #1`r`n    Print #1, CStr(probeValue)`r`n    Close #1`r`nEnd Sub"
            Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
                ExpectedSha256 = $created.Sha256; StartLine = 1; Count = $created.Lines; Text = $code } | Out-Null
            $settled = $false
            for ($attempt = 0; $attempt -lt 10; $attempt++) {
                $source = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
                Start-Sleep -Milliseconds 100
                $stable = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
                if ($stable.Sha256 -eq $source.Sha256) { $settled = $true; break }
                $source = $stable
            }
            if (-not $settled) { throw 'The temporary module did not settle after editing.' }
            Invoke-Vbe @{ Command = 'select_code'; Project = $project; Module = $module;
                ExpectedSha256 = $source.Sha256; StartLine = $breakLine } | Out-Null
            $breakCommand = Get-CommandById 51
            if (-not $breakCommand) { throw 'The native Toggle Breakpoint command is unavailable.' }
            Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
                ExpectedSha256 = $source.Sha256; StartLine = $breakLine; ExpectedMode = 2;
                Action = 'toggle_breakpoint'; ControlId = 51; ControlCaption = $breakCommand.Caption } | Out-Null
            $breakpointSet = $true
            $source = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
            Invoke-Vbe @{ Command = 'select_code'; Project = $project; Module = $module;
                ExpectedSha256 = $source.Sha256; StartLine = 2 } | Out-Null
            $runCommand = Get-CommandById 186
            if (-not $runCommand) { throw 'The native Run Sub command is unavailable.' }
            Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
                ExpectedSha256 = $source.Sha256; StartLine = 2; ExpectedMode = 2;
                Action = 'run'; ControlId = 186; ControlCaption = $runCommand.Caption } | Out-Null
            for ($attempt = 0; $attempt -lt 30; $attempt++) {
                $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
                if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq $breakLine) { $stopped = $true; break }
                Start-Sleep -Milliseconds 100
            }
            if (-not $stopped) { throw "The VBA project did not stop at line $breakLine (mode=$($state.Mode))." }
            Invoke-Vbe @{ Command = 'open_debug_pane'; Action = 'locals' } | Out-Null
            if ($InspectNativePanes) {
                Invoke-Vbe @{ Command = 'open_debug_pane'; Action = 'immediate' } | Out-Null
                & (Join-Path $PSScriptRoot '..\probes\Inspect-VbeNativeWindows.ps1') -View Panes -HostProcessId $HostProcessId
                & (Join-Path $PSScriptRoot '..\probes\Inspect-VbeAccessibility.ps1') -Target MsaaChildren -HostProcessId $HostProcessId
            }
            for ($attempt = 0; $attempt -lt 10; $attempt++) {
                $windows = Invoke-Vbe @{ Command = 'debug_windows' }
                $localsProbe = $windows.Locals
                if (@($localsProbe.Items).Count -gt 0) { break }
                Start-Sleep -Milliseconds 100
            }
            $stepCommand = Get-CommandById 194
            if (-not $stepCommand) { throw 'The native Step Over command is unavailable in break mode.' }
            Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
                ExpectedSha256 = $source.Sha256; StartLine = $breakLine; ExpectedMode = 1;
                Action = 'step_over'; ControlId = 194; ControlCaption = $stepCommand.Caption } | Out-Null
            for ($attempt = 0; $attempt -lt 30; $attempt++) {
                $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
                if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq ($breakLine + 1)) { $stepped = $true; break }
                Start-Sleep -Milliseconds 100
            }
            if (-not $stepped) { throw "Step Over did not reach line $($breakLine + 1) (mode=$($state.Mode))." }
            $continueCommand = Get-CommandById 186
            if (-not $continueCommand) { throw 'The native Continue command is unavailable in break mode.' }
            Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
                ExpectedSha256 = $source.Sha256; StartLine = ($breakLine + 1); ExpectedMode = 1;
                Action = 'continue'; ControlId = 186; ControlCaption = $continueCommand.Caption } | Out-Null
            for ($attempt = 0; $attempt -lt 30; $attempt++) {
                $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
                if ($state.Mode -eq 2 -and (Test-Path -LiteralPath $outputPath) -and
                    (Get-Content -LiteralPath $outputPath -Raw).Trim() -eq '2') {
                    $continued = $true
                    break
                }
                Start-Sleep -Milliseconds 100
            }
            if (-not $continued) { throw "Continue did not finish with expected value 2 (mode=$($state.Mode))." }
        }
        finally {
            $cleanupError = $null
            $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
            if ($state.Mode -eq 1) {
                Invoke-Vbe @{ Command = 'debug_global'; Project = $project; ExpectedMode = 1; Action = 'reset' } | Out-Null
                for ($attempt = 0; $attempt -lt 20; $attempt++) {
                    $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
                    if ($state.Mode -eq 2) { break }
                    Start-Sleep -Milliseconds 100
                }
            }
            if ($state.Mode -ne 2) { throw 'The SOLIDWORKS VBE did not return to design mode.' }
            if ($breakpointSet -and $source) {
                try {
                    $source = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
                    Invoke-Vbe @{ Command = 'select_code'; Project = $project; Module = $module;
                        ExpectedSha256 = $source.Sha256; StartLine = $breakLine } | Out-Null
                    $breakCommand = Get-CommandById 51
                    if (-not $breakCommand) { throw 'Cannot remove the disposable breakpoint.' }
                    Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
                        ExpectedSha256 = $source.Sha256; StartLine = $breakLine; ExpectedMode = 2;
                        Action = 'toggle_breakpoint'; ControlId = 51; ControlCaption = $breakCommand.Caption } | Out-Null
                    $breakpointRemovalInvoked = $true
                }
                catch { $cleanupError = "Breakpoint removal was not verified: $_" }
            }
            $current = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
            if (@($current | Where-Object { $_.Name -eq $module }).Count -eq 1) {
                $projectState = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
                $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $module }
                Invoke-Vbe @{ Command = 'remove_component'; Project = $project; Module = $module;
                    ExpectedProjectVersion = $projectState.Version; ExpectedComponentVersion = $component.Version } | Out-Null
            }
            $last = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
            $after = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
            if ($last.Count -ne $baseline.Count -or $after.Sha256 -ne $original.Sha256) {
                throw 'The SOLIDWORKS project did not return to its initial code/component inventory.'
            }
            if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }
            if ($cleanupError) { throw $cleanupError }
        }
        [pscustomobject]@{ HostProcessId = $HostProcessId; BreakpointHit = $stopped;
            StepOverVerified = $stepped; ContinueVerified = $continued;
            BreakpointRemovalInvoked = $breakpointRemovalInvoked;
            LocalsAvailable = $localsProbe.Available; LocalsError = $localsProbe.Error;
            LocalsRows = @($localsProbe.Items).Count; LocalsSnapshot = ($localsProbe | ConvertTo-Json -Compress -Depth 3);
            OriginalModuleShaPreserved = $true; ModuleInventoryRestored = $true } | Format-List

    }
    'Debug' {
        $ErrorActionPreference = 'Stop'





        if (@($projects | Where-Object { $_.Name -eq $project -and $_.FileName -eq $projectFile -and $_.Mode -eq 2 }).Count -ne 1) {
            throw 'The expected disposable SOLIDWORKS project is not uniquely available in design mode.'
        }
        $module = 'CodexSwDebugProbe'
        $baseline = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
        if (@($baseline | Where-Object { $_.Name -eq $module }).Count -ne 0) {
            throw 'The disposable module name already exists.'
        }
        $original = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
        $outputPath = Join-Path $env:TEMP ("VBAi-sw-debug-{0}.txt" -f $HostProcessId)
        if (Test-Path -LiteralPath $outputPath) { throw 'The disposable output path already exists.' }
        $verified = $false
        try {
            $created = Invoke-Vbe @{ Command = 'create_module'; Project = $project; Module = $module; ExpectedMode = 2 }
            $code = "Option Explicit`r`nPublic Sub CodexSwDebugSmoke()`r`n    Open `"$outputPath`" For Output As #1`r`n    Print #1, `"CodexSwDebugSmoke:42`"`r`n    Close #1`r`nEnd Sub"
            Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
                ExpectedSha256 = $created.Sha256; StartLine = 1; Count = $created.Lines; Text = $code } | Out-Null
            $source = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
            $compiled = Invoke-Vbe @{ Command = 'compile_project'; Project = $project; ExpectedMode = 2 }
            $diagnostic = Invoke-Vbe @{ Command = 'debug_dialog' }
            if ($diagnostic.Visible) { throw 'Compilation raised a native diagnostic dialog.' }
            $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
            if ($state.Mode -ne 2) { throw 'Compilation did not leave the project in design mode.' }
            Invoke-Vbe @{ Command = 'open_debug_pane'; Action = 'immediate' } | Out-Null
            Invoke-Vbe @{ Command = 'select_code'; Project = $project; Module = $module;
                ExpectedSha256 = $source.Sha256; StartLine = 2 } | Out-Null
            $runCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = '' }) |
            Where-Object { $_.Id -eq 186 -and $_.Enabled } | Select-Object -First 1
            if (-not $runCommand) { throw 'The native Run Sub command is unavailable.' }
            $run = Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
                ExpectedSha256 = $source.Sha256; StartLine = 2; ExpectedMode = 2;
                Action = 'run'; ControlId = 186; ControlCaption = $runCommand.Caption }
            $observed = $false
            for ($attempt = 0; $attempt -lt 30; $attempt++) {
                $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
                if ($state.Mode -eq 2 -and (Test-Path -LiteralPath $outputPath) -and
                    (Get-Content -LiteralPath $outputPath -Raw) -match 'CodexSwDebugSmoke:42') {
                    $observed = $true
                    break
                }
                Start-Sleep -Milliseconds 100
            }
            if (-not $observed) { throw "The native run did not produce the expected output in design mode (mode=$($state.Mode))." }
            $verified = $true
        }
        finally {
            try {
                $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
                if ($state.Mode -eq 1) {
                    Invoke-Vbe @{ Command = 'debug_global'; Project = $project; ExpectedMode = 1; Action = 'reset' } | Out-Null
                }
            }
            catch { Write-Warning "Could not verify/reset the debugger: $_" }
            $current = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
            if (@($current | Where-Object { $_.Name -eq $module }).Count -eq 1) {
                $projectState = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
                $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $module }
                Invoke-Vbe @{ Command = 'remove_component'; Project = $project; Module = $module;
                    ExpectedProjectVersion = $projectState.Version; ExpectedComponentVersion = $component.Version } | Out-Null
            }
            $last = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
            $after = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
            if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }
            if ($last.Count -ne $baseline.Count -or $after.Sha256 -ne $original.Sha256) {
                throw 'The SOLIDWORKS test project did not return to its original component/code inventory.'
            }
        }
        [pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project;
            CompileAndRunVerified = $verified; OriginalModuleShaPreserved = $true;
            ModuleInventoryRestored = $true } | Format-List

    }
    'ClassFormRoundTrip' {
        $ErrorActionPreference = 'Stop'





        if (@($projects | Where-Object { $_.Name -eq $project -and $_.FileName -eq $projectFile -and $_.Mode -eq 2 }).Count -ne 1) {
            throw 'The expected disposable SOLIDWORKS VBA project is unavailable.'
        }
        $name = if ($Kind -eq 'Class') { 'CodexSwClassRoundTrip' } else { 'CodexSwFormRoundTrip' }
        $extension = if ($Kind -eq 'Class') { '.cls' } else { '.frm' }
        $path = Join-Path $env:TEMP ("VBAi-sw-{0}-{1}{2}" -f $Kind.ToLowerInvariant(), $HostProcessId, $extension)
        $frx = [IO.Path]::ChangeExtension($path, '.frx')
        if ((Test-Path -LiteralPath $path) -or (Test-Path -LiteralPath $frx)) {
            throw 'A disposable export path already exists.'
        }
        $baseline = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
        if (@($baseline | Where-Object { $_.Name -eq $name }).Count -ne 0) {
            throw 'The disposable component name already exists.'
        }
        $original = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
        $verified = $false
        try {
            if ($Kind -eq 'Class') {
                $created = Invoke-Vbe @{ Command = 'create_class'; Project = $project;
                    Module = $name; ExpectedMode = 2 }
                $code = "Option Explicit`r`nPublic Function Value() As Long`r`n    Value = 42`r`nEnd Function"
                Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $name;
                    ExpectedSha256 = $created.Sha256; StartLine = 1; Count = $created.Lines; Text = $code } | Out-Null
            }
            else {
                Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $name } | Out-Null
                $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $name }
                Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $name;
                    ExpectedFormVersion = $tree.FormVersion; ControlType = 'Forms.Label.1';
                    Control = 'lblProbe'; Left = 12; Top = 12; Width = 80; Height = 20;
                    Caption = 'SolidWorks probe' } | Out-Null
                $state = Invoke-Vbe @{ Command = 'form_state'; Project = $project; Form = $name }
                Invoke-Vbe @{ Command = 'set_form_property'; Project = $project; Form = $name;
                    ExpectedFormVersion = $state.Version; Property = 'Caption'; Value = 'VBAi form probe' } | Out-Null
                $state = Invoke-Vbe @{ Command = 'form_state'; Project = $project; Form = $name }
                Invoke-Vbe @{ Command = 'set_form_property'; Project = $project; Form = $name;
                    ExpectedFormVersion = $state.Version; Property = 'Width'; Value = 240 } | Out-Null
                foreach ($entry in @(
                        @{ Property = 'Caption'; Value = 'Label from VBAi' },
                        @{ Property = 'Font.Name'; Value = 'Arial' },
                        @{ Property = 'Font.Size'; Value = 14 }
                )) {
                    $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $name }
                    Invoke-Vbe @{ Command = 'set_form_node_property'; Project = $project; Form = $name;
                        ControlPath = 'Controls/lblProbe'; ExpectedTreeVersion = $tree.TreeVersion;
                        Property = $entry.Property; Value = $entry.Value } | Out-Null
                }
                $state = Invoke-Vbe @{ Command = 'form_state'; Project = $project; Form = $name }
                $label = @($state.Controls | Where-Object { $_.Name -eq 'lblProbe' }) | Select-Object -First 1
                if ($state.Caption -ne 'VBAi form probe' -or [Math]::Abs($state.Width - 240) -gt 0.1 -or
                    $label.Caption -ne 'Label from VBAi' -or $label.FontName -ne 'Arial' -or
                    [Math]::Abs($label.FontSize - 14) -gt 0.1) {
                    throw 'SOLIDWORKS did not retain the form/label designer properties.'
                }
            }
            $codeBefore = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $name }
            $treeBefore = if ($Kind -eq 'Form') { Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $name } } else { $null }
            $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $name }
            $exported = Invoke-Vbe @{ Command = 'export_component'; Project = $project;
                Module = $name; Path = $path; ExpectedComponentVersion = $component.Version }
            if (-not (Test-Path -LiteralPath $path) -or $exported.Bytes -lt 1) {
                throw 'The exported component file is missing or empty.'
            }
            $projectState = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
            $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $name }
            $removed = Invoke-Vbe @{ Command = 'remove_component'; Project = $project; Module = $name;
                ExpectedProjectVersion = $projectState.Version; ExpectedComponentVersion = $component.Version }
            $imported = Invoke-Vbe @{ Command = 'import_component'; Project = $project;
                Path = $path; ExpectedProjectVersion = $removed.Version }
            if (-not $imported.Applied -or $imported.ImportedName -ne $name) {
                throw 'The imported component was not identified.'
            }
            $codeAfter = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $name }
            if ($codeAfter.Sha256 -ne $codeBefore.Sha256) {
                throw 'The code SHA changed during export/import.'
            }
            if ($Kind -eq 'Form') {
                $treeAfter = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $name }
                $stateAfter = Invoke-Vbe @{ Command = 'form_state'; Project = $project; Form = $name }
                $labelAfter = @($stateAfter.Controls | Where-Object { $_.Name -eq 'lblProbe' }) | Select-Object -First 1
                if ($treeAfter.TreeVersion -ne $treeBefore.TreeVersion -or
                    @($treeAfter.Controls | Where-Object { $_.Name -eq 'lblProbe' }).Count -ne 1 -or
                    $stateAfter.Caption -ne 'VBAi form probe' -or [Math]::Abs($stateAfter.Width - 240) -gt 0.1 -or
                    $labelAfter.Caption -ne 'Label from VBAi' -or $labelAfter.FontName -ne 'Arial' -or
                    [Math]::Abs($labelAfter.FontSize - 14) -gt 0.1) {
                    throw 'The UserForm designer tree changed during export/import.'
                }
            }
            $verified = $true
        }
        finally {
            $current = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
            if (@($current | Where-Object { $_.Name -eq $name }).Count -eq 1) {
                $projectState = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
                $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $name }
                Invoke-Vbe @{ Command = 'remove_component'; Project = $project; Module = $name;
                    ExpectedProjectVersion = $projectState.Version; ExpectedComponentVersion = $component.Version } | Out-Null
            }
            $last = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
            $originalAfter = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
            if ($last.Count -ne $baseline.Count -or $originalAfter.Sha256 -ne $original.Sha256) {
                throw 'The original SOLIDWORKS VBA project inventory was not restored.'
            }
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
            if (Test-Path -LiteralPath $frx) { Remove-Item -LiteralPath $frx -Force }
        }
        [pscustomobject]@{ HostProcessId = $HostProcessId; Kind = $Kind;
            RoundTripVerified = $verified; ModuleInventoryRestored = $true;
            OriginalModuleShaPreserved = $true } | Format-List

    }
    'ModuleRoundTrip' {
        $ErrorActionPreference = 'Stop'





        $target = @($projects | Where-Object { $_.Name -eq $project -and $_.FileName -eq $projectFile -and $_.Mode -eq 2 })
        if ($target.Count -ne 1) {
            throw "The expected SOLIDWORKS project matched $($target.Count) entries. Name=$($projects[-1].Name -eq $project) Path=$($projects[-1].FileName -eq $projectFile) Mode=$($projects[-1].Mode -eq 2); expected=$projectFile; actual=$($projects[-1].FileName)."
        }
        $baseline = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
        $module = 'CodexSwRoundTrip'
        if (@($baseline | Where-Object { $_.Name -eq $module }).Count -ne 0) {
            throw 'The disposable module name already exists.'
        }
        $existing = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
        $path = Join-Path $env:TEMP ("VBAi-sw-module-{0}.bas" -f $HostProcessId)
        if (Test-Path -LiteralPath $path) { throw 'The disposable export path already exists.' }
        $roundTrip = $false
        $cleanupVerified = $false
        try {
            $created = Invoke-Vbe @{ Command = 'create_module'; Project = $project;
                Module = $module; ExpectedMode = 2 }
            $code = "Option Explicit`r`nPublic Sub CodexSwSmoke()`r`n    Debug.Print 42`r`nEnd Sub"
            Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
                ExpectedSha256 = $created.Sha256; StartLine = 1; Count = $created.Lines; Text = $code } | Out-Null
            $before = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
            $procedure = Invoke-Vbe @{ Command = 'list_procedures'; Project = $project; Module = $module }
            if (@($procedure.Procedures | Where-Object { $_.Name -eq 'CodexSwSmoke' }).Count -ne 1) {
                throw 'The SOLIDWORKS VBE did not recognize the test procedure.'
            }
            $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $module }
            $exported = Invoke-Vbe @{ Command = 'export_component'; Project = $project;
                Module = $module; Path = $path; ExpectedComponentVersion = $component.Version }
            if (-not (Test-Path -LiteralPath $path) -or $exported.Bytes -lt 1) {
                throw 'The module export file is missing or empty.'
            }
            $projectState = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
            $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $module }
            $removed = Invoke-Vbe @{ Command = 'remove_component'; Project = $project; Module = $module;
                ExpectedProjectVersion = $projectState.Version; ExpectedComponentVersion = $component.Version }
            $imported = Invoke-Vbe @{ Command = 'import_component'; Project = $project;
                Path = $path; ExpectedProjectVersion = $removed.Version }
            if (-not $imported.Applied -or $imported.ImportedName -ne $module) {
                throw 'The module import was not confirmed.'
            }
            $after = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
            if ($after.Sha256 -ne $before.Sha256) { throw 'The imported VBA code SHA differs.' }
            $roundTrip = $true
        }
        finally {
            $current = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
            if (@($current | Where-Object { $_.Name -eq $module }).Count -eq 1) {
                $projectState = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
                $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $module }
                Invoke-Vbe @{ Command = 'remove_component'; Project = $project; Module = $module;
                    ExpectedProjectVersion = $projectState.Version; ExpectedComponentVersion = $component.Version } | Out-Null
            }
            $last = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
            $original = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
            $cleanupVerified = @($last | Where-Object { $_.Name -eq $module }).Count -eq 0 -and
            $original.Sha256 -eq $existing.Sha256 -and $last.Count -eq $baseline.Count
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
            if (-not $cleanupVerified) { throw 'The SOLIDWORKS test project did not return to its original component/code inventory.' }
        }
        [pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project;
            RoundTripVerified = $roundTrip; OriginalModuleShaPreserved = $cleanupVerified;
            ModuleInventoryRestored = $cleanupVerified } | Format-List

    }
}

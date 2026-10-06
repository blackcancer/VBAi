<#
.SYNOPSIS
Runs one explicit debugger probe against a disposable Excel fixture.
.DESCRIPTION
PID scenarios attach to the selected owned host and keep its existing lifetime.
Assembly scenarios create their own isolated fixture. Each invocation runs only
one scenario. BreakpointStage retains its separate -Stage execution contract.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('BreakpointStage','BreakpointCommandState','ClearAllBreakpoints','CompileProject','StepOver','NestedCallStack','QuickWatch','EditWatchType','RuntimeDialog','ShowNextStatementGuard','RegisteredTree')]
    [string]$Scenario,
    [ValidateRange(1,2147483647)][int]$HostProcessId,
    [switch]$OwnedDisposableHost,
    [ValidateSet('insert','state','toggle','run','continue','step_into','remove')][string]$Stage,
    [ValidateNotNullOrEmpty()][string]$AssemblyPath,
    [ValidateNotNullOrEmpty()][string]$OutputDirectory,
    [switch]$UseBridge,
    [switch]$AllowTemporaryVbaAccess
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'VbeProbe.Common.ps1')
$VbeProbeJsonDepth = 8
$VbeProbeRawError = $false
$VbeProbeTraceRequests = $false
if ($Scenario -in @('BreakpointCommandState','RegisteredTree')) {
    $required = @('AssemblyPath','OutputDirectory')
    $allowed = @('AssemblyPath','OutputDirectory','AllowTemporaryVbaAccess')
    if ($Scenario -eq 'RegisteredTree') { $required += 'UseBridge'; $allowed += 'UseBridge' }
} else {
    $required = @('HostProcessId','OwnedDisposableHost')
    $allowed = @('HostProcessId','OwnedDisposableHost')
    if ($Scenario -eq 'BreakpointStage') {
        $required += 'Stage'; $allowed += 'Stage'
        $VbeProbeJsonDepth = 10; $VbeProbeRawError = $true
    }
}
Assert-VbeProbeParameters -Bound $PSBoundParameters -Required $required -Allowed $allowed

switch ($Scenario) {
    'BreakpointStage' {
        $ErrorActionPreference = 'Stop'

        $hosts = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
        if ($hosts.Count -ne 1 -or $hosts[0].Id -ne $HostProcessId) {
            throw 'The breakpoint probe requires one isolated Excel process matching HostProcessId.'
        }



        $project = 'VBAProject'
        $module = 'ThisWorkbook'
        if ($Stage -eq 'insert') {
            $before = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
            if ($before.Code.Length -ne 0) { throw 'The probe module is not empty.' }
            $code = @'
Public Sub CodexBreakpointProbe()
    Dim probeValue As Long
    probeValue = 1
    probeValue = probeValue + 1
    Debug.Print probeValue
End Sub
'@
            Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
                ExpectedSha256 = $before.Sha256; StartLine = 1; Count = 0; Text = $code }
            $after = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
            Write-Output "Inserted temporary probe: $($after.Sha256)"
            return
        }

        $source = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
        if ($source.Code -notmatch 'CodexBreakpointProbe') { throw 'Probe procedure is missing.' }
        $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
        if ($Stage -eq 'state') { $state; return }

        $action = switch ($Stage) {
            'toggle' { 'toggle_breakpoint' }
            'remove' { 'toggle_breakpoint' }
            'run' { 'run' }
            'continue' { 'continue' }
            'step_into' { 'step_into' }
        }
        $controlId = switch ($Stage) {
            'toggle' { 51 }
            'remove' { 51 }
            'run' { 186 }
            'continue' { 186 }
            'step_into' { 188 }
        }
        $targetLine = if ($Stage -eq 'toggle' -or $Stage -eq 'remove') { 4 }
        elseif ($Stage -eq 'run') { 1 }
        else { [int]$state.Selection.StartLine }
        $null = Invoke-Vbe @{ Command = 'select_code'; Project = $project; Module = $module;
            ExpectedSha256 = $source.Sha256; StartLine = $targetLine }
        $query = switch ($Stage) {
            'toggle' { 'point' }
            'remove' { 'point' }
            'run' { 'Sub/UserForm' }
            'continue' { 'Contin' }
            'step_into' { 'pas' }
        }
        $commands = @(Invoke-Vbe @{ Command = 'list_commands'; Query = $query })
        $control = $commands | Where-Object { $_.Id -eq $controlId -and $_.Enabled } | Select-Object -First 1
        if (-not $control) { throw "Native VBE command $controlId is absent or disabled." }
        Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
            ExpectedSha256 = $source.Sha256; StartLine = $targetLine;
            ExpectedMode = $state.Mode; Action = $action;
            ControlId = $controlId; ControlCaption = $control.Caption }
        Invoke-Vbe @{ Command = 'debug_state'; Project = $project }

    }
    'BreakpointCommandState' {
        $ErrorActionPreference = 'Stop'
        . $VbeProbeInitializeAssembly

        try {
            . $VbeProbeOpenAssemblyExcel
            $vbe.MainWindow.Visible = $true
            $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
            $project = $book.VBProject
            $module = $project.VBComponents.Add(1); $module.Name = 'BreakpointStateProbe'
            $module.CodeModule.AddFromString("Public Sub ProbeState()`r`n    Debug.Print 1`r`n    Debug.Print 2`r`nEnd Sub")
            $pane = $module.CodeModule.CodePane
            $pane.Show(); $pane.SetSelection(2, 1, 2, 1)
            $toggle = $vbe.CommandBars.FindControl(1, 51)
            $clear = $vbe.CommandBars.FindControl(1, 579)
            if ($null -eq $toggle -or $toggle.Caption -notmatch 'point|breakpoint' -or -not $toggle.Enabled) { throw 'Expected native toggle command is unavailable.' }
            $rows = @()
            $rows += [pscustomobject]@{ Stage = 'before'; Line = 2; ToggleState = $toggle.State; ClearEnabled = $clear.Enabled }
            $toggle.Execute()
            $rows += [pscustomobject]@{ Stage = 'added'; Line = 2; ToggleState = $toggle.State; ClearEnabled = $clear.Enabled }
            $pane.SetSelection(3, 1, 3, 1)
            $rows += [pscustomobject]@{ Stage = 'other-line'; Line = 3; ToggleState = $toggle.State; ClearEnabled = $clear.Enabled }
            $pane.SetSelection(2, 1, 2, 1); $toggle.Execute()
            $rows += [pscustomobject]@{ Stage = 'removed'; Line = 2; ToggleState = $toggle.State; ClearEnabled = $clear.Enabled }
            $rows | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'breakpoint-command-state.json') -Encoding UTF8
            $rows | Format-Table
        } finally {
            # Restore trust before Quit, which can block in some COM teardown paths.
            if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
            else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
            if ($null -ne $otherBook) { $otherBook.Close($false) }
            if ($null -ne $book) { $book.Close($false) }
            if ($null -ne $excel) { $excel.Quit() }
            if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) {
                # This process was created after the no-Excel precondition and contains only
                # the disposable workbook. Quit can leave it alive because of COM references.
                Stop-Process -Id $probeProcess.Id -Force
            }
        }

    }
    'ClearAllBreakpoints' {
        $ErrorActionPreference = 'Stop'
        $excel = Get-Process -Id $HostProcessId -ErrorAction Stop
        if ($excel.ProcessName -ne 'EXCEL') { throw 'A disposable Excel process is required.' }



        $before = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
        if ($before.Code.Length -ne 0) { throw 'Use a fresh workbook with an empty ThisWorkbook module.' }
        $code = @'
Public Sub CodexClearProbe()
    Dim value As Long
    value = 1
    value = value + 1
    Debug.Print "CodexClearProbe:" & CStr(value)
End Sub
'@
        $null = Invoke-Vbe @{ Command = 'replace_lines'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $before.Sha256; StartLine = 1; Count = 0; Text = $code }
        $source = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
        $null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 3 }
        $breakCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'point' }) |
        Where-Object { $_.Id -eq 51 -and $_.Enabled } | Select-Object -First 1
        if (-not $breakCommand) { throw 'Native Toggle Breakpoint command 51 is unavailable.' }
        foreach ($line in @(3, 4)) {
            $null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
                ExpectedSha256 = $source.Sha256; StartLine = $line; ExpectedMode = 2;
                Action = 'toggle_breakpoint'; ControlId = 51; ControlCaption = $breakCommand.Caption }
        }
        $null = Invoke-Vbe @{ Command = 'open_debug_pane'; Action = 'immediate' }
        $null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 1 }
        $runCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'Sub/UserForm' }) |
        Where-Object { $_.Id -eq 186 -and $_.Enabled } | Select-Object -First 1
        if (-not $runCommand) { throw 'Native Run Sub command 186 is unavailable.' }
        function Start-Probe {
            $null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
                ExpectedSha256 = $source.Sha256; StartLine = 1 }
            $null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
                ExpectedSha256 = $source.Sha256; StartLine = 1; ExpectedMode = 2;
                Action = 'run'; ControlId = 186; ControlCaption = $runCommand.Caption }
        }
        Start-Probe
        $state = $null
        for ($attempt = 0; $attempt -lt 30; $attempt++) {
            $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
            if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq 3) { break }
            Start-Sleep -Milliseconds 100
        }
        if ($state.Mode -ne 1 -or $state.Selection.StartLine -ne 3) {
            throw 'The probe did not stop at its first breakpoint.'
        }
        $null = Invoke-Vbe @{ Command = 'debug_global'; Project = 'VBAProject'; ExpectedMode = 1;
            Action = 'reset' }
        $reset = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
        if ($reset.Mode -ne 2) { throw 'Reset did not restore design mode.' }
        $cleared = Invoke-Vbe @{ Command = 'debug_global'; Project = 'VBAProject'; ExpectedMode = 2;
            Action = 'clear_all_breakpoints' }
        Start-Probe
        $text = $null
        for ($attempt = 0; $attempt -lt 30; $attempt++) {
            $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
            if ($state.Mode -eq 1) { throw "Execution still stopped at line $($state.Selection.StartLine)." }
            $windows = Invoke-Vbe @{ Command = 'debug_windows' }
            $text = $windows.Immediate.Text
            if ($text -match 'CodexClearProbe:2') { break }
            Start-Sleep -Milliseconds 100
        }
        if ($state.Mode -ne 2 -or $text -notmatch 'CodexClearProbe:2') {
            throw "The procedure did not finish with its expected output. Mode=$($state.Mode); Immediate=$text"
        }
        [pscustomobject]@{
            HostProcessId = $HostProcessId
            FirstStopLine = 3
            ClearCommandId = $cleared.ControlId
            FinalMode = $state.Mode
            OutputObserved = 'CodexClearProbe:2'
        } | Format-List

    }
    'CompileProject' {
        $ErrorActionPreference = 'Stop'
        $hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
        if ($hostProcess.ProcessName -ne 'EXCEL') { throw 'An isolated Excel process is required.' }



        $project = 'VBAProject'
        $module = 'ThisWorkbook'
        $initial = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
        if ($initial.Code.Length -ne 0) { throw 'Use a fresh disposable workbook with an empty ThisWorkbook module.' }

        $invalid = "Option Explicit`r`nPublic Sub CodexCompileProbe()`r`n    missingValue = 1`r`nEnd Sub"
        Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
            ExpectedSha256 = $initial.Sha256; StartLine = 1; Count = 0; Text = $invalid } | Out-Null
        $failure = Invoke-Vbe @{ Command = 'compile_project'; Project = $project; ExpectedMode = 2 }
        $location = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
        if ($failure.Compiled -or $failure.Verification -ne 'NativeDiagnosticCaptured' -or
            $failure.Diagnostic -notmatch 'Variable non définie|Variable not defined' -or
            $location.ActiveModule -ne $module -or $location.Selection.StartLine -ne 3) {
            throw 'The expected native compile diagnostic or source selection was not captured.'
        }

        $current = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
        $corrected = "Option Explicit`r`nPublic Sub CodexCompileProbe()`r`n    Dim definedValue As Long`r`n    definedValue = 1`r`nEnd Sub"
        $lineCount = @($current.Code -split '\r?\n').Count
        Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
            ExpectedSha256 = $current.Sha256; StartLine = 1; Count = $lineCount; Text = $corrected } | Out-Null
        $success = Invoke-Vbe @{ Command = 'compile_project'; Project = $project; ExpectedMode = 2 }
        if (-not $success.Compiled -or $success.Diagnostic -or
            $success.Verification -ne 'NoNativeDiagnosticObserved') {
            throw 'The corrected VBA code did not compile without a native diagnostic.'
        }

        [pscustomobject]@{
            HostProcessId = $HostProcessId
            ErrorDiagnostic = $failure.Diagnostic
            ErrorLine = $location.Selection.StartLine
            CorrectedCompiled = $success.Compiled
            ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
        } | Format-List

    }
    'StepOver' {
        $ErrorActionPreference = 'Stop'
        $excel = Get-Process -Id $HostProcessId -ErrorAction Stop
        if ($excel.ProcessName -ne 'EXCEL') { throw 'A disposable Excel process is required.' }



        $before = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
        if ($before.Code.Length -ne 0) { throw 'Use a fresh workbook with an empty ThisWorkbook module.' }
        $code = @'
Public Sub CodexStepOuter()
    Dim value As Long
    value = 1
    CodexStepInner value
    Debug.Print value
End Sub
Private Sub CodexStepInner(ByRef value As Long)
    value = value + 5
End Sub
'@
        $null = Invoke-Vbe @{ Command = 'replace_lines'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $before.Sha256; StartLine = 1; Count = 0; Text = $code }
        $source = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
        $null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 4 }
        $breakCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'point' }) |
        Where-Object { $_.Id -eq 51 -and $_.Enabled } | Select-Object -First 1
        if (-not $breakCommand) { throw 'Native Toggle Breakpoint command 51 is unavailable.' }
        $null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 4; ExpectedMode = 2;
            Action = 'toggle_breakpoint'; ControlId = 51; ControlCaption = $breakCommand.Caption }
        $null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 1 }
        $runCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'Sub/UserForm' }) |
        Where-Object { $_.Id -eq 186 -and $_.Enabled } | Select-Object -First 1
        if (-not $runCommand) { throw 'Native Run Sub command 186 is unavailable.' }
        $null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 1; ExpectedMode = 2;
            Action = 'run'; ControlId = 186; ControlCaption = $runCommand.Caption }
        $state = $null
        for ($attempt = 0; $attempt -lt 30; $attempt++) {
            $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
            if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq 4) { break }
            Start-Sleep -Milliseconds 100
        }
        if ($state.Mode -ne 1 -or $state.Selection.StartLine -ne 4) {
            throw 'The outer procedure did not stop at its call line.'
        }
        $stepCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'pas à pas principal' }) |
        Where-Object { $_.Enabled -and $_.Caption -match 'principal' } | Select-Object -First 1
        if (-not $stepCommand) { throw 'Native Step Over command is unavailable.' }
        $null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 4; ExpectedMode = 1;
            Action = 'step_over'; ControlId = $stepCommand.Id; ControlCaption = $stepCommand.Caption }
        for ($attempt = 0; $attempt -lt 30; $attempt++) {
            $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
            if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq 5) { break }
            Start-Sleep -Milliseconds 100
        }
        if ($state.Mode -ne 1 -or $state.Selection.StartLine -ne 5) {
            throw 'Step Over did not stop at the following outer line.'
        }
        $null = Invoke-Vbe @{ Command = 'open_debug_pane'; Action = 'locals' }
        $windows = Invoke-Vbe @{ Command = 'debug_windows' }
        $valueRow = @($windows.Locals.Items) |
        Where-Object { $_.Expression -eq 'value' -and $_.Value -eq '6' } | Select-Object -First 1
        if (-not $valueRow) {
            throw "The called procedure's effect was not visible in Locals: $($windows.Locals | ConvertTo-Json -Compress -Depth 5)"
        }
        [pscustomobject]@{
            HostProcessId = $HostProcessId
            StepCommandId = $stepCommand.Id
            BreakLine = 4
            StepOverLine = $state.Selection.StartLine
            LocalValue = $valueRow.Value
        } | Format-List

    }
    'NestedCallStack' {
        $ErrorActionPreference = 'Stop'
        $excel = Get-Process -Id $HostProcessId -ErrorAction Stop
        if ($excel.ProcessName -ne 'EXCEL') { throw 'A disposable Excel process is required.' }



        $before = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
        if ($before.Code.Length -ne 0) { throw 'Use a fresh workbook with an empty ThisWorkbook module.' }
        $code = @'
Public Sub CodexStackOuter()
    Dim outerValue As Long
    outerValue = 1
    CodexStackInner
    Debug.Print outerValue
End Sub
Private Sub CodexStackInner()
    Dim innerValue As Long
    innerValue = 2
    innerValue = innerValue + 1
End Sub
'@
        $null = Invoke-Vbe @{ Command = 'replace_lines'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $before.Sha256; StartLine = 1; Count = 0; Text = $code }
        $source = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
        $null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 10 }

        $breakCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'point' }) |
        Where-Object { $_.Id -eq 51 -and $_.Enabled } | Select-Object -First 1
        if (-not $breakCommand) { throw 'Native Toggle Breakpoint command 51 is unavailable.' }
        $null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 10; ExpectedMode = 2;
            Action = 'toggle_breakpoint'; ControlId = 51; ControlCaption = $breakCommand.Caption }
        $null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 1 }
        $runCommand = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'Sub/UserForm' }) |
        Where-Object { $_.Id -eq 186 -and $_.Enabled } | Select-Object -First 1
        if (-not $runCommand) { throw 'Native Run Sub command 186 is unavailable.' }
        $null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 1; ExpectedMode = 2;
            Action = 'run'; ControlId = 186; ControlCaption = $runCommand.Caption }

        $state = $null
        for ($attempt = 0; $attempt -lt 20; $attempt++) {
            $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
            if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq 10) { break }
            Start-Sleep -Milliseconds 100
        }
        if ($state.Mode -ne 1 -or $state.Selection.StartLine -ne 10) {
            throw 'The inner procedure did not stop at line 10.'
        }
        $null = Invoke-Vbe @{ Command = 'open_debug_pane'; Action = 'locals' }
        $windows = Invoke-Vbe @{ Command = 'debug_windows'; IncludeCallStack = $true }
        if (-not $windows.CallStack.Available -or $windows.CallStack.Error -or
            @($windows.CallStack.Frames).Count -lt 2 -or
            -not (@($windows.CallStack.Frames) -match 'CodexStackInner') -or
            -not (@($windows.CallStack.Frames) -match 'CodexStackOuter')) {
            throw "Nested Call Stack not observed: $($windows.CallStack | ConvertTo-Json -Compress)"
        }

        [pscustomobject]@{
            HostProcessId = $HostProcessId
            BreakLine = $state.Selection.StartLine
            Frames = ($windows.CallStack.Frames -join ' | ')
            ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
        } | Format-List

    }
    'QuickWatch' {
        $ErrorActionPreference = 'Stop'
        $excel = Get-Process -Id $HostProcessId -ErrorAction Stop
        if ($excel.ProcessName -ne 'EXCEL') { throw 'A disposable Excel process is required.' }




        $initial = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
        if ($initial.Code.Length -ne 0) { throw 'Use a fresh workbook with an empty ThisWorkbook module.' }

        $probe = Join-Path $PSScriptRoot 'Test-Debugger.ps1'
        $null = & $probe -Scenario BreakpointStage -HostProcessId $HostProcessId -OwnedDisposableHost -Stage insert
        $null = & $probe -Scenario BreakpointStage -HostProcessId $HostProcessId -OwnedDisposableHost -Stage toggle
        $null = & $probe -Scenario BreakpointStage -HostProcessId $HostProcessId -OwnedDisposableHost -Stage run

        $state = $null
        for ($attempt = 0; $attempt -lt 20; $attempt++) {
            $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
            if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq 4) { break }
            Start-Sleep -Milliseconds 100
        }
        if ($state.Mode -ne 1 -or $state.Selection.StartLine -ne 4) {
            throw 'The disposable macro did not stop on line 4.'
        }

        $source = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
        $watch = Invoke-Vbe @{ Command = 'quick_watch'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; ExpectedMode = 1; StartLine = 4;
            StartColumn = 5; EndColumn = 15; Expression = 'probeValue';
            Procedure = 'CodexBreakpointProbe' }
        if ($watch.Expression -ne 'probeValue' -or $watch.Value -ne '1' -or
            $watch.Context -ne 'VBAProject.ThisWorkbook.CodexBreakpointProbe' -or
            $watch.Verification -ne 'NativeDialogReadback') {
            throw "Unexpected Quick Watch result: $($watch | ConvertTo-Json -Compress)"
        }

        [pscustomobject]@{
            HostProcessId = $HostProcessId
            Expression = $watch.Expression
            Value = $watch.Value
            Context = $watch.Context
            Verification = $watch.Verification
            ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
        } | Format-List

    }
    'EditWatchType' {
        $ErrorActionPreference = 'Stop'
        $excel = Get-Process -Id $HostProcessId -ErrorAction Stop
        if ($excel.ProcessName -ne 'EXCEL') { throw 'A disposable Excel process is required.' }



        $initial = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
        if ($initial.Code.Length -ne 0) { throw 'Use a fresh workbook with an empty ThisWorkbook module.' }
        $probe = Join-Path $PSScriptRoot 'Test-Debugger.ps1'
        $null = & $probe -Scenario BreakpointStage -HostProcessId $HostProcessId -OwnedDisposableHost -Stage insert
        $null = & $probe -Scenario BreakpointStage -HostProcessId $HostProcessId -OwnedDisposableHost -Stage toggle
        $null = & $probe -Scenario BreakpointStage -HostProcessId $HostProcessId -OwnedDisposableHost -Stage run

        $state = $null
        for ($attempt = 0; $attempt -lt 20; $attempt++) {
            $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
            if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq 4) { break }
            Start-Sleep -Milliseconds 100
        }
        if ($state.Mode -ne 1 -or $state.Selection.StartLine -ne 4) {
            throw 'The disposable macro did not stop on line 4.'
        }

        $expression = 'probeValue = 2'
        $context = 'ThisWorkbook.CodexBreakpointProbe'
        $null = Invoke-Vbe @{ Command = 'add_watch'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedMode = 1; Procedure = 'CodexBreakpointProbe'; Expression = $expression;
            WatchType = 'expression' }
        $before = Invoke-Vbe @{ Command = 'debug_windows' }
        $watchBefore = @($before.Watches.Items | Where-Object {
                $_.Expression -eq $expression -and $_.Context -eq $context })
        if ($watchBefore.Count -ne 1 -or $watchBefore[0].Value -notmatch 'Faux|False') {
            throw 'The initial expression watch was not read back as false.'
        }

        $edit = Invoke-Vbe @{ Command = 'edit_watch'; Project = 'VBAProject'; ExpectedMode = 1;
            Expression = $expression; NewExpression = $expression; Context = $context;
            WatchType = 'break_when_true' }
        if (-not $edit.Edited -or $edit.Verification -ne 'ReadbackVerified') {
            throw 'The edited watch was not read back.'
        }

        $source = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
        $commands = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'Continuer' })
        $control = $commands | Where-Object { $_.Id -eq 186 -and $_.Enabled } | Select-Object -First 1
        if (-not $control) { throw 'Native Continue command 186 is unavailable.' }
        $null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 4; ExpectedMode = 1;
            Action = 'continue'; ControlId = 186; ControlCaption = $control.Caption }

        $after = $null
        for ($attempt = 0; $attempt -lt 20; $attempt++) {
            $after = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
            if ($after.Mode -eq 1 -and $after.Selection.StartLine -eq 5) { break }
            Start-Sleep -Milliseconds 100
        }
        $windows = Invoke-Vbe @{ Command = 'debug_windows' }
        $watchAfter = @($windows.Watches.Items | Where-Object {
                $_.Expression -eq $expression -and $_.Context -eq $context })
        if ($after.Mode -ne 1 -or $after.Selection.StartLine -ne 5 -or
            $watchAfter.Count -ne 1 -or $watchAfter[0].Value -notmatch 'Vrai|True') {
            throw 'The edited break_when_true watch did not stop execution at line 5.'
        }

        [pscustomobject]@{
            HostProcessId = $HostProcessId
            Expression = $expression
            Before = $watchBefore[0].Value
            After = $watchAfter[0].Value
            BreakLine = $after.Selection.StartLine
            ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
        } | Format-List

    }
    'RuntimeDialog' {
        $ErrorActionPreference = 'Stop'
        $hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
        if ($hostProcess.ProcessName -ne 'EXCEL') { throw 'An isolated Excel process is required.' }



        $project = 'VBAProject'
        $module = 'ThisWorkbook'
        $before = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
        if ($before.Code.Length -ne 0) { throw 'Use a fresh disposable workbook with an empty ThisWorkbook module.' }
        $code = "Public Sub CodexRuntimeProbe()`r`n    Err.Raise 11`r`nEnd Sub"
        Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
            ExpectedSha256 = $before.Sha256; StartLine = 1; Count = 0; Text = $code } | Out-Null
        $source = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $module }
        Invoke-Vbe @{ Command = 'select_code'; Project = $project; Module = $module;
            ExpectedSha256 = $source.Sha256; StartLine = 1 } | Out-Null
        $control = @(Invoke-Vbe @{ Command = 'list_commands' } | Where-Object {
                $_.Id -eq 186 -and $_.Enabled -and $_.Caption -match 'Sub/UserForm'
        }) | Select-Object -First 1
        if (-not $control) { throw 'The native Run Sub/UserForm command is unavailable.' }
        Invoke-Vbe @{ Command = 'invoke_debug'; Project = $project; Module = $module;
            ExpectedSha256 = $source.Sha256; StartLine = 1; ExpectedMode = 2;
            Action = 'run'; ControlId = 186; ControlCaption = $control.Caption } | Out-Null

        $dialog = $null
        for ($attempt = 0; $attempt -lt 30; $attempt++) {
            $dialog = Invoke-Vbe @{ Command = 'debug_dialog' }
            if ($dialog.Available) { break }
            Start-Sleep -Milliseconds 100
        }
        if (-not $dialog.Available -or $dialog.Error -or
            $dialog.Diagnostic -notmatch "Erreur d'exécution '11'|Run-time error '11'" -or
            @($dialog.Buttons | Where-Object { $_ -eq 'OK' }).Count -ne 1) {
            throw 'The expected native run-time error dialog was not captured.'
        }
        $closed = Invoke-Vbe @{ Command = 'respond_debug_dialog';
            Diagnostic = $dialog.Diagnostic; Button = 'OK' }
        $after = Invoke-Vbe @{ Command = 'debug_dialog' }
        $state = Invoke-Vbe @{ Command = 'debug_state'; Project = $project }
        if ($closed.Verification -ne 'DialogClosed' -or $after.Available -or $state.Mode -ne 2)
        { throw 'The native diagnostic did not close into design mode.' }

        [pscustomobject]@{
            HostProcessId = $HostProcessId
            Diagnostic = $dialog.Diagnostic
            Button = $closed.Button
            ModeAfter = $state.Mode
            ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue)
        } | Format-List

    }
    'ShowNextStatementGuard' {
        $ErrorActionPreference = 'Stop'
        $excelProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
        if ($excelProcess.ProcessName -ne 'EXCEL') { throw 'A disposable Excel process is required.' }



        $initial = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
        $code = @'
Public Sub CodexShowProbe()
    Dim value As Long
    value = 1
    value = value + 1
    Debug.Print "CodexShowProbe:" & CStr(value)
End Sub
'@
        if ($initial.Code.Trim().Length -eq 0) {
            $null = Invoke-Vbe @{ Command = 'replace_lines'; Project = 'VBAProject'; Module = 'ThisWorkbook';
                ExpectedSha256 = $initial.Sha256; StartLine = 1; Count = 0; Text = $code }
        } elseif (($initial.Code.Trim() -replace "`r`n", "`n") -ne $code.Trim()) {
            throw 'Use a fresh workbook or the exact CodexShowProbe source.'
        }
        $source = Invoke-Vbe @{ Command = 'read_module'; Project = 'VBAProject'; Module = 'ThisWorkbook' }
        $null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 4 }
        $breakControl = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'point' }) |
        Where-Object { $_.Id -eq 51 -and $_.Enabled } | Select-Object -First 1
        if (-not $breakControl) { throw 'Native Toggle Breakpoint command 51 is unavailable.' }
        $null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 4; ExpectedMode = 2;
            Action = 'toggle_breakpoint'; ControlId = 51; ControlCaption = $breakControl.Caption }
        $runControl = @(Invoke-Vbe @{ Command = 'list_commands'; Query = 'Sub/UserForm' }) |
        Where-Object { $_.Id -eq 186 -and $_.Enabled } | Select-Object -First 1
        if (-not $runControl) { throw 'Native Run Sub command 186 is unavailable.' }
        $null = Invoke-Vbe @{ Command = 'invoke_debug'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 1; ExpectedMode = 2;
            Action = 'run'; ControlId = 186; ControlCaption = $runControl.Caption }
        $state = $null
        for ($attempt = 0; $attempt -lt 30; $attempt++) {
            $state = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
            if ($state.Mode -eq 1 -and $state.Selection.StartLine -eq 4) { break }
            Start-Sleep -Milliseconds 100
        }
        if ($state.Mode -ne 1 -or $state.Selection.StartLine -ne 4) {
            throw "Did not stop at line 4: $($state | ConvertTo-Json -Compress -Depth 5)"
        }
        $null = Invoke-Vbe @{ Command = 'select_code'; Project = 'VBAProject'; Module = 'ThisWorkbook';
            ExpectedSha256 = $source.Sha256; StartLine = 2 }
        $before = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
        if ($before.Selection.StartLine -ne 2) { throw 'Selection did not move away from the execution line.' }
        $show = Invoke-Vbe @{ Command = 'debug_global'; Project = 'VBAProject'; ExpectedMode = 1;
            Action = 'show_next_statement' }
        $after = $null
        for ($attempt = 0; $attempt -lt 30; $attempt++) {
            $after = Invoke-Vbe @{ Command = 'debug_state'; Project = 'VBAProject' }
            if ($after.Selection.StartLine -eq 4) { break }
            Start-Sleep -Milliseconds 100
        }
        if ($after.Mode -ne 1 -or $after.Selection.StartLine -ne 4) {
            throw "Show Next Statement did not restore line 4: $($after | ConvertTo-Json -Compress -Depth 5)"
        }
        $null = Invoke-Vbe @{ Command = 'debug_global'; Project = 'VBAProject'; ExpectedMode = 1;
            Action = 'reset' }
        [pscustomobject]@{
            HostProcessId = $HostProcessId
            BreakLine = 4
            MovedSelectionLine = $before.Selection.StartLine
            RestoredSelectionLine = $after.Selection.StartLine
            ShowControlId = $show.ControlId
        } | Format-List

    }
    'RegisteredTree' {
        $ErrorActionPreference='Stop'
        . $VbeProbeInitializeRegistered
        $excel=$null;$book=$null;$vbe=$null;$probeProcess=$null;$toolbarName=$null;$option=$null;$optionChanged=$false;$report=[ordered]@{}



        try{
            if($initialAccess -ne 1){New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value 1 -PropertyType DWord -Force|Out-Null}
            $report.Identity=Open-ProbeExcel
            $book=$excel.Workbooks.Add();$document=Join-Path $outputRoot ('Debugger-'+[Guid]::NewGuid().ToString('N')+'.xlsm');$book.SaveAs($document,52)
            $module=$book.VBProject.VBComponents.Add(1);$module.Name='LargeTreeProbe'
            $module.CodeModule.AddFromString(@"
Option Explicit
Public Sub InspectLargeTree()
    Dim values(1 To 1000) As Long
    Dim number As Double, money As Currency, stamp As Date
    Dim text As String, flag As Boolean, absent As Variant, index As Long
    For index = 1 To 1000
        values(index) = index * 3
    Next index
    number = 12.5: money = 7.25: stamp = DateSerial(2026, 9, 28)
    text = "VBAi debugger": flag = True: absent = Null
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = "Paused"
    Stop
    ThisWorkbook.Worksheets(1).Range("A1").Value2 = values(1000)
End Sub
"@)
            Invoke-Bridge @{Command='open_debug_pane';Action='locals'}|Out-Null
            $source=Invoke-Bridge @{Command='read_module';Project=$document;Module=$module.Name}
            $report.Start=Invoke-Bridge @{Command='run_sub';Project=$document;Module=$module.Name;Procedure='InspectLargeTree';ExpectedMode=2;ExpectedSha256=$source.Sha256}
            for($i=0;$i -lt 30;$i++){Start-Sleep -Milliseconds 200;$mode=Invoke-Bridge @{Command='debug_state';Project=$document};if($mode.Mode -eq 1){break}}
            if($mode.Mode -ne 1 -or $book.Worksheets.Item(1).Range('A1').Value2 -ne 'Paused'){throw 'Owned fixture did not pause at Stop.'}
            $before=Invoke-Bridge @{Command='debug_windows'};$report.Before=$before
            $row=@($before.Locals.Items|Where-Object Expression -eq 'values')
            if($row.Count -ne 1 -or @($row[0].PathSegments).Count -eq 0){throw 'Large-array root was not uniquely exposed.'}
            $report.Expand=Invoke-Bridge @{Command='debug_item';Pane='locals';Action='expand';PathSegments=@($row[0].PathSegments)}
            $expanded=Invoke-Bridge @{Command='debug_windows'};$report.Expanded=$expanded
            $children=@($expanded.Locals.Items|Where-Object { $_.PathSegments.Count -gt $row[0].PathSegments.Count -and $_.PathSegments[0] -eq $row[0].PathSegments[0] })
            if($children.Count -eq 0){throw 'No array child exposed after expansion.'}
            $verifiedValues=0
            foreach($child in $children){
                if($child.Expression -notmatch '^values\(([0-9]+)\)$' -or [int]$Matches[1] -lt 1 -or [int]$Matches[1] -gt 1000 -or $child.Value -ne [string]([int]$Matches[1]*3)){throw 'Array child value differs from its deterministic fixture.'}
                $verifiedValues++
            }
            if(@($children|Group-Object Expression|Where-Object Count -ne 1).Count){throw 'Duplicate array observations remain.'}
            $report.Array=@{ExpectedTotal=1000;ExposedChildren=$children.Count;VerifiedValues=$verifiedValues;Coverage=$expanded.Locals.Coverage;Exhaustive=($children.Count -eq 1000)}
            $report.Collapse=Invoke-Bridge @{Command='debug_item';Pane='locals';Action='collapse';PathSegments=@($row[0].PathSegments)}
            $collapsed=Invoke-Bridge @{Command='debug_windows'};$report.Collapsed=$collapsed
            if(@($collapsed.Locals.Items|Where-Object { $_.PathSegments.Count -gt $row[0].PathSegments.Count -and $_.PathSegments[0] -eq $row[0].PathSegments[0] }).Count){throw 'Array children remain exposed after collapse.'}
            foreach($name in @('number','money','stamp','text','flag','absent','index')){
                if(@($before.Locals.Items|Where-Object { $_.Expression -eq $name -and $_.Type -and $null -ne $_.Value }).Count -ne 1){throw ('Scalar row missing: '+$name)}
            }
            $resume=$vbe.CommandBars.FindControl(1,186)
            $report.Resume=Invoke-Bridge @{Command='invoke_debug';Project=$document;Module=$module.Name;ExpectedMode=1;ExpectedSha256=$source.Sha256;StartLine=13;Action='continue';ControlId=186;ControlCaption=$resume.Caption}
            Start-Sleep -Milliseconds 300
            $after=Invoke-Bridge @{Command='debug_state';Project=$document}
            if($after.Mode -ne 2 -or $book.Worksheets.Item(1).Range('A1').Value2 -ne 3000){throw 'Fixture did not resume exactly once.'}
            $read=Invoke-Bridge @{Command='read_module';Project=$document;Module=$module.Name}
            if($read.Sha256 -ne $source.Sha256){throw 'Debugger actions altered the source.'}
            $report.RuntimeMarker=3000;$report.CodeUnchanged=$true
        }catch{$report|ConvertTo-Json -Depth 30|Set-Content -LiteralPath (Join-Path $outputRoot 'partial-results.json') -Encoding UTF8;throw}
        finally{
            try{
                if($null -ne $book){$state=Invoke-Bridge @{Command='debug_state';Project=$document};if($state.Mode -eq 1){Invoke-Bridge @{Command='debug_global';Project=$document;ExpectedMode=1;Action='reset'}|Out-Null}}
                Close-ProbeExcel
            }finally{
                if($hadAccess){New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value $initialAccess -PropertyType DWord -Force|Out-Null}else{Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue}
            }
        }
        $report|ConvertTo-Json -Depth 30|Set-Content -LiteralPath (Join-Path $outputRoot 'registered-debugger-tree.json') -Encoding UTF8
        Write-Output 'PASS native large array expand/collapse, scalar types and resume; exposed-row count recorded.'

    }
}

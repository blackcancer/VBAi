[CmdletBinding()]
param([string]$LauncherPath)
if([string]::IsNullOrWhiteSpace($LauncherPath)){$LauncherPath=Join-Path $PSScriptRoot 'Invoke-IsolatedTests.ps1'}
$ErrorActionPreference='Stop'
# Execute the real generated catch/finally with finite injected process/writer/wait behavior.
# No helper, Office, registry write or native wait is invoked.
$source=[IO.File]::ReadAllText((Resolve-Path -LiteralPath $LauncherPath).Path)
function Get-RealCatchFinally([string]$text){
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseInput($text,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw 'Launcher parse failed'}
    $assignment=@($ast.FindAll({param($a) $a -is [Management.Automation.Language.AssignmentStatementAst] -and $a.Left.Extent.Text -eq '$body'},$true))
    if($assignment.Count -ne 1){throw 'Generated body assignment missing'}
    $literal=@($assignment[0].Right.FindAll({param($a) $a -is [Management.Automation.Language.StringConstantExpressionAst]},$true))
    if($literal.Count -ne 1){throw 'Exact literal generated body required'}
    $body=$literal[0].Value
    $bodyAst=[Management.Automation.Language.Parser]::ParseInput($body,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw 'Generated body parse failed'}
    $target=@($bodyAst.FindAll({param($a) $a -is [Management.Automation.Language.TryStatementAst] -and $a.Finally -and $a.Finally.Extent.Text.Contains('$process.Dispose()')},$true))
    if($target.Count -ne 1 -or $target[0].CatchClauses.Count -ne 1){throw 'Exact process catch/finally required'}
    $code='try { throw "Injected original Start/receipt failure" } '+$target[0].CatchClauses[0].Extent.Text+' finally '+$target[0].Finally.Extent.Text
    # Only read-only wait and delay are injected; actual catch, writer, guard and Dispose code remain extracted.
    $wait='[IsolatedHelperProcessNative]::WaitForSingleObject($originalHandle,250)'
    $sleep='[Threading.Thread]::Sleep(250)'
    if(-not $code.Contains($wait) -or -not $code.Contains($sleep)){throw 'Exact wait/sleep sites required'}
    $code=$code.Replace($wait,'(Invoke-ProbeWait $originalHandle 250)').Replace($sleep,'Invoke-ProbeSleep 250')
    # A dynamically compiled scriptblock has no PSScriptRoot. Inject only its receipt destination.
    $code=$code.Replace("(Join-Path `$PSScriptRoot 'launcher-retained.json')","'probe-launcher-retained.json'")
    return $code
}

$variants=@{Current=(Get-RealCatchFinally $source)}
function Run-Probe([string]$variant,[string]$case,[bool]$startedValue,[bool]$enteredValue,[bool]$returnedValue,[bool]$writerFails,[bool]$observeExit){
    $script:trace=[Collections.Generic.List[string]]::new();$script:waits=0
    $script:writerFails=$writerFails;$script:observeExit=$observeExit
    function Write-Durable([string]$path,$value){
        if($path.EndsWith('launcher-retained.json')){
            $script:trace.Add('RetainedReceiptAttempt')
            if($script:writerFails){throw 'Injected retained receipt IO failure'}
        }else{$script:trace.Add('Terminal')}
    }
    function Invoke-ProbeWait([IntPtr]$handle,[uint32]$timeout){
        if($handle -ne [IntPtr]42 -or $timeout -ne 250){throw 'Unexpected wait parameters'}
        $script:waits++;$script:trace.Add('Wait'+$script:waits)
        if($script:waits -eq 1){return [uint32]258}
        if(-not $script:observeExit){throw 'Injected finite probe stop: exit remains unobserved'}
        $script:trace.Add('ExitObserved');return [uint32]0
    }
    function Invoke-ProbeSleep([int]$milliseconds){$script:trace.Add('SleepInjected')}
    $process=[pscustomobject]@{Handle=[IntPtr]42;Id=7}
    $process|Add-Member ScriptMethod Dispose {$script:trace.Add('Dispose')}
    $plan=[pscustomobject]@{Terminal='probe-terminal.json'}
    $started=$startedValue;$startEntered=$enteredValue;$startReturned=$returnedValue
    $observedExit=$false;$originalHandle=[IntPtr]42;$failure=$null;$code=0;$escaped=$null
    try{. ([scriptblock]::Create($variants[$variant]))}catch{$escaped=$_.Exception.Message}
    [pscustomobject]@{Variant=$variant;Case=$case;Trace=$script:trace.ToArray();ObservedExit=$observedExit;Escaped=$escaped;Failure=$failure;WaitCount=$script:waits}
}
$results=@()
foreach($variant in @('Current')){
    $results+=Run-Probe $variant 'Started_ReceiptFails_ThenExit' $true $true $true $true $true
    $results+=Run-Probe $variant 'UnknownStart_ReceiptFails_ExitUnobserved' $false $true $false $true $false
    $results+=Run-Probe $variant 'UnknownStart_ReceiptSucceeds_ExitUnobserved' $false $true $false $false $false
    $results+=Run-Probe $variant 'Started_ReceiptSucceeds_ThenExit' $true $true $true $false $true
    $results+=Run-Probe $variant 'NeverStarted' $false $false $false $false $false
    $results+=Run-Probe $variant 'StartReturnedFalse' $false $true $true $false $false
}
foreach($r in $results){
    if($r.Case -notin @('NeverStarted','StartReturnedFalse') -and $r.Trace -notcontains 'RetainedReceiptAttempt'){throw 'Injected writer was not reached'}
    if($r.Variant -eq 'Current'){
        if($r.Case -like '*ThenExit'){
            if($r.WaitCount -ne 2 -or -not $r.ObservedExit -or $r.Escaped -or [array]::IndexOf($r.Trace,'Dispose') -le [array]::IndexOf($r.Trace,'ExitObserved')){throw 'Proposed wait/exit ordering failed'}
        }elseif($r.Case -like '*ExitUnobserved'){
            if($r.WaitCount -ne 2 -or $r.ObservedExit -or $r.Trace -contains 'Dispose' -or $r.Trace -contains 'Terminal'){throw 'Proposed unknown lifetime released'}
        }else{if($r.WaitCount -ne 0 -or $r.Trace -notcontains 'Dispose' -or $r.Trace -notcontains 'Terminal'){throw 'Known no-child path changed'}}
    }
}
[pscustomobject]@{Passed=$true;Cases=$results.Count;Scope='Pure generated launcher retention fault injection'}
$results|Select-Object Case,WaitCount,ObservedExit,@{n='Trace';e={$_.Trace -join ' > '}}
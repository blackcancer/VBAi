#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [Parameter(Mandatory=$true)][string]$RecoveryRoot,
    [Parameter(Mandatory=$true)][string]$ClosureReceipt,
    [string]$RecoveryHostStartup)
$ErrorActionPreference='Stop'
if(-not [IO.Path]::IsPathRooted($RecoveryRoot) -or (Test-Path -LiteralPath $RecoveryRoot)){throw 'A fresh one-shot recovery root is required.'}
$plan=Get-Content -LiteralPath (Join-Path $EvidenceRoot 'q026-plan.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$phaseRoots=@(Get-ChildItem -LiteralPath (Join-Path $EvidenceRoot 'native/phases') -Directory)
if($phaseRoots.Count -ne 1){throw 'Exactly one native phase directory is required.'}
$retained=@(Get-ChildItem -LiteralPath $phaseRoots[0].FullName -Filter '*-HostRetained-*.json')
$baselineFile=@(Get-ChildItem -LiteralPath $phaseRoots[0].FullName -Filter '*-BaselineComplete-*.json')
if($retained.Count -ne 1 -or $baselineFile.Count -ne 1){throw 'Complete baseline and committed-entry ledger are required.'}
$ledger=Get-Content -LiteralPath $retained[0].FullName -Raw -Encoding UTF8 | ConvertFrom-Json
$baseline=Get-Content -LiteralPath $baselineFile[0].FullName -Raw -Encoding UTF8 | ConvertFrom-Json
$closed=Get-Content -LiteralPath $ClosureReceipt -Raw -Encoding UTF8 | ConvertFrom-Json
$originalProcessId=$ledger.ProcessId
$originalStartUtc=$ledger.ProcessStartUtc
if($RecoveryHostStartup){
    # An expired read cannot be replayed in the original process. A separately owned
    # recovery host may compensate only a proven font commit after that process died.
    $replacement=Get-Content -LiteralPath $RecoveryHostStartup -Raw -Encoding UTF8 | ConvertFrom-Json
    if($ledger.Data.Request.Command -cne 'read_vbe_options' -or $ledger.Data.Response -ne $null -or
        -not $replacement.Owned -or $replacement.LoadedAssemblyMvid -ne $plan.ProductMvid -or
        $replacement.ProcessId -eq $originalProcessId){throw 'Separate owned recovery host is allowed only after a failed read, on the exact frozen product.'}
    if(Get-Process -Id $originalProcessId -ErrorAction SilentlyContinue){throw 'Original host still exists; no successor recovery.'}
    $ledger.ProcessId=$replacement.ProcessId
    $ledger.ProcessStartUtc=$replacement.HostStartedUtc
}
if(-not $closed.EnumerationSucceeded -or -not $closed.OptionsDialogAbsent -or $closed.ProcessId -ne $ledger.ProcessId -or
    $closed.ProcessStartUtc -ne $ledger.ProcessStartUtc -or $baseline.ProductMvid -ne $plan.ProductMvid){throw 'Exact owned closed-dialog and frozen-candidate identity required.'}
[IO.Directory]::CreateDirectory($RecoveryRoot) | Out-Null
$bridge=Join-Path $PSScriptRoot '../Invoke-VBAi.ps1'
function Save($name,$value){$value | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath (Join-Path $RecoveryRoot $name) -Encoding UTF8}
function Guard {
    $p=Get-Process -Id $ledger.ProcessId -ErrorAction Stop
    if($p.ProcessName -ne 'EXCEL' -or $p.StartTime.ToUniversalTime() -ne [DateTime]::Parse($ledger.ProcessStartUtc).ToUniversalTime()){throw 'Owned PID/start changed; no dispatch.'}
    if((Get-FileHash -LiteralPath $plan.InstalledProduct).Hash -cne $plan.ProductSha256){throw 'Frozen installed candidate changed.'}
    $others=@(Get-Process EXCEL,WINWORD,POWERPNT,MSACCESS,MSPUB,SLDWORKS -ErrorAction SilentlyContinue | Where-Object {$_.Id -ne $ledger.ProcessId})
    if($others.Count){throw 'Competing VBE host appeared; retain without dispatch.'}
}
function Read($name){
    Guard;Save ($name+'-intent.json') @{Command='read_vbe_options';Utc=[DateTime]::UtcNow.ToString('o');InvocationCount=1}
    $raw=& $bridge -HostProcessId $ledger.ProcessId -ResponseTimeoutSeconds 90 -RequestJson '{"Command":"read_vbe_options"}'
    $raw | Set-Content -LiteralPath (Join-Path $RecoveryRoot ($name+'-reply.json')) -Encoding UTF8
    $reply=$raw | ConvertFrom-Json
    if(-not $reply.Ok -or -not $reply.Data.DialogClosed -or $reply.Data.OptionsVersion -notmatch '^[0-9a-fA-F]{64}$' -or -not $reply.Data.Tabs){throw 'Complete terminal read not verified; retain without replay.'}
    return $reply.Data
}
function Control($state,$entry){
    $tab=@($state.Tabs | Where-Object {$_.Tab -ceq $entry.Pane})
    if($tab.Count -ne 1){throw 'Exact tab unavailable.'}
    if($entry.Category){$category=@($tab[0].FormatCategories | Where-Object {$_.Category -ceq $entry.Category});if($category.Count -ne 1){throw 'Exact category unavailable.'};$controls=$category[0].Palettes}
    else{$controls=$tab[0].Controls}
    $control=@($controls | Where-Object {$_.Name -ceq $entry.Property -and $_.Type -ne 'ControlType.Text'})
    if($control.Count -ne 1 -or $control[0].Error){throw 'Exact readable control unavailable.'}
    return $control[0]
}
Save 'claim.json' @{State='EXPLICIT_RECOVERY_ONCE';ProcessId=$ledger.ProcessId;ProcessStartUtc=$ledger.ProcessStartUtc;
    ProductMvid=$plan.ProductMvid;OriginalProcessId=$originalProcessId;OriginalProcessStartUtc=$originalStartUtc;
    SeparateOwnedRecoveryHost=[bool]$RecoveryHostStartup;FailedMutationReplayed=$false;NativeQualification='FAILED';CleanupAllowed=$false}
try {
    Guard
    $statusRaw=& $bridge -HostProcessId $ledger.ProcessId -ResponseTimeoutSeconds 20 -Command status
    $statusRaw | Set-Content -LiteralPath (Join-Path $RecoveryRoot 'status.json') -Encoding UTF8
    $status=$statusRaw | ConvertFrom-Json
    if(-not $status.Ok -or $status.Data.HostProcessId -ne $ledger.ProcessId -or $status.Data.AssemblyModuleVersionId -ne $plan.ProductMvid){throw 'Loaded candidate identity mismatch.'}
    $current=Read 'initial'
    if($RecoveryHostStartup){
        $fontEntries=@($ledger.Data.CommittedRestoreEntries | Where-Object {$_.Property -cin @('Font','Police :') -and -not $_.Category})
        if($fontEntries.Count -ne 1){throw 'One positively committed original font entry is required.'}
        $normalized=$current | ConvertTo-Json -Depth 100 -Compress | ConvertFrom-Json
        $font=Control $normalized $fontEntries[0]
        if($font.Value -cnotin @('Courier New','Courier New (Occidental)','Courier New (Western)','Consolas','Consolas (Occidental)','Consolas (Western)')){throw 'Unexpected font state; no compensation.'}
        $font.Value=$fontEntries[0].Value
        if(($normalized.Tabs | ConvertTo-Json -Depth 100 -Compress) -cne ($baseline.Data.Tabs | ConvertTo-Json -Depth 100 -Compress)){throw 'State differs from baseline beyond the positively committed font; no compensation.'}
        Save 'separate-host-precheck.json' @{FailedCommand='read_vbe_options';OriginalHostAbsent=$true;
            BaselineEqualAfterFontNormalization=$true;OriginalBaselineVersion=$baseline.Data.OptionsVersion;Readback=$current}
    } elseif($null -eq $ledger.Data.Request -and @($ledger.Data.CommittedRestoreEntries).Count -eq 0){
        $isolationFiles=@(Get-ChildItem -LiteralPath $phaseRoots[0].FullName -Filter '*-HostExclusivityObservation-*.json' | Sort-Object Name)
        if(-not $isolationFiles.Count){throw 'No known isolation-stop observation; no recovery dispatch.'}
        $isolation=Get-Content -LiteralPath $isolationFiles[-1].FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        if($isolation.Sequence+1 -ne $ledger.Sequence -or -not $isolation.Data.OwnedAlive -or
            @($isolation.Data.Competitors).Count -eq 0 -or -not $isolation.Data.BeforeDispatch -or
            $ledger.Data.Error -notmatch '^System.InvalidOperationException: The owned Excel identity or exclusive VBE-host interval changed;' -or
            $current.OptionsVersion -cne $baseline.Data.OptionsVersion -or
            ($current.Tabs | ConvertTo-Json -Depth 100 -Compress) -cne ($baseline.Data.Tabs | ConvertTo-Json -Depth 100 -Compress)){
            throw 'Known stop before dispatch and complete unchanged baseline required; no compensation.'
        }
        Save 'isolation-stop-precheck.json' @{BeforeDispatch=$true;CommittedEntries=0;CompleteBaselineEqual=$true;CompensationWrites=0}
    } else {
        $failedControl=@($current.Tabs | Where-Object {$_.Tab -ceq $ledger.Data.Request.Pane} | ForEach-Object {$_.Controls} | Where-Object {$_.Name -ceq $ledger.Data.Request.Property -and $_.Type -ne 'ControlType.Text'})
        $baselineControl=@($baseline.Data.Tabs | Where-Object {$_.Tab -ceq $ledger.Data.Request.Pane} | ForEach-Object {$_.Controls} | Where-Object {$_.Name -ceq $ledger.Data.Request.Property -and $_.Type -ne 'ControlType.Text'})
        if($failedControl.Count -ne 1 -or $baselineControl.Count -ne 1 -or $failedControl[0].Value -cne $baselineControl[0].Value){throw 'Failed transaction control is not at baseline; retain for explicit recovery.'}
    }
    $entries=@($ledger.Data.CommittedRestoreEntries);[Array]::Reverse($entries);$sequence=0
    foreach($entry in $entries){
        $sequence++;$name='restore-'+$sequence.ToString('D2');$control=Control $current $entry
        if($control.Value -ceq $entry.Value){Save ($name+'-already-matched.json') $entry;continue}
        Guard
        $request=@{Command='set_vbe_option';Pane=$entry.Pane;Property=$entry.Property;Value=$entry.Value;Query=$entry.Category;ExpectedOptionsVersion=$current.OptionsVersion}
        Save ($name+'-intent.json') @{Request=$request;Before=$current;InvocationCount=1;Utc=[DateTime]::UtcNow.ToString('o')}
        $raw=& $bridge -HostProcessId $ledger.ProcessId -ResponseTimeoutSeconds 90 -RequestJson ($request | ConvertTo-Json -Compress)
        $raw | Set-Content -LiteralPath (Join-Path $RecoveryRoot ($name+'-reply.json')) -Encoding UTF8
        $reply=$raw | ConvertFrom-Json
        if(-not $reply.Ok -or -not $reply.Data.CommitRequested -or -not $reply.Data.ControlValueVerified -or -not $reply.Data.DialogClosed){throw 'Restoration outcome not verified; retain without replay.'}
        $current=Read ($name+'-readback')
        if((Control $current $entry).Value -cne $entry.Value){throw 'Exact restoration readback mismatch; retain without replay.'}
        Save ($name+'-verified.json') @{Entry=$entry;Revision=$current.OptionsVersion;InvocationCount=1}
    }
    if($current.OptionsVersion -cne $baseline.Data.OptionsVersion -or
        ($current.Tabs | ConvertTo-Json -Depth 100 -Compress) -cne ($baseline.Data.Tabs | ConvertTo-Json -Depth 100 -Compress)){throw 'Complete baseline revision and Tabs not restored.'}
    Save 'terminal.json' @{State='PREFERENCES_RECOVERED';Qualified=$false;ProcessId=$ledger.ProcessId;ProcessStartUtc=$ledger.ProcessStartUtc;
        OriginalProcessId=$originalProcessId;SeparateOwnedRecoveryHost=[bool]$RecoveryHostStartup;
        BaselineVersion=$baseline.Data.OptionsVersion;Readback=$current;FailedMutationReplayed=$false;HostShutdownInvoked=$false}
} catch {Save 'failure.json' @{State='RECOVERY_STOPPED_RETAINED';Error=$_.Exception.ToString();ReplayAllowed=$false};throw}

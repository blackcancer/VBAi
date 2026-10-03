#requires -Version 5.1
# Pure evidence validation. No host, bridge, registry or UI access.
function Resolve-Q026RecoveryBaseline {
    param([Parameter(Mandatory=$true)][object[]]$Records,
        [Parameter(Mandatory=$true)]$Ledger,
        [Parameter(Mandatory=$true)][string]$ProductMvid)
    $baseline=@($Records | Where-Object {$_.Phase -ceq 'BaselineComplete'})
    if($baseline.Count -eq 1){return @{Record=$baseline[0];Kind='ScenarioBaseline'}}
    if($baseline.Count -ne 0){throw 'Ambiguous scenario baseline.'}
    if($null -ne $Ledger.Data.Request -or $null -ne $Ledger.Data.Response -or
        @($Ledger.Data.CommittedRestoreEntries).Count -ne 0 -or
        $Ledger.Data.Error -notmatch '^System.InvalidOperationException: The owned Excel identity or exclusive VBE-host interval changed;'){
        throw 'A warmup baseline is allowed only for a known pre-dispatch isolation stop with no commits.'
    }
    $before=@($Records | Where-Object {$_.Phase -ceq 'GuardBreakpointWarmupBeforeReply'})
    $refusal=@($Records | Where-Object {$_.Phase -ceq 'GuardBreakpointWarmupVerifiedRefusal'})
    $verified=@($Records | Where-Object {$_.Phase -ceq 'GuardBreakpointWarmupVerified'})
    $isolation=@($Records | Where-Object {$_.Phase -ceq 'HostExclusivityObservation' -and $_.Sequence -eq $Ledger.Sequence-1})
    if($before.Count -ne 1 -or $refusal.Count -ne 1 -or $verified.Count -ne 1 -or $isolation.Count -ne 1){
        throw 'Complete unique warmup and immediate isolation-stop evidence required.'
    }
    foreach($record in @($before[0],$refusal[0],$verified[0],$isolation[0],$Ledger)){
        if($record.ProcessId -ne $Ledger.ProcessId -or $record.ProcessStartUtc -cne $Ledger.ProcessStartUtc -or
            $record.ProductMvid -cne $ProductMvid){throw 'Warmup evidence ownership mismatch.'}
    }
    if($before[0].Sequence -ge $refusal[0].Sequence -or $refusal[0].Sequence -ge $verified[0].Sequence -or
        $verified[0].Sequence -ge $isolation[0].Sequence -or
        $verified[0].Data.NativePreferenceWrites -ne 0 -or $verified[0].Data.FailedMutationReplayed -ne $false -or
        -not $isolation[0].Data.OwnedAlive -or -not $isolation[0].Data.BeforeDispatch -or
        @($isolation[0].Data.Competitors).Count -eq 0){throw 'Warmup or isolation stop was not fully verified.'}
    $intervening=@($Records | Where-Object {$_.Sequence -gt $verified[0].Sequence -and $_.Sequence -lt $Ledger.Sequence})
    if(@($intervening | Where-Object {$_.Phase -cnotin @('ScenarioMatrix','HostExclusivityObservation')}).Count){
        throw 'A native dispatch may have occurred after warmup; no warmup-baseline recovery.'
    }
    $response=$before[0].Data.Response
    $old=$response.Data
    $after=$refusal[0].Data.IndependentClosedReadback
    if($response.Ok -ne $true -or $old.DialogClosed -ne $true -or $after.DialogClosed -ne $true -or
        $old.OptionsVersion -notmatch '^[a-fA-F0-9]{64}$' -or $old.OptionsVersion -cne $after.OptionsVersion -or
        -not $old.Tabs -or -not $after.Tabs -or $refusal[0].Data.Response.Ok -ne $false -or
        $refusal[0].Data.Response.Error -cne 'VBE options changed since inspection; read them again.' -or
        $null -ne $refusal[0].Data.Response.Data -or $refusal[0].Data.MutationRetried -ne $false){
        throw 'Complete terminal unchanged warmup readback required.'
    }
    foreach($state in @($response,$old,$after,$refusal[0].Data.Response)){
        foreach($name in @('Pending','Uncertain','DeliveryUncertain','VerificationPending')){
            $value=$state.$name
            if($null -ne $value -and $value -ne $false){throw 'Uncertain warmup evidence cannot authorize recovery.'}
        }
    }
    if((ConvertTo-Json -InputObject $old.Tabs -Depth 100 -Compress) -cne
        (ConvertTo-Json -InputObject $after.Tabs -Depth 100 -Compress)){throw 'Warmup structures differ despite their revision.'}
    # Keep the original receipt intact; provide the verified readback as the recovery oracle.
    return @{Kind='VerifiedWarmupBeforeScenario';Record=[pscustomobject]@{
        ProductMvid=$ProductMvid;ProcessId=$Ledger.ProcessId;ProcessStartUtc=$Ledger.ProcessStartUtc;
        Data=$after;SourcePhase=$refusal[0].Phase;SourceSequence=$refusal[0].Sequence}}
}

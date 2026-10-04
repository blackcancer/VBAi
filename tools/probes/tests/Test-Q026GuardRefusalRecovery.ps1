#requires -Version 5.1
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../Resolve-Q026RecoveryBaseline.ps1')
$mvid='5cc513d1-5569-4835-bf6c-cf70a18274fb'
function Fixture {
    $request=@{Command='set_vbe_option';Pane='Editor Format';Property='__Q026_READ_ONLY_GUARD_WARMUP__';Value='No preference mutation';ExpectedOptionsVersion=('0'*64)}
    $reply=@{Ok=$false;Error='VBE options changed since inspection; read them again.';Data=$null}
    $state=@{DialogClosed=$true;OptionsVersion=('a'*64);Tabs=@(@{Tab='Editor Format';Controls=@(@{Name='Font';Value='Consolas'})})}
    $records=@(
        @{Phase='GuardBreakpointWarmupBeforeIntent';Sequence=2;Data=@{Command='read_vbe_options'}},
        @{Phase='GuardBreakpointWarmupBeforeReply';Sequence=3;Data=@{Response=@{Ok=$true;Data=$state}}},
        @{Phase='GuardBreakpointWarmupIntent';Sequence=5;Data=$request},
        @{Phase='GuardBreakpointWarmupRefusalReply';Sequence=6;Data=$reply},
        @{Phase='ClosureObservation';Sequence=7;Data=@{}},
        @{Phase='GuardBreakpointWarmupNativeClosureObservation';Sequence=8;Data=@{ProcessIdentityVerified=$true;ObservationOnly=$true;EnumerationSucceeded=$false;OptionsDialogAbsent=$false}},
        @{Phase='HostRetained';Sequence=9;Data=@{Request=$request;Response=$reply;CommittedRestoreEntries=@();Error='System.InvalidOperationException: Native Options-window absence for the exact owned process was not independently verified.'}}
    )
    foreach($record in $records){$record.ProcessId=1234;$record.ProcessStartUtc='2026-10-04T09:00:00Z';$record.ProductMvid=$mvid}
    return ,($records|ConvertTo-Json -Depth 30|ConvertFrom-Json)
}
$passed=0
function Refuses($name,[scriptblock]$change){
    $r=Fixture; & $change $r; $refused=$false
    try{$null=Resolve-Q026RecoveryBaseline -Records $r -Ledger $r[-1] -ProductMvid $mvid}catch{$refused=$true}
    if(-not $refused){throw ('Unsafe guard recovery accepted: '+$name)}
    $script:passed++
}
$r=Fixture
$result=Resolve-Q026RecoveryBaseline -Records $r -Ledger $r[-1] -ProductMvid $mvid
if($result.Kind-cne 'KnownWarmupRefusalBeforeScenario' -or $result.Record.Data.OptionsVersion-cne ('a'*64)){throw 'Known guard refusal rejected.'}
$passed++
Refuses 'foreign process' {param($r)$r[0].ProcessId=5678}
Refuses 'foreign start' {param($r)$r[1].ProcessStartUtc='2026-10-04T09:01:00Z'}
Refuses 'foreign product' {param($r)$r[2].ProductMvid='other'}
Refuses 'positive commit' {param($r)$r[-1].Data.CommittedRestoreEntries=@(@{Property='Font'})}
Refuses 'real property' {param($r)$r[2].Data.Property='Font';$r[-1].Data.Request.Property='Font'}
Refuses 'changed expected revision' {param($r)$r[2].Data.ExpectedOptionsVersion=('a'*64)}
Refuses 'different request in ledger' {param($r)$r[-1].Data.Request.Pane='Other'}
Refuses 'unexpected value' {param($r)$r[2].Data.Value='12'}
Refuses 'unexpected command' {param($r)$r[2].Data.Command='write_module'}
Refuses 'different refusal in ledger' {param($r)$r[-1].Data.Response.Error='other'}
Refuses 'successful response' {param($r)$r[3].Data.Ok=$true}
Refuses 'ambiguous response data' {param($r)$r[3].Data.Data=@{DialogClosed=$true}}
Refuses 'uncertain response' {param($r)$r[3].Data|Add-Member DeliveryUncertain $true}
Refuses 'pending baseline' {param($r)$r[1].Data.Response|Add-Member Pending $true}
Refuses 'missing tabs' {param($r)$r[1].Data.Response.Data.Tabs=$null}
Refuses 'unverified initial closure' {param($r)$r[1].Data.Response.Data.DialogClosed=$false}
Refuses 'zero baseline revision' {param($r)$r[1].Data.Response.Data.OptionsVersion=('0'*64)}
Refuses 'ownership not verified' {param($r)$r[5].Data.ProcessIdentityVerified=$false}
Refuses 'closure was an action' {param($r)$r[5].Data.ObservationOnly=$false}
Refuses 'wrong phase order' {param($r)$r[5].Sequence=10}
Refuses 'duplicate baseline' {param($r)$r[4].Phase='GuardBreakpointWarmupBeforeReply'}
Refuses 'other native intent' {param($r)$r[4].Phase='FontIntent'}
Refuses 'wrong baseline read' {param($r)$r[0].Data.Command='set_vbe_option'}
Refuses 'different retained error' {param($r)$r[-1].Data.Error='Unknown delivery'}
[pscustomobject]@{State='PASSED';Passed=$passed;Failed=0;NativeDispatches=0;OfficeActivated=$false}|ConvertTo-Json

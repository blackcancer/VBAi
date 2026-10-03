#requires -Version 5.1
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../Resolve-Q026RecoveryBaseline.ps1')
function Fixture {
    $state=@{DialogClosed=$true;OptionsVersion=('a'*64);Tabs=@(@{Tab='Editor Format';Controls=@(@{Name='Font';Value='Consolas'})})}
    $after=$state | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $records=@(
        @{Phase='GuardBreakpointWarmupBeforeReply';Sequence=3;Data=@{Response=@{Ok=$true;Data=$state}}},
        @{Phase='GuardBreakpointWarmupVerifiedRefusal';Sequence=12;Data=@{Response=@{Ok=$false;Error='VBE options changed since inspection; read them again.';Data=$null};IndependentClosedReadback=$after;MutationRetried=$false}},
        @{Phase='GuardBreakpointWarmupVerified';Sequence=13;Data=@{NativePreferenceWrites=0;FailedMutationReplayed=$false}},
        @{Phase='ScenarioMatrix';Sequence=14;Data=@()},
        @{Phase='HostExclusivityObservation';Sequence=15;Data=@{OwnedAlive=$true;BeforeDispatch=$true;Competitors=@(@{ProcessId=4321})}},
        @{Phase='HostRetained';Sequence=16;Data=@{Request=$null;Response=$null;CommittedRestoreEntries=@();Error='System.InvalidOperationException: The owned Excel identity or exclusive VBE-host interval changed; retain without further native dispatch.'}}
    )
    foreach($record in $records){$record.ProcessId=1234;$record.ProcessStartUtc='2026-10-03T17:00:00Z';$record.ProductMvid='frozen'}
    return ,($records | ConvertTo-Json -Depth 30 | ConvertFrom-Json)
}
$passed=0
function ExpectRefusal($name,[scriptblock]$change) {
    $records=Fixture
    & $change $records
    $refused=$false
    try { $null=Resolve-Q026RecoveryBaseline -Records $records -Ledger $records[-1] -ProductMvid 'frozen' }
    catch {$refused=$true}
    if(-not $refused){throw ('Unsafe case accepted: '+$name)}
    $script:passed++
}
$records=Fixture
$result=Resolve-Q026RecoveryBaseline -Records $records -Ledger $records[-1] -ProductMvid 'frozen'
if($result.Kind -ne 'VerifiedWarmupBeforeScenario' -or $result.Record.Data.OptionsVersion -ne ('a'*64)){
    throw 'Complete unchanged warmup was rejected.'
}
$passed++
ExpectRefusal 'foreign PID' {param($r) $r[0].ProcessId=5678}
ExpectRefusal 'foreign start' {param($r) $r[1].ProcessStartUtc='2026-10-03T17:01:00Z'}
ExpectRefusal 'foreign binary' {param($r) $r[2].ProductMvid='another'}
ExpectRefusal 'missing full readback' {param($r) $r[1].Data.IndependentClosedReadback.Tabs=$null}
ExpectRefusal 'changed structure with equal hash' {param($r) $r[1].Data.IndependentClosedReadback.Tabs[0].Controls[0].Value='Courier New'}
ExpectRefusal 'pending warmup' {param($r) $r[1].Data.IndependentClosedReadback | Add-Member Pending $true}
ExpectRefusal 'uncertain refusal' {param($r) $r[1].Data.Response | Add-Member DeliveryUncertain $true}
ExpectRefusal 'ambiguous warmup receipt' {param($r) $r[3].Phase='GuardBreakpointWarmupBeforeReply'}
ExpectRefusal 'known preference write' {param($r) $r[2].Data.NativePreferenceWrites=1}
ExpectRefusal 'intervening native dispatch' {param($r) $r[3].Phase='FontBeforeWriteIntent'}
ExpectRefusal 'positive ledger' {param($r) $r[-1].Data.CommittedRestoreEntries=@(@{Property='Font'})}
ExpectRefusal 'unknown request' {param($r) $r[-1].Data.Request=@{Command='set_vbe_option'}}
ExpectRefusal 'unknown error' {param($r) $r[-1].Data.Error='Uncertain delivery'}
ExpectRefusal 'unverified dialog closure' {param($r) $r[0].Data.Response.Data.DialogClosed=$false}
ExpectRefusal 'isolation sequence gap' {param($r) $r[4].Sequence=14}
$scenario=[pscustomobject]@{Phase='BaselineComplete';Data=@{OptionsVersion='existing'}}
$result=Resolve-Q026RecoveryBaseline -Records @($scenario) -Ledger $records[-1] -ProductMvid 'frozen'
if($result.Kind -ne 'ScenarioBaseline' -or $result.Record -ne $scenario){throw 'Existing scenario baseline changed.'}
$passed++
[pscustomobject]@{State='PASSED';Passed=$passed;Failed=0;NativeDispatches=0;OfficeActivated=$false} | ConvertTo-Json

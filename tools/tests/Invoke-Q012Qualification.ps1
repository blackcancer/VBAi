#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$CandidateAssembly,
    [Parameter(Mandatory=$true)][string]$BuildOutputRoot,
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
    [Parameter(Mandatory=$true)][string]$DesktopHelperAssembly,
    [string]$TestProject,
    [ValidateSet('All','Access','Publisher')][string]$HostScope='All',
    [ValidateRange(1,30)][int[]]$ScenarioNumbers=(1..15),
    [switch]$StopOnNativeFailure,
    [ValidateRange(0,3600)][int]$NativeScenarioTimeoutSeconds=0,
    [ValidateRange(0,3600)][int]$ManagedGateTimeoutSeconds=0,
    [string]$BlockedHostReason='Host excluded by the reviewed plan; prior outcomes are not promoted to acceptance.',
    [ValidateSet('Debug','Release')][string]$Configuration='Debug',
    [switch]$MetadataGetterProbe,
    [string]$PublisherSerializedSeed,
    [switch]$Execute
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
if($PSVersionTable.PSEdition -ne 'Desktop' -or -not [Environment]::Is64BitProcess){throw 'Windows PowerShell 5.1 x64 required.'}
$repo=(Resolve-Path -LiteralPath (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent)).Path
if(-not $TestProject){$TestProject=Join-Path $repo 'tests/VBAi.Tests/VBAi.Tests.csproj'}
$registrationScript=Join-Path $repo 'tools/testing-explorer/Set-TestExplorerCandidate.ps1'
$testAssembly=Join-Path $BuildOutputRoot ('VBAi.Tests/'+$Configuration+'/net48/VBAi.Tests.dll')
$unitNames=@('VbeProjectGeneralOperationTests','VbeProjectGeneralNativeTests','VbeProjectGeneralProjectTests',
    'VbeDebugGeneralCommandTests','LlmVbeToolsProjectGeneralTests','BridgeServerTests','VbeSessionContractTests',
    'VbeOtherHostPersistenceTests','OfficeProjectReopenIdentityTests','OfficePublisherStartupBindingTests',
    'OfficeVbeFixturePublisherTestCleanupTests','OfficeVbeFixturePublisherBootstrapTests','OfficeVbeFixturePublisherSerializedSeedTests','OfficeMetadataMutationEvidenceTests','VbeScalarPropertyTests','AccessHelpContextDispatchTests','AccessHelpContextProjectTests','IsolatedTestDesktopTests','OfficeVbeFixtureDesktopTests','OfficeVbeFixtureDesktopAddInConnectionTests','OfficeVbeFixtureDesktopStartupRecoveryTests','OfficeVbeFixturePublisherOwnershipTests','AccessSaveConfirmationTests','LlmVbeToolsBoundaryTests','ChatWindowStateTests')
$unitFilter=($unitNames | ForEach-Object {
    $selector='FullyQualifiedName~VBAi.Tests.Unit.'+$_
    if($_ -in @('BridgeServerTests','VbeSessionContractTests')){return $selector+'.General'}
    if($_ -ceq 'ChatWindowStateTests'){$selector+'.CachedMetadata'}else{$selector}
}) -join '|'
$unitFilter+='|FullyQualifiedName~VBAi.Tests.Unit.BridgeServerTests.PendingGeneralBlocksBridgeNativeRoutesBeforeAnyEntry|FullyQualifiedName~VBAi.Tests.Unit.BridgeServerTests.UncertainGeneralBlocksBridgeNativeRoutesBeforeAnyEntry|FullyQualifiedName~VBAi.Tests.Unit.BridgeServerTests.NativeBridgeAdmissionExcludesGeneralUntilWorkerReturnsOrThrows'
$unitFilter+='|FullyQualifiedName~VBAi.Tests.Unit.VbeProjectLegacyHelpMetadataTests|FullyQualifiedName~VBAi.Tests.Unit.VbeProjectComponentsTests|FullyQualifiedName~VBAi.Tests.Unit.VbeProjectScalarFailureTests|FullyQualifiedName~VBAi.Tests.Unit.ProjectHelpTests|FullyQualifiedName~VBAi.Tests.Unit.ProjectHelpNativeBoundaryTests|FullyQualifiedName~VBAi.Tests.Unit.VbeSessionContractTests'
$scenarioRows=@(
    'Access|OfficeAdapterOnlyQualificationTests|Access16ActiveModuleOnlyAdapterSaveReopen',
    'Access|OfficeAdapterOnlyQualificationTests|Access16ModuleAndClassAdapterSaveReopen',
    'Access|OfficeAdapterOnlyReferenceQualificationTests|Access16ReferenceAdditionAdapterSaveReopen',
    'Access|OfficeAdapterOnlyReferenceQualificationTests|Access16ReferenceFileAdditionAdapterSaveReopen',
    'Access|OfficeAdapterOnlyReferenceQualificationTests|Access16ReferenceRemovalAdapterSaveReopen',
    'Access|OfficeAdapterOnlyMetadataQualificationTests|Access16DescriptionAdapterSaveReopen',
    'Publisher|OfficeAdapterOnlyQualificationTests|PublisherAdapterOnlySaveReopen',
    'Publisher|OfficeAdapterOnlyReferenceQualificationTests|PublisherReferenceAdditionAdapterSaveReopen',
    'Publisher|OfficeAdapterOnlyReferenceQualificationTests|PublisherReferenceFileAdditionAdapterSaveReopen',
    'Publisher|OfficeAdapterOnlyReferenceQualificationTests|PublisherReferenceRemovalAdapterSaveReopen',
    'Publisher|OfficeAdapterOnlyMetadataQualificationTests|PublisherDescriptionAdapterSaveReopen',
    'Access|OfficeAdapterOnlyMetadataQualificationTests|Access16HelpFilePathAdapterSaveReopen',
    'Publisher|OfficeAdapterOnlyMetadataQualificationTests|PublisherHelpFilePathAdapterSaveReopen',
    'Access|OfficeAdapterOnlyMetadataQualificationTests|Access16HelpContextIdAdapterSaveReopen',
    'Publisher|OfficeAdapterOnlyMetadataQualificationTests|PublisherHelpContextIdAdapterSaveReopen',
    'Publisher|PublisherGeneralQualificationTests|PublisherNativeHelpFileSaveReopen',
    'Publisher|PublisherGeneralQualificationTests|PublisherNativeHelpContextSaveReopen',
    'Publisher|PublisherGeneralQualificationTests|PublisherNativeAnsiHelpFileSaveReopen',
    'Publisher|PublisherGeneralQualificationTests|PublisherNativeUnicodeHelpFileRefusedBeforeWrite',
    'Access|AccessGeneralQualificationTests|AccessNativeAnsiHelpFileSaveReopen',
    'Access|AccessGeneralQualificationTests|AccessNativeHelpContextSaveReopen',
    'Publisher|OfficeAdapterOnlyQualificationTests|PublisherSerializedAdapterOnlySaveReopen',
    'Publisher|OfficeAdapterOnlyReferenceQualificationTests|PublisherSerializedReferenceAdditionAdapterSaveReopen',
    'Publisher|OfficeAdapterOnlyReferenceQualificationTests|PublisherSerializedReferenceFileAdditionAdapterSaveReopen',
    'Publisher|OfficeAdapterOnlyReferenceQualificationTests|PublisherSerializedReferenceRemovalAdapterSaveReopen',
    'Publisher|OfficeAdapterOnlyMetadataQualificationTests|PublisherSerializedDescriptionAdapterSaveReopen',
    'Access|LegacyHelpMetadataRefusalTests|AccessLegacyHelpFileRefusedBeforeWrite',
    'Access|LegacyHelpMetadataRefusalTests|AccessLegacyHelpContextRefusedBeforeWrite',
    'Publisher|LegacyHelpMetadataRefusalTests|PublisherLegacyHelpFileRefusedBeforeWrite',
    'Publisher|LegacyHelpMetadataRefusalTests|PublisherLegacyHelpContextRefusedBeforeWrite'
)
$scenarios=@()
foreach($row in $scenarioRows){
    $parts=$row.Split('|')
    $scenarios+=[pscustomobject][ordered]@{Number=$scenarios.Count+1;Host=$parts[0];Method=$parts[2];
        FullyQualifiedName='VBAi.Tests.Integration.'+$parts[1]+'.'+$parts[2]}
}
if($ScenarioNumbers.Count -eq 0 -or @($ScenarioNumbers | Select-Object -Unique).Count -ne $ScenarioNumbers.Count){
    throw 'Select at least one distinct planned scenario number; duplicates are refused.'
}
function Write-Report([string]$Path,$Value){
    [IO.File]::WriteAllText($Path,(ConvertTo-Json -InputObject $Value -Depth 100),[Text.UTF8Encoding]::new($false))
}
function Absolute-Existing([string]$Path,[bool]$Directory=$false){
    if(-not [IO.Path]::IsPathRooted($Path)){throw ('Absolute path required: '+$Path)}
    $kind=if($Directory){'Container'}else{'Leaf'}
    if(-not(Test-Path -LiteralPath $Path -PathType $kind)){throw ('Required path absent: '+$Path)}
    return (Resolve-Path -LiteralPath $Path).Path
}
function Hash-File([string]$Path){return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash}
function Git-Read([string[]]$Arguments){
    $gitPath=(Get-Command git.exe -ErrorAction Stop).Source
    $data=@(& $gitPath -c core.quotepath=false -C $repo @Arguments)
    if($LASTEXITCODE -ne 0){throw 'Read-only Git inventory failed.'}
    return $data
}
function Source-Snapshot{
    $paths=@(Git-Read @('ls-files','-c','-o','--exclude-standard') | Sort-Object -Unique)
    return @($paths | ForEach-Object {
        $absolute=Join-Path $repo $_
        if(-not(Test-Path -LiteralPath $absolute -PathType Leaf)){throw ('Source absent: '+$_)}
        [pscustomobject][ordered]@{Path=$_;Sha256=(Hash-File $absolute)}
    })
}
function Binary-Snapshot{
    $directories=@((Split-Path $CandidateAssembly -Parent),(Split-Path $testAssembly -Parent),(Split-Path $DesktopHelperAssembly -Parent)) | Sort-Object -Unique
    return @($directories | ForEach-Object {
        Get-ChildItem -LiteralPath $_ -File -Recurse | ForEach-Object {
            if(($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0){throw 'Binary reparse points refused.'}
            [pscustomobject][ordered]@{Path=$_.FullName;Sha256=(Hash-File $_.FullName)}
        }
    } | Sort-Object Path -Unique)
}
function Canonical-Json($Value){return ConvertTo-Json -InputObject $Value -Depth 100 -Compress}
function Require-Frozen($Plan){
    foreach($entry in @(@{Path=$Plan.Dotnet;Hash=$Plan.DotnetSha256},@{Path=$Plan.GitExecutable;Hash=$Plan.GitExecutableSha256},
        @{Path=$Plan.AccessExecutable;Hash=$Plan.AccessExecutableSha256},@{Path=$Plan.PublisherExecutable;Hash=$Plan.PublisherExecutableSha256})){
        if((Hash-File $entry.Path) -cne $entry.Hash){throw 'Command/native executable hash changed before another test.'}
    }
    if([string](Git-Read @('rev-parse','HEAD')) -cne $Plan.SourceRevision){throw 'Source revision changed.'}
    if((Canonical-Json @(Source-Snapshot)) -cne (Canonical-Json @($Plan.SourceFiles))){throw 'Source inventory/hash drift refused before another test.'}
    if((Canonical-Json @(Binary-Snapshot)) -cne (Canonical-Json @($Plan.BinaryFiles))){throw 'Frozen binary inventory/hash drift refused before another test.'}
}
function Host-Inventory([string]$HostKind){
    $name=if($HostKind -eq 'Access'){'MSACCESS'}else{'MSPUB'}
    foreach($process in @(Get-Process -Name $name -ErrorAction SilentlyContinue)){
        try{[pscustomobject]@{Host=$HostKind;Pid=$process.Id;StartedUtc=$process.StartTime.ToUniversalTime().ToString('o')}}
        finally{$process.Dispose()}
    }
}
function Installed-Image([string]$Name){
    $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::Registry64)
    $key=$null
    try{
        $key=$base.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\'+$Name,$false)
        if($null -eq $key){throw ('Installed x64 App Paths absent: '+$Name)}
        return Absolute-Existing ([string]$key.GetValue($null))
    }finally{if($null -ne $key){$key.Dispose()};$base.Dispose()}
}
function Run-Tests([string]$Filter,[string]$Directory,[string]$Name,[bool]$Single){
    if(Test-Path -LiteralPath $Directory){throw 'Each test requires a new evidence directory; replay refused.'}
    [IO.Directory]::CreateDirectory($Directory) | Out-Null
    $log=Join-Path $Directory 'console.log'
    $arguments=@('test',$TestProject,'-c',$Configuration,'--no-build','--no-restore',
        ('-p:BuildOutputRoot='+$BuildOutputRoot),'-p:BuildProjectReferences=false','--filter',$Filter,
        '--logger',('trx;LogFileName='+$Name+'.trx'),'--results-directory',$Directory)
    $hangTimeout=if($Single){$NativeScenarioTimeoutSeconds}else{$ManagedGateTimeoutSeconds}
    if($hangTimeout -gt 0){
        $arguments+=@('--blame-hang-timeout',($hangTimeout.ToString()+'s'),'--blame-hang-dump-type','none')
    }
    Write-Report (Join-Path $Directory 'invocation-intent.json') @{Filter=$Filter;InvocationCount=1;StartingUtc=[DateTime]::UtcNow.ToString('o');Arguments=$arguments;Desktop=$env:VBAi_TEST_DESKTOP_NAME}
    & $plan.Dotnet @arguments *> $log
    $code=$LASTEXITCODE
    Write-Report (Join-Path $Directory 'invocation-terminal.json') @{ExitCode=$code;ReturnedUtc=[DateTime]::UtcNow.ToString('o');InvocationCount=1}
    $trx=Join-Path $Directory ($Name+'.trx')
    $outcomes=@();$errorText=$null
    try{
        if(-not(Test-Path -LiteralPath $trx -PathType Leaf)){throw 'Terminal TRX absent; no result inferred.'}
        [xml]$xml=Get-Content -LiteralPath $trx -Raw -Encoding UTF8
        $definitions=@{}
        foreach($definition in @($xml.SelectNodes("//*[local-name()='TestDefinitions']/*[local-name()='UnitTest']"))){
            $method=$definition.SelectSingleNode("*[local-name()='TestMethod']")
            if($null -ne $method){
                $definitions[$definition.GetAttribute('id')]=[pscustomobject]@{Class=$method.GetAttribute('className').Split(',')[0].Trim();Method=$method.GetAttribute('name')}
            }
        }
        $outcomes=@($xml.SelectNodes("//*[local-name()='UnitTestResult']") | ForEach-Object {
            $testId=$_.GetAttribute('testId')
            if(-not $definitions.ContainsKey($testId)){throw 'TRX result has no exact test definition.'}
            $definition=$definitions[$testId]
            [pscustomobject]@{Name=$_.GetAttribute('testName');Outcome=$_.GetAttribute('outcome');Class=$definition.Class;Method=$definition.Method}
        })
        if($outcomes.Count -lt 1 -or ($Single -and $outcomes.Count -ne 1)){throw 'Unexpected scenario count; original TRX retained.'}
        if($Single){
            $expected=$Filter.Substring('FullyQualifiedName='.Length)
            if(($outcomes[0].Class+'.'+$outcomes[0].Method) -cne $expected){throw 'TRX method differs from the planned single native scenario.'}
        }else{
            foreach($class in $unitNames){
                if(-not @($outcomes | Where-Object {$_.Class -ceq ('VBAi.Tests.Unit.'+$class)}).Count){
                    throw ('Required regression class has no TRX outcomes: '+$class)
                }
            }
        }
    }catch{$errorText=$_.Exception.ToString()}
    $passed=($code -eq 0 -and $null -eq $errorText -and @($outcomes | Where-Object Outcome -ne 'Passed').Count -eq 0)
    return [pscustomobject]@{State=if($passed){'PASS'}else{'FAIL'};ExitCode=$code;Trx=$trx;Console=$log;
        Outcomes=$outcomes;EvidenceError=$errorText;InvocationCount=1}
}
$CandidateAssembly=Absolute-Existing $CandidateAssembly
$BuildOutputRoot=Absolute-Existing $BuildOutputRoot $true
$TestProject=Absolute-Existing $TestProject
$DesktopHelperAssembly=Absolute-Existing $DesktopHelperAssembly
$testAssembly=Absolute-Existing $testAssembly
$registrationScript=Absolute-Existing $registrationScript
if($PublisherSerializedSeed){$PublisherSerializedSeed=Absolute-Existing $PublisherSerializedSeed}
$publisherSeedHash=if($PublisherSerializedSeed){Hash-File $PublisherSerializedSeed}else{''}
$publisherSeedProvenance=@()
if($PublisherSerializedSeed){
    foreach($receiptName in @('shutdown-before-reopen-183824.json','adapter-only-progress.json')){
        $receiptPath=Absolute-Existing (Join-Path (Split-Path $PublisherSerializedSeed -Parent) $receiptName)
        $publisherSeedProvenance+=@{Path=$receiptPath;Sha256=(Hash-File $receiptPath)}
    }
}
if(-not [IO.Path]::IsPathRooted($EvidenceDirectory)){throw 'Absolute evidence directory required.'}
$EvidenceDirectory=[IO.Path]::GetFullPath($EvidenceDirectory)
foreach($output in @($BuildOutputRoot,(Split-Path $CandidateAssembly -Parent),(Split-Path $DesktopHelperAssembly -Parent))){
    if($EvidenceDirectory.StartsWith($output.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase) -or $EvidenceDirectory -ieq $output){
        throw 'Evidence must be separate from frozen binary outputs.'
    }
}
$testProduct=Absolute-Existing (Join-Path (Split-Path $testAssembly -Parent) 'VBAi.dll')
if((Hash-File $testProduct) -cne (Hash-File $CandidateAssembly)){throw 'Test output references another product candidate.'}
$identity=[Reflection.Assembly]::ReflectionOnlyLoadFrom($CandidateAssembly)
if($identity.GetName().Name -cne 'VBAi'){throw 'Expected VBAi candidate.'}
$mvid=$identity.ManifestModule.ModuleVersionId.ToString('D')
$planPath=Join-Path $EvidenceDirectory 'plan.json'
if(-not $Execute){
    if(Test-Path -LiteralPath $EvidenceDirectory){throw 'Prepare requires a fresh evidence directory.'}
    [IO.Directory]::CreateDirectory($EvidenceDirectory) | Out-Null
    $preview=& $registrationScript -CandidateAssemblyPath $CandidateAssembly -ExpectedMvid ([Guid]$mvid)
    Write-Report (Join-Path $EvidenceDirectory 'registration-preview.json') $preview
    $dotnet=(Get-Command dotnet.exe -ErrorAction Stop).Source
    $gitExecutable=(Get-Command git.exe -ErrorAction Stop).Source
    $accessExecutable=Installed-Image 'MSACCESS.EXE'
    $publisherExecutable=Installed-Image 'MSPUB.EXE'
    $plan=[ordered]@{Format='VBAi.Q012.Campaign.1';PreparedUtc=[DateTime]::UtcNow.ToString('o');Repository=$repo;
        SourceRevision=[string](Git-Read @('rev-parse','HEAD'));SourceBranch=[string](Git-Read @('branch','--show-current'));
        CandidateAssembly=$CandidateAssembly;CandidateMvid=$mvid;CandidateSha256=(Hash-File $CandidateAssembly);
        BuildOutputRoot=$BuildOutputRoot;TestProject=$TestProject;TestAssembly=$testAssembly;Configuration=$Configuration;
        DesktopHelperAssembly=$DesktopHelperAssembly;EvidenceDirectory=$EvidenceDirectory;Dotnet=$dotnet;DotnetSha256=(Hash-File $dotnet);
        AccessExecutable=$accessExecutable;AccessExecutableSha256=(Hash-File $accessExecutable);
        PublisherExecutable=$publisherExecutable;PublisherExecutableSha256=(Hash-File $publisherExecutable);
        GitExecutable=$gitExecutable;GitExecutableSha256=(Hash-File $gitExecutable);
        SourceFiles=@(Source-Snapshot);BinaryFiles=@(Binary-Snapshot);UnitFilter=$unitFilter;NativeScenarios=$scenarios;
        ExistingAccess=@(Host-Inventory 'Access');ExistingPublisher=@(Host-Inventory 'Publisher');
        HostScope=$HostScope;BlockedHostReason=$BlockedHostReason;SelectedScenarioNumbers=@($ScenarioNumbers);
        StopOnNativeFailure=[bool]$StopOnNativeFailure;NativeScenarioTimeoutSeconds=$NativeScenarioTimeoutSeconds;
        ManagedGateTimeoutSeconds=$ManagedGateTimeoutSeconds;
        PublisherSerializedSeed=$PublisherSerializedSeed;PublisherSerializedSeedSha256=$publisherSeedHash;
        PublisherSerializedSeedProvenance=$publisherSeedProvenance;
        NativeInvocationLimit=1;NativeSaveReplay=$false;ForceTermination=$false;InputDesktopFallback=$false;MetadataGetterProbe=[bool]$MetadataGetterProbe;
        Scope='Existing ACCDB/PUB save; no first SaveAs, macro execution, trust changes or signatures.'}
    Write-Report $planPath $plan
    Write-Output ('PREPARED: '+$planPath)
    return
}
if(-not(Test-Path -LiteralPath $planPath -PathType Leaf)){throw 'Prepare the reviewed frozen plan first.'}
$plan=Get-Content -LiteralPath $planPath -Raw -Encoding UTF8 | ConvertFrom-Json
$plannedStopOnFailure=$false
$plannedNativeTimeout=0
$plannedManagedTimeout=0
$plannedPublisherSeed='';$plannedPublisherSeedHash=''
if($null -ne $plan.PSObject.Properties['PublisherSerializedSeed']){$plannedPublisherSeed=[string]$plan.PublisherSerializedSeed;$plannedPublisherSeedHash=[string]$plan.PublisherSerializedSeedSha256}
$plannedPublisherSeedProvenance=@()
if($null -ne $plan.PSObject.Properties['PublisherSerializedSeedProvenance']){$plannedPublisherSeedProvenance=@($plan.PublisherSerializedSeedProvenance)}
if($null -ne $plan.PSObject.Properties['StopOnNativeFailure']){$plannedStopOnFailure=[bool]$plan.StopOnNativeFailure}
if($null -ne $plan.PSObject.Properties['NativeScenarioTimeoutSeconds']){$plannedNativeTimeout=[int]$plan.NativeScenarioTimeoutSeconds}
if($null -ne $plan.PSObject.Properties['ManagedGateTimeoutSeconds']){$plannedManagedTimeout=[int]$plan.ManagedGateTimeoutSeconds}
if($plan.Format -cne 'VBAi.Q012.Campaign.1' -or $plan.Repository -cne $repo -or
    $plan.CandidateAssembly -cne $CandidateAssembly -or $plan.CandidateMvid -cne $mvid -or
    $plan.BuildOutputRoot -cne $BuildOutputRoot -or $plan.TestProject -cne $TestProject -or
    $plan.TestAssembly -cne $testAssembly -or $plan.Configuration -cne $Configuration -or
    $plan.DesktopHelperAssembly -cne $DesktopHelperAssembly -or $plan.EvidenceDirectory -cne $EvidenceDirectory -or
    $plan.HostScope -cne $HostScope -or $plan.BlockedHostReason -cne $BlockedHostReason -or
    $plan.MetadataGetterProbe -ne [bool]$MetadataGetterProbe -or
    $plannedPublisherSeed -cne [string]$PublisherSerializedSeed -or $plannedPublisherSeedHash -cne $publisherSeedHash -or
    (Canonical-Json $plannedPublisherSeedProvenance) -cne (Canonical-Json $publisherSeedProvenance) -or
    $plannedStopOnFailure -ne [bool]$StopOnNativeFailure -or $plannedNativeTimeout -ne $NativeScenarioTimeoutSeconds -or
    $plannedManagedTimeout -ne $ManagedGateTimeoutSeconds -or
    (Canonical-Json @($plan.SelectedScenarioNumbers)) -cne (Canonical-Json @($ScenarioNumbers)) -or
    $plan.UnitFilter -cne $unitFilter -or (Canonical-Json @($plan.NativeScenarios)) -cne (Canonical-Json $scenarios)){
    throw 'Plan identity or fixed scenario inventory changed.'
}
if((Hash-File $plan.Dotnet) -cne $plan.DotnetSha256 -or
    $plan.AccessExecutable -cne (Installed-Image 'MSACCESS.EXE') -or
    $plan.PublisherExecutable -cne (Installed-Image 'MSPUB.EXE') -or
    (Hash-File $plan.AccessExecutable) -cne $plan.AccessExecutableSha256 -or
    (Hash-File $plan.PublisherExecutable) -cne $plan.PublisherExecutableSha256 -or
    $plan.GitExecutable -cne (Get-Command git.exe -ErrorAction Stop).Source -or
    (Hash-File $plan.GitExecutable) -cne $plan.GitExecutableSha256){throw 'Frozen executable selection/hash changed.'}
Require-Frozen $plan
$desktop=$env:VBAi_TEST_DESKTOP_NAME
$desktopType=[Reflection.Assembly]::LoadFrom($DesktopHelperAssembly).GetType('VBAi.Tests.Integration.IsolatedTestDesktop',$true)
$flags=[Reflection.BindingFlags]'Static,NonPublic'
$desktopType.GetMethod('RequireCurrent',$flags).Invoke($null,@($desktop)) | Out-Null
$claimPath=Join-Path $EvidenceDirectory 'execution-claim.json'
$claim=[IO.File]::Open($claimPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
$claim.Dispose()
Write-Report $claimPath @{State='CLAIMED_ONCE';PlanSha256=(Hash-File $planPath);Pid=$PID;Desktop=$desktop;Utc=[DateTime]::UtcNow.ToString('o')}
$summary=[ordered]@{State='RUNNING';FullMatrixQualified=$false;CandidateMvid=$mvid;CandidateSha256=$plan.CandidateSha256;Desktop=$desktop;
    Unit=$null;Scenarios=@();RegistrationApply=$null;RegistrationRestore=$null;Failure=$null;CompletedUtc=$null}
$summaryPath=Join-Path $EvidenceDirectory 'summary.json'
$pattern='^VBAi_RUN_|^VBAi_TEST_|^VBAi_OFFICE_RESULTS$|^VBAi_SOLIDWORKS_PID$|^VBAI_EDITOR_|^VBAI_NATIVE_'
$desktopVariables=@('VBAi_TEST_DESKTOP_NAME','VBAi_TEST_DESKTOP_SENTINEL_HWND','VBAi_TEST_DESKTOP_SENTINEL_PID','VBAi_TEST_DESKTOP_SENTINEL_TID')
$variables=@(Get-ChildItem Env: | Where-Object {$_.Name -match $pattern})
$backup=Join-Path $EvidenceDirectory 'candidate-registration.clixml'
$applyAttempted=$false
try{
    foreach($variable in $variables){if($variable.Name -notin $desktopVariables){Remove-Item -LiteralPath ('Env:\'+$variable.Name)}}
    $summary.Unit=Run-Tests $unitFilter (Join-Path $EvidenceDirectory 'unit') 'unit' $false
    Write-Report $summaryPath $summary
    Require-Frozen $plan
    if($summary.Unit.State -ne 'PASS'){
        $summary.State='UNIT_FAILED_NATIVE_NOT_RUN'
        foreach($scenario in $scenarios){$summary.Scenarios+=@{Scenario=$scenario;State='NOT_RUN';Reason='Required unit gate failed';InvocationCount=0}}
    }else{
        $applyAttempted=$true
        $summary.RegistrationApply=& $registrationScript -CandidateAssemblyPath $CandidateAssembly -ExpectedMvid ([Guid]$mvid) -Apply -ReportPath $backup
        Write-Report $summaryPath $summary
        $env:VBAi_RUN_OFFICE_TESTS='1';$env:VBAi_TEST_ACCESS_EXE=$plan.AccessExecutable;$env:VBAi_TEST_PUBLISHER_EXE=$plan.PublisherExecutable
        if($MetadataGetterProbe){$env:VBAi_RUN_OFFICE_METADATA_GETTER_PROBE='1'}
        if($PublisherSerializedSeed){$env:VBAi_TEST_PUBLISHER_SERIALIZED_SEED=$PublisherSerializedSeed}
        $nativeFailureScenario=$null
        foreach($scenario in $scenarios){
            if($scenario.Number -notin $ScenarioNumbers){
                $summary.Scenarios+=@{Scenario=$scenario;State='NOT_RUN';Reason='Outside the frozen diagnostic selection';InvocationCount=0}
            }elseif($null -ne $nativeFailureScenario){
                $summary.Scenarios+=@{Scenario=$scenario;State='BLOCKED';Reason=('Campaign stopped after native scenario '+$nativeFailureScenario+' failed; no repeat before its cause is resolved');InvocationCount=0}
            }elseif($HostScope -ne 'All' -and $scenario.Host -ne $HostScope){
                $summary.Scenarios+=@{Scenario=$scenario;State='BLOCKED';Reason=$BlockedHostReason;InvocationCount=0}
            }else{
                # Full source/binary checks guard actual invocations and the final result.
                # Unselected rows cannot dispatch and need no repeated repository hashing.
                Require-Frozen $plan
                if($PublisherSerializedSeed -and (Hash-File $PublisherSerializedSeed) -cne $publisherSeedHash){throw 'Frozen Publisher seed changed before dispatch.'}
                foreach($receipt in $publisherSeedProvenance){if((Hash-File $receipt.Path) -cne $receipt.Sha256){throw 'Frozen Publisher seed provenance changed before dispatch.'}}
                $desktopType.GetMethod('RequireCurrent',$flags).Invoke($null,@($desktop)) | Out-Null
                $existing=@(Host-Inventory $scenario.Host)
                if($existing.Count){
                    $summary.Scenarios+=@{Scenario=$scenario;State='BLOCKED';Reason='Existing/retained same-host process; ownership would refuse';Processes=$existing;InvocationCount=0}
                    Write-Report $summaryPath $summary
                    continue
                }
                $directory=Join-Path $EvidenceDirectory ('native/'+$scenario.Number.ToString('00')+'-'+$scenario.Method)
                $env:VBAi_OFFICE_RESULTS=Join-Path $directory 'host-evidence'
                $result=Run-Tests ('FullyQualifiedName='+$scenario.FullyQualifiedName) $directory 'native' $true
                $summary.Scenarios+=@{Scenario=$scenario;State=$result.State;Result=$result;RemainingProcesses=@(Host-Inventory $scenario.Host);InvocationCount=1}
                if($StopOnNativeFailure -and $result.State -ne 'PASS'){$nativeFailureScenario=$scenario.Number}
            }
            Write-Report $summaryPath $summary
        }
        Require-Frozen $plan
        $selectedResults=@($summary.Scenarios | Where-Object {$_.Scenario.Number -in $ScenarioNumbers})
        $summary.State=if($selectedResults.Count -ne $ScenarioNumbers.Count -or @($selectedResults | Where-Object {$_.State -ne 'PASS'}).Count){
            'FAILED_OR_BLOCKED'
        }elseif($ScenarioNumbers.Count -lt $scenarios.Count){'DIAGNOSTIC_PASS'}else{'PASS'}
    }
}catch{
    $summary.State='FAILED';$summary.Failure=$_.Exception.ToString()
    foreach($scenario in $scenarios){
        if(-not @($summary.Scenarios | Where-Object {$_.Scenario.FullyQualifiedName -ceq $scenario.FullyQualifiedName}).Count){
            $summary.Scenarios+=@{Scenario=$scenario;State='NOT_RUN';Reason='Campaign terminal refusal before invocation';InvocationCount=0}
        }
    }
}finally{
    if($applyAttempted){
        try{
            if(Test-Path -LiteralPath ($backup+'.after.clixml') -PathType Leaf){$summary.RegistrationRestore=& $registrationScript -Restore -BackupPath $backup}
            else{throw 'No completed applied-state backup; registry inspection required, automatic restore refused.'}
        }catch{$summary.State='FAILED_OR_BLOCKED';$summary.RegistrationRestore=@{Verified=$false;Error=$_.Exception.ToString();Backup=$backup;AutomaticRetry=$false}}
    }
    foreach($variable in @(Get-ChildItem Env: | Where-Object {$_.Name -match $pattern})){
        if($variable.Name -notin $desktopVariables){Remove-Item -LiteralPath ('Env:\'+$variable.Name)}
    }
    foreach($variable in $variables){Set-Item -LiteralPath ('Env:\'+$variable.Name) -Value $variable.Value}
    $summary.FullMatrixQualified=($summary.State -eq 'PASS' -and
        $applyAttempted -and $summary.RegistrationRestore.Restored -eq $true -and
        $summary.RegistrationRestore.Verified -eq $true)
    $summary.CompletedUtc=[DateTime]::UtcNow.ToString('o')
    Write-Report $summaryPath $summary
}
Write-Output ('Q012 '+$summary.State+': '+$summaryPath)
if($summary.State -notin @('PASS','DIAGNOSTIC_PASS')){exit 1}
exit 0

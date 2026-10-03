#requires -Version 5.1
param(
    [switch]$Prepare,
    [string]$EvidenceRoot,
    [string]$InstalledDirectory,
    [string]$PlanPath = (Join-Path $PSScriptRoot 'q026-plan.json'),
    [ValidateSet('FullFormat','Margin')][string]$Scenario = 'FullFormat'
)
$ErrorActionPreference = 'Stop'
function Write-Json($path, $value) {
    $value | ConvertTo-Json -Depth 18 | Set-Content -LiteralPath $path -Encoding UTF8
}
if ($Prepare) {
    if (-not [IO.Path]::IsPathRooted($EvidenceRoot) -or (Test-Path -LiteralPath $EvidenceRoot)) { throw 'A fresh absolute evidence root is required.' }
    $repository = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    $installed = Join-Path $InstalledDirectory 'VBAi.dll'
    if (-not [IO.Path]::IsPathRooted($InstalledDirectory) -or -not (Test-Path -LiteralPath $installed)) { throw 'A frozen installed candidate is required.' }
    [IO.Directory]::CreateDirectory($EvidenceRoot) | Out-Null
    $build = Join-Path $EvidenceRoot 'build'
    $product = Join-Path $build 'VBAi/Debug/net48'
    [IO.Directory]::CreateDirectory($product) | Out-Null
    Get-ChildItem -LiteralPath $InstalledDirectory -Force | Copy-Item -Destination $product -Recurse
    foreach ($project in @('tests/VBAi.Q026.Tests/VBAi.Q026.Tests.csproj','tests/VBAi.Desktop.Helper/VBAi.Desktop.Helper.csproj')) {
        $log = Join-Path $EvidenceRoot (([IO.Path]::GetFileNameWithoutExtension($project)) + '-build.log')
        & dotnet build (Join-Path $repository $project) -c Debug "-p:BuildOutputRoot=$build" "-p:FrozenProductDirectory=$product" --verbosity minimal *> $log
        if ($LASTEXITCODE -ne 0) { throw "Preparation failed; no host launched. See $log." }
    }
    $test = Join-Path $build 'VBAi.Q026.Tests/Debug/net48/VBAi.Tests.dll'
    $helper = Join-Path $build 'VBAi.Desktop.Helper/Debug/net48/VBAi.Desktop.Helper.exe'
    $hash = (Get-FileHash -LiteralPath $installed).Hash
    if ((Get-FileHash -LiteralPath (Join-Path ([IO.Path]::GetDirectoryName($test)) 'VBAi.dll')).Hash -cne $hash) { throw 'Test and installed candidates differ.' }
    $script = Join-Path $EvidenceRoot 'Invoke-FrozenQ026.ps1'
    Copy-Item -LiteralPath $PSCommandPath -Destination $script
    $files = @(Get-ChildItem -LiteralPath ([IO.Path]::GetDirectoryName($test)) -File | Where-Object {$_.Extension -in @('.dll','.exe','.config')} | ForEach-Object { @{Path=$_.FullName;Sha256=(Get-FileHash -LiteralPath $_.FullName).Hash} })
    $files += @{Path=$helper;Sha256=(Get-FileHash -LiteralPath $helper).Hash}
    $files += @{Path=$script;Sha256=(Get-FileHash -LiteralPath $script).Hash}
    $nativeMethod=if($Scenario -eq 'Margin'){'NativeMarginCheckboxRoundTripAndRestoreCompleteOptionsVersion'}else{'NativeFormatChoicesRoundTripAndRestoreCompleteOptionsVersion'}
    $matrix=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'q026-scenarios.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if($Scenario -eq 'Margin'){
        $matrix.Cases=@($matrix.Cases | Where-Object {$_.Id -in @('read-stability','margin','complete-restoration','normal-exit')})
        $matrix.Scope='Owned disposable Excel margin-checkbox diagnostic only; not the full Format matrix'
    }
    $plan = @{Scope=('Q-026 owned Excel '+$Scenario+' mutation/restoration; not all-host qualification');Scenario=$Scenario;Repository=$repository;
        SourceCommit=(& git -C $repository rev-parse HEAD);SourceStatus=@(& git -C $repository status --porcelain);
        InstalledProduct=$installed;ProductSha256=$hash;ProductMvid=([Reflection.Assembly]::ReflectionOnlyLoadFrom($installed)).ManifestModule.ModuleVersionId.ToString('D');
        EvidenceRoot=$EvidenceRoot;TestAssembly=$test;HelperAssembly=$helper;FrozenFiles=$files;
        Matrix=$matrix;
        NativeMethod=('VBAi.Tests.Integration.Hosts.Excel.ExcelFormatOptionsTests.'+$nativeMethod);
        QuietHostPeriodSeconds=30;NoNativeReplay=$true;NoForceTermination=$true;PreparedUtc=[DateTime]::UtcNow.ToString('o')}
    Write-Json (Join-Path $EvidenceRoot 'q026-plan.json') $plan
    Write-Output ('Prepared '+$EvidenceRoot)
    exit 0
}
$plan=Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
if(-not $env:VBAi_TEST_DESKTOP_NAME){throw 'A verified private-desktop worker is required; no foreground fallback.'}
foreach($file in $plan.FrozenFiles){if((Get-FileHash -LiteralPath $file.Path).Hash -cne $file.Sha256){throw 'Frozen qualification file changed: '+$file.Path}}
if((Get-FileHash -LiteralPath $plan.InstalledProduct).Hash -cne $plan.ProductSha256){throw 'Installed candidate changed after preparation.'}
$ledger=Join-Path $plan.EvidenceRoot 'campaign.json'
if(Test-Path -LiteralPath $ledger){throw 'One-shot campaign already claimed; do not replay.'}
$record=@{SourceCommit=$plan.SourceCommit;ProductMvid=$plan.ProductMvid;NativeState='NOT_RUN';ManagedState='STARTED_ONCE';NoNativeReplay=$true}
Write-Json $ledger $record
$env:VBAi_RUN_EXCEL_TESTS='0'
$managed=Join-Path $plan.EvidenceRoot 'managed'
& dotnet vstest $plan.TestAssembly '/TestCaseFilter:TestCategory=Unit' '/Logger:trx;LogFileName=managed.trx' "/ResultsDirectory:$managed" *> (Join-Path $plan.EvidenceRoot 'managed.log')
$record.ManagedExitCode=$LASTEXITCODE
$managedTrx=Join-Path $managed 'managed.trx'
if(-not(Test-Path -LiteralPath $managedTrx)){throw 'No managed terminal report; native work is forbidden.'}
[xml]$managedReport=Get-Content -LiteralPath $managedTrx
$record.ManagedCounters=$managedReport.TestRun.ResultSummary.Counters.OuterXml
$managedOk=$record.ManagedExitCode -eq 0 -and [int]$managedReport.TestRun.ResultSummary.Counters.executed -gt 0 -and [int]$managedReport.TestRun.ResultSummary.Counters.total -eq [int]$managedReport.TestRun.ResultSummary.Counters.passed
if(-not $managedOk){$record.ManagedState='FAILED';Write-Json $ledger $record;exit 1}
$record.ManagedState='PASS'
$quiet=[Diagnostics.Stopwatch]::StartNew()
do {
    $other=@(Get-Process EXCEL,WINWORD,POWERPNT,MSACCESS,MSPUB,SLDWORKS -ErrorAction SilentlyContinue)
    if($other.Count){$quiet.Restart();$record.WaitingHostPids=@($other.Id);$record.NativeState='WAITING_OTHER_HOSTS';Write-Json $ledger $record}
    else {$record.WaitingHostPids=@();$record.NativeState='OBSERVING_QUIET_HOST_PERIOD';Write-Json $ledger $record}
    Start-Sleep -Milliseconds 1000
} while($quiet.Elapsed.TotalSeconds -lt $plan.QuietHostPeriodSeconds)
# Read registry metadata only; refuse another candidate before any host launch.
$record.NativeState='VERIFYING_REGISTRATION';Write-Json $ledger $record
$registry=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryView]::Registry64)
$key=$registry.OpenSubKey('Software\Classes\CLSID\{8E854243-087F-4D6C-9E0E-8622B0E50883}\InprocServer32')
try {
    if(-not $key){throw 'The reviewed per-user x64 registration is absent.'}
    $codeBase=[string]$key.GetValue('CodeBase')
    $record.RegisteredCodeBase=$codeBase
    $record.ExpectedInstalledProduct=$plan.InstalledProduct
    $record.RegistrationIdentity=[Security.Principal.WindowsIdentity]::GetCurrent().Name
    Write-Json $ledger $record
    if(-not $codeBase -or ([Uri]$codeBase).LocalPath -ine $plan.InstalledProduct){throw 'Registered CodeBase differs from the frozen installed candidate; no host launched.'}
} catch {
    $record.NativeState='PRELAUNCH_REFUSED';$record.PrelaunchError=$_.ToString()
    Write-Json $ledger $record
    throw
} finally {if($key){$key.Dispose()};$registry.Dispose()}
if(@(Get-Process EXCEL,WINWORD,POWERPNT,MSACCESS,MSPUB,SLDWORKS -ErrorAction SilentlyContinue).Count){throw 'A competing host appeared after the quiet period; no launch.'}
$env:VBAi_RUN_EXCEL_TESTS='1'
$env:VBAi_EXCEL_RESULTS=Join-Path $plan.EvidenceRoot 'native/hosts'
$env:VBAi_TEST_FORMAT_OPTIONS_OUTPUT=Join-Path $plan.EvidenceRoot 'native/phases'
$record.NativeState='STARTED_ONCE';$record.NativeStartedUtc=[DateTime]::UtcNow.ToString('o');Write-Json $ledger $record
$native=Join-Path $plan.EvidenceRoot 'native'
& dotnet vstest $plan.TestAssembly ("/TestCaseFilter:FullyQualifiedName="+$plan.NativeMethod) '/Logger:trx;LogFileName=format.trx' "/ResultsDirectory:$native" *> (Join-Path $plan.EvidenceRoot 'native.log')
$record.NativeExitCode=$LASTEXITCODE
$nativeTrx=Join-Path $native 'format.trx'
if(Test-Path -LiteralPath $nativeTrx){
    [xml]$nativeReport=Get-Content -LiteralPath $nativeTrx
    $record.NativeCounters=$nativeReport.TestRun.ResultSummary.Counters.OuterXml
    $record.NativeState=if($record.NativeExitCode -eq 0 -and [int]$nativeReport.TestRun.ResultSummary.Counters.executed -eq 1 -and [int]$nativeReport.TestRun.ResultSummary.Counters.passed -eq 1){'PASS'}else{'FAILED_OR_SKIPPED'}
}else{$record.NativeState='NO_TERMINAL_REPORT'}
$record.RemainingExcelPids=@(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$record.CompletedUtc=[DateTime]::UtcNow.ToString('o');Write-Json $ledger $record
if($record.NativeState -ne 'PASS' -or $record.RemainingExcelPids.Count){exit 1}
exit 0

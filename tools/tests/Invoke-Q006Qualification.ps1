#requires -Version 5.1
param(
    [switch]$Prepare,
    [string]$EvidenceRoot,
    [string]$InstalledDirectory,
    [switch]$RecheckDesktopGuardOnly,
    [string]$ScenarioIds,
    [string]$PlanPath = (Join-Path $PSScriptRoot 'q006-plan.json')
)
$ErrorActionPreference = 'Stop'
Import-Module Microsoft.PowerShell.Utility
function Write-Json($path, $value) {
    $value | ConvertTo-Json -Depth 16 | Set-Content -LiteralPath $path -Encoding UTF8
}
if ($Prepare) {
    if (-not [IO.Path]::IsPathRooted($EvidenceRoot) -or (Test-Path -LiteralPath $EvidenceRoot)) { throw 'A fresh absolute evidence root is required.' }
    $repository = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    if (-not [IO.Path]::IsPathRooted($InstalledDirectory) -or -not (Test-Path -LiteralPath (Join-Path $InstalledDirectory 'VBAi.dll'))) { throw 'An existing frozen installed candidate is required.' }
    [IO.Directory]::CreateDirectory($EvidenceRoot) | Out-Null
    $build = Join-Path $EvidenceRoot 'build'
    $product = Join-Path $build 'VBAi/Debug/net48'
    [IO.Directory]::CreateDirectory($product) | Out-Null
    Copy-Item -LiteralPath (Join-Path $InstalledDirectory 'VBAi.dll') -Destination $product
    # Preserve installed dependencies/assets without rebuilding or changing registration.
    Get-ChildItem -LiteralPath $InstalledDirectory -Force | Where-Object { $_.Name -ne 'VBAi.dll' } | Copy-Item -Destination $product -Recurse
    foreach ($project in @('tests/VBAi.Q006.Tests/VBAi.Q006.Tests.csproj','tests/VBAi.Desktop.Helper/VBAi.Desktop.Helper.csproj')) {
        $log = Join-Path $EvidenceRoot (([IO.Path]::GetFileNameWithoutExtension($project)) + '-build.log')
        & dotnet build (Join-Path $repository $project) -c Debug "-p:BuildOutputRoot=$build" "-p:FrozenProductDirectory=$product" -p:BuildProjectReferences=false --verbosity minimal *> $log
        if ($LASTEXITCODE -ne 0) { throw "Qualification preparation failed; inspect $log. No native test was launched." }
    }
    $testAssembly = Join-Path $build 'VBAi.Q006.Tests/Debug/net48/VBAi.Tests.dll'
    $testProduct = Join-Path ([IO.Path]::GetDirectoryName($testAssembly)) 'VBAi.dll'
    $hash = (Get-FileHash -LiteralPath (Join-Path $InstalledDirectory 'VBAi.dll') -Algorithm SHA256).Hash
    if ((Get-FileHash -LiteralPath $testProduct).Hash -cne $hash) { throw 'Test output differs from the installed product.' }
    $script = Join-Path $EvidenceRoot 'Invoke-FrozenQ006.ps1'
    Copy-Item -LiteralPath $PSCommandPath -Destination $script
    $scenarios = @(
        @{Id='managed-guards'; Filter='TestCategory=Unit'; Native=$false; Oracle='Private desktop arguments, bounded scalar candidates, identity/revision/mode guards, uncertain shutdown refusal'}
        @{Id='scalar-unsupported'; Method='ExcelLocalScalarInspectionTests.InstalledBridgeSkipsUnsupportedScalarPageWithoutQuickWatch'; Native=$true; Oracle='No QuickWatch; array/Variant/object refusal; terminal phases; normal exit'}
        @{Id='scalar-long'; Method='ExcelLocalScalarInspectionTests.InstalledBridgeReadsOneLongScalarWithNativePhaseEvidence'; Native=$true; Oracle='Long=42, one observer, unchanged source/selection/mode; normal exit'}
        @{Id='scalar-original-page'; Method='ExcelLocalScalarInspectionTests.InstalledBridgeReadsFullScalarPageWithThreeNativeObservers'; Native=$true; Oracle='Three values and three refusals, exactly one request and three observers; normal exit'}
        @{Id='scalar-types-first'; Method='ExcelLocalScalarInspectionTests.InstalledBridgeReadsSupportedScalarFirstPageWithNativePhaseEvidence'; Native=$true; Oracle='Offset 0/limit 4: Long/String/Boolean/Byte values, four observers, terminal phases and normal exit'}
        @{Id='scalar-types-second'; Method='ExcelLocalScalarInspectionTests.InstalledBridgeReadsSupportedScalarSecondPageWithNativePhaseEvidence'; Native=$true; Oracle='Offset 4/limit 4: Integer/LongLong/LongPtr/Single values, four observers, terminal phases and normal exit'}
        @{Id='scalar-types-last'; Method='ExcelLocalScalarInspectionTests.InstalledBridgeReadsSupportedScalarLastPageWithNativePhaseEvidence'; Native=$true; Oracle='Offset 8/limit 6: Double/Currency/Date values, three refusals, three observers, terminal phases and normal exit'}
        @{Id='paramarray'; Method='ExcelProcedureValuesTests.NativeParamArrayCallsPreserveArityValuesAndSingleInvocation'; Native=$true; Oracle='Empty/30 scalar/Null/array ParamArray values, exact invocation counts; normal exit'}
        @{Id='variant-arrays'; Method='ExcelProcedureValuesTests.NativeVariantArraysRoundTripWithBoundsAndOneInvocation'; Native=$true; Oracle='Vector/matrix bounds and exact native values, no replay during polling; normal exit'}
        @{Id='save-fresh-reopen'; Method='ExcelQualificationPersistenceTests.ProductSavePreservesModuleClassAndFormInFreshOwnedProcessWithoutHelperSaving'; Native=$true; Oracle='Dirty state then product Save, module/class/form/Label disk readback in a fresh process, no helper Save; both normal exits'}
        @{Id='protection-reopen'; Method='ExcelIdeSurfaceTests.ProjectProtectionPersistsAfterNativeSaveAndReopen'; Native=$true; Oracle='Native project lock after product Save and discard-close/reopen; normal exit'}
    )
    foreach ($scenario in $scenarios) { if ($scenario.Method) { $scenario.Filter = 'FullyQualifiedName=VBAi.Tests.Integration.' + $scenario.Method } }
    if($RecheckDesktopGuardOnly){$scenarios[0].Filter='FullyQualifiedName~IsolatedTestDesktopTests';$scenarios[0].Oracle='Recheck only the changed Office command-line construction and desktop argument validation; earlier other managed guards are separate retained evidence'}
    if($ScenarioIds){
        $selected=@($ScenarioIds.Split(','))
        foreach($id in $selected){if(-not @($scenarios | Where-Object {$_.Id -ceq $id}).Count){throw "Unknown prepared scenario: $id"}}
        $scenarios=@($scenarios | Where-Object {$selected -ccontains $_.Id})
    }
    $files = @(Get-ChildItem -LiteralPath ([IO.Path]::GetDirectoryName($testAssembly)) -File | Where-Object { $_.Extension -in @('.dll','.exe','.config') } | ForEach-Object { @{Path=$_.FullName;Sha256=(Get-FileHash -LiteralPath $_.FullName).Hash} })
    $helper = Join-Path $build 'VBAi.Desktop.Helper/Debug/net48/VBAi.Desktop.Helper.exe'
    $files += @{Path=$helper;Sha256=(Get-FileHash -LiteralPath $helper).Hash}
    $files += @{Path=$script;Sha256=(Get-FileHash -LiteralPath $script).Hash}
    $plan = @{Scope='Q-006 owned Excel evidence; not all-host release acceptance or historical crash causality';Repository=$repository;SourceCommit=(& git -C $repository rev-parse HEAD);
        SourceStatus=(& git -C $repository status --porcelain);InstalledProduct=(Join-Path $InstalledDirectory 'VBAi.dll');ProductSha256=$hash;
        EvidenceRoot=$EvidenceRoot;BuildRoot=$build;TestAssembly=$testAssembly;HelperAssembly=$helper;CampaignScript=$script;FrozenFiles=$files;Scenarios=$scenarios;
        NoAutomaticRetries=$true;NoDesktopSwitch=$true;NoForceTermination=$true;PreparedUtc=[DateTime]::UtcNow.ToString('o')}
    Write-Json (Join-Path $EvidenceRoot 'q006-plan.json') $plan
    $plan | ConvertTo-Json -Depth 4
    exit 0
}
$plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $env:VBAi_TEST_DESKTOP_NAME) { throw 'Execution requires the reviewed isolated-desktop worker; there is no interactive fallback.' }
foreach ($file in $plan.FrozenFiles) { if ((Get-FileHash -LiteralPath $file.Path).Hash -cne $file.Sha256) { throw "Frozen candidate changed: $($file.Path)" } }
if ((Get-FileHash -LiteralPath $plan.InstalledProduct).Hash -cne $plan.ProductSha256) { throw 'Installed product changed after preparation.' }
$ledger = Join-Path $plan.EvidenceRoot 'campaign.json'
if (Test-Path -LiteralPath $ledger) { throw 'This campaign is one-shot; previous outcomes must not be overwritten or replayed.' }
$records = @(); $allPassed = $true; $refused = $false
foreach ($scenario in $plan.Scenarios) {
    $record = @{Id=$scenario.Id;Filter=$scenario.Filter;Oracle=$scenario.Oracle;State='NOT_RUN';Utc=[DateTime]::UtcNow.ToString('o')}
    $records += $record
    if ($refused) { Write-Json $ledger @{Scenarios=$records;Completed=$false;NoRetries=$true}; continue }
    if (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -ne 0) { $record.State='REFUSED_EXISTING_HOST'; $refused=$true; $allPassed=$false; Write-Json $ledger @{Scenarios=$records;Completed=$false;NoRetries=$true}; continue }
    $root = Join-Path $plan.EvidenceRoot $scenario.Id
    [IO.Directory]::CreateDirectory($root) | Out-Null
    $env:VBAi_VBE_INSPECTION_TRACE = Join-Path $root 'inspection.jsonl'
    $env:VBAi_EXCEL_RESULTS = Join-Path $root 'hosts'
    $env:VBAi_RUN_EXCEL_TESTS = if ($scenario.Native) { '1' } else { '0' }
    $record.State='STARTED_ONCE'; Write-Json $ledger @{Scenarios=$records;Completed=$false;NoRetries=$true}
    & dotnet vstest $plan.TestAssembly "/TestCaseFilter:$($scenario.Filter)" "/Logger:trx;LogFileName=$($scenario.Id).trx" "/ResultsDirectory:$root" *> (Join-Path $root 'test.log')
    $record.ExitCode=$LASTEXITCODE
    $trx = Join-Path $root ($scenario.Id + '.trx')
    if (Test-Path -LiteralPath $trx) {
        [xml]$report = Get-Content -LiteralPath $trx
        $record.Counters = $report.TestRun.ResultSummary.Counters.OuterXml
        $record.State = if ($record.ExitCode -eq 0 -and [int]$report.TestRun.ResultSummary.Counters.executed -gt 0 -and [int]$report.TestRun.ResultSummary.Counters.total -eq [int]$report.TestRun.ResultSummary.Counters.passed) { 'PASS' } else { 'FAILED_OR_SKIPPED' }
    } else { $record.State='NO_TERMINAL_REPORT' }
    if ($record.State -ne 'PASS') { $allPassed=$false }
    $remaining = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $record.RemainingExcelPids=$remaining
    if ($remaining.Count -ne 0 -or $record.State -eq 'NO_TERMINAL_REPORT' -or (-not $scenario.Native -and $record.State -ne 'PASS')) { $refused=$true }
    Write-Json $ledger @{Scenarios=$records;Completed=$false;NoRetries=$true}
}
Write-Json $ledger @{Scenarios=$records;Completed=$true;AllPassed=$allPassed;NoRetries=$true;EndedUtc=[DateTime]::UtcNow.ToString('o')}
if (-not $allPassed) { exit 1 }
exit 0

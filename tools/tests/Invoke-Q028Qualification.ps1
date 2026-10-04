#requires -Version 5.1
param([switch]$Prepare, [string]$EvidenceRoot, [string]$ModelRoot, [string]$ProductSourceCommit,
    [string]$Model = 'qwen2.5:7b-instruct',
    [string]$OllamaExecutable = 'C:\Users\init-\AppData\Local\Programs\Ollama\ollama.exe')
$ErrorActionPreference = 'Stop'
function Write-Json($path, $value) { $value | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $path -Encoding UTF8 }
function Free-Port {
    $socket = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
    try { $socket.Start(); return $socket.LocalEndpoint.Port } finally { $socket.Stop() }
}
if ($Prepare) {
    if (-not [IO.Path]::IsPathRooted($EvidenceRoot) -or (Test-Path -LiteralPath $EvidenceRoot)) { throw 'Fresh absolute evidence root required.' }
    $repo = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
    $product = Join-Path $repo 'artifacts/build/VBAi/Debug/net48/VBAi.dll'
    $tests = Join-Path $repo 'artifacts/build/VBAi.Tests/Debug/net48/VBAi.Tests.dll'
    $helper = Join-Path $repo 'artifacts/build/VBAi.Desktop.Helper/Debug/net48/VBAi.Desktop.Helper.exe'
    $manifest = Join-Path $ModelRoot ('manifests/registry.ollama.ai/library/' + $Model.Replace(':','/'))
    foreach ($path in @($product,$tests,$helper,$OllamaExecutable,$manifest)) {
        if (-not [IO.Path]::IsPathRooted($path) -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing absolute prerequisite: $path" }
    }
    [IO.Directory]::CreateDirectory($EvidenceRoot) | Out-Null
    $payload = Get-Content -LiteralPath $manifest -Raw -Encoding UTF8 | ConvertFrom-Json
    $modelFiles = @($manifest)
    foreach ($blob in @($payload.config) + @($payload.layers)) {
        $path = Join-Path $ModelRoot ('blobs/' + $blob.digest.Replace(':','-'))
        if ((Get-Item -LiteralPath $path).Length -ne $blob.size) { throw 'Model blob size differs from its manifest.' }
        if ((Get-FileHash -LiteralPath $path).Hash -ine $blob.digest.Substring(7)) { throw 'Model blob content hash differs.' }
        $modelFiles += $path
    }
    $hosts = @()
    $toolExecutables = @('node.exe','dotnet.exe','git.exe' | ForEach-Object {
        $tool = (Get-Command $_ -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
        @{Path=$tool;Sha256=(Get-FileHash -LiteralPath $tool).Hash}
    })
    foreach ($name in @('Excel','Word','PowerPoint','Access','Publisher','Outlook')) {
        $exeName = @{Excel='EXCEL';Word='WINWORD';PowerPoint='POWERPNT';Access='MSACCESS';Publisher='MSPUB';Outlook='OUTLOOK'}[$name]
        $exe = Join-Path 'C:/Program Files/Microsoft Office/root/Office16' ($exeName + '.EXE')
        $hosts += @{Name=$name;Executable=$exe;ProcessName=$exeName;Version=(Get-Item -LiteralPath $exe).VersionInfo.FileVersion;Sha256=(Get-FileHash -LiteralPath $exe).Hash}
    }
    if (-not $ProductSourceCommit) { $ProductSourceCommit = & git -C $repo rev-parse HEAD }
    $managedClasses = @('OllamaQualificationEndpointTests','OllamaQualificationModelTests',
        'OllamaQualificationProfileTests','OllamaSyntheticWireCaptureTests','OllamaOfficeStreamOracleTests',
        'LlmChatClientCoverageTests','StreamTests','ChatWindowStateTests','LlmVbeToolsBoundaryTests',
        'LlmVbeAsyncValidationTests','LlmVbeToolContractTests','LlmProjectPrivacyTests','ProjectPrivacyBoundaryTests',
        'CatalogBoundaryTests','ToolCatalogTests','PrivateDesktopUiActionTests','QualificationDesktopGuardTests',
        'OfficeVbeFixtureDesktopTests','OfficeVbeFixtureDesktopStartupRecoveryTests','OutlookPrivateDesktopTests',
        'OfficeOwnedShutdownEvidenceTests','OutlookVbaTestFixtureShutdownTests')
    $managedFilter = '(TestCategory=Unit|TestCategory=Scenario)&TestCategory!=OllamaUi&(' +
        (($managedClasses | ForEach-Object {'FullyQualifiedName~VBAi.Tests.Unit.'+$_+'.'}) -join '|') +
        '|FullyQualifiedName~VBAi.Tests.NativeExportTraceTests.)'
    $scenarios = @(
        @{Id='managed';Filter=$managedFilter;Oracle='Focused transport/chat/privacy/desktop/Office lifecycle regressions all pass; includes the Node lookup regression. This gate does not replace the failed broad aggregate. The real-provider OllamaUi case runs separately; native/provider opt-ins are absent'}
        @{Id='tool-roundtrip';Filter='FullyQualifiedName=VBAi.Tests.Integration.OllamaQualificationTests.LocalModelStreamsAndCompletesSyntheticToolRoundTrip';Oracle='Visible HTTP deltas, exactly one qualification_echo with scalar VB_AI_42, exact final marker; retained synthetic wire'}
        @{Id='cancel-recovery';Filter='FullyQualifiedName=VBAi.Tests.Integration.OllamaQualificationTests.LocalModelCancellationDoesNotPoisonTheNextConversation';Oracle='Cancel after first nonempty fragment; independent fresh request completes; exact wire retained'}
        @{Id='detached-ui';Filter='FullyQualifiedName=VBAi.Tests.Unit.ChatWindowStateTests.LocalOllamaShownChatStreamsStopsAndCompletesNextSend';Oracle='Shown real controls render text while busy, one Stop, visible interruption and next complete UI_READY_42; simulated VBE explicitly distinct'}
    )
    foreach ($hostRow in $hosts) {
        $scenarios += @{Id=($hostRow.Name.ToLowerInvariant()+'-embedded');Host=$hostRow.Name;
            Filter=('FullyQualifiedName=VBAi.Tests.Integration.OllamaOfficeQualificationTests.'+$hostRow.Name+'EmbeddedAssistantStreamsStopsRecoversAndReadsNativeMarker');
            Oracle='Exact installed candidate and private native owner; actual embedded assistant and own scope; streamed text while Stop is active; one Stop; visible cancellation and next reply; unprompted native marker via read_module; all source/references unchanged; original normal host exit. No VBA execution or mail send.'}
    }
    $frozen = Join-Path $EvidenceRoot 'Invoke-FrozenQ028.ps1'
    Copy-Item -LiteralPath $PSCommandPath -Destination $frozen
    $files = @($frozen,$OllamaExecutable) + $modelFiles
    $files += @($toolExecutables.Path)
    $files += @(Get-ChildItem (Split-Path $product),(Split-Path $tests),(Split-Path $helper) -Recurse -File | Select-Object -ExpandProperty FullName)
    $files += @((Join-Path $repo 'tools/testing-explorer/Set-TestExplorerCandidate.ps1'),(Join-Path $repo 'tools/tests/Invoke-IsolatedDesktopWorker.ps1'))
    $backendPort = Free-Port; $proxyPort = Free-Port
    if ($backendPort -eq $proxyPort) { throw 'Port allocation collided; no campaign launched.' }
    $settings = Join-Path $EvidenceRoot 'bounded.runsettings'
    '<RunSettings><RunConfiguration><TargetPlatform>x64</TargetPlatform><MaxCpuCount>1</MaxCpuCount><TestSessionTimeout>900000</TestSessionTimeout><TreatNoTestsAsError>true</TreatNoTestsAsError></RunConfiguration></RunSettings>' | Set-Content -LiteralPath $settings -Encoding UTF8
    $files += $settings
    $plan = @{Scope='Q028 real Ollama and six classic Office VBE hosts; no SOLIDWORKS, Visio, Project or global VBE preference changes';
        ProductSourceCommit=$ProductSourceCommit;RunSettings=$settings;CaseTimeoutMilliseconds=900000;
        SourceCommit=(& git -C $repo rev-parse HEAD);SourceStatus=@(& git -C $repo status --porcelain);Repository=$repo;
        EvidenceRoot=$EvidenceRoot;Product=$product;TestAssembly=$tests;Helper=$helper;
        ProductMvid=([Reflection.Assembly]::ReflectionOnlyLoadFrom($product).ManifestModule.ModuleVersionId.ToString('D'));
        ProductSha256=(Get-FileHash $product).Hash;OllamaExe=$OllamaExecutable;OllamaVersion=(Get-Item $OllamaExecutable).VersionInfo.ProductVersion;
        ModelRoot=$ModelRoot;Model=$Model;ModelDigest=(Get-FileHash $manifest).Hash;BackendPort=$backendPort;ProxyPort=$proxyPort;
        Temperature=0;TopP=0.8;ContextLength=8192;NumParallel=1;Device='CPU';CloudDisabled=$true;
        Hosts=$hosts;Scenarios=$scenarios;NoRetry=$true;NoDesktopSwitch=$true;NoForceTerminationOfOffice=$true;
        ToolExecutables=$toolExecutables;
        BackendShutdown='Stop only the exact newly created synthetic headless Ollama server after requests settle; never call this a normal Office exit';
        FrozenFiles=@($files | Select-Object -Unique | ForEach-Object {@{Path=$_;Sha256=(Get-FileHash -LiteralPath $_).Hash}});
        PreparedUtc=[DateTime]::UtcNow.ToString('o')}
    Write-Json (Join-Path $EvidenceRoot 'q028-plan.json') $plan
    Write-Output (Join-Path $EvidenceRoot 'q028-plan.json'); exit 0
}

$plan = Get-Content (Join-Path $PSScriptRoot 'q028-plan.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if ($PSVersionTable.PSEdition -ne 'Desktop' -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA' -or -not $env:VBAi_TEST_DESKTOP_NAME) { throw 'The reviewed private x64 Desktop PowerShell STA worker is required.' }
foreach ($file in $plan.FrozenFiles) { if ((Get-FileHash -LiteralPath $file.Path).Hash -cne $file.Sha256) { throw "Frozen file changed: $($file.Path)" } }
$root = $plan.EvidenceRoot
$ledgerPath = Join-Path $root 'campaign.json'
if (Test-Path -LiteralPath $ledgerPath) { throw 'This campaign is already claimed; no replay.' }
$rows = @($plan.Scenarios | ForEach-Object { [pscustomobject]@{Id=$_.Id;Host=$_.Host;Oracle=$_.Oracle;State='NOT_RUN';InvocationCount=0;Trx=$null;Error=$null} })
$ledger = @{State='RUNNING';Qualified=$false;Desktop=$env:VBAi_TEST_DESKTOP_NAME;Scenarios=$rows;ProductMvid=$plan.ProductMvid;ProductSha256=$plan.ProductSha256;StartedUtc=[DateTime]::UtcNow.ToString('o')}
function Flush { Write-Json $ledgerPath $ledger }
Flush
$proxy = $null; $server = $null; $settingsBaseline = $null; $settingsApplied = $null; $registered = $false
$backup = Join-Path $root 'candidate-registration.clixml'
$register = Join-Path $plan.Repository 'tools/testing-explorer/Set-TestExplorerCandidate.ps1'
$settingsFields = @('ProviderName','OllamaEndpoint','OllamaModel','OllamaTemperature','OllamaTopP','VbeEditApproval')
function Run-Case($scenario, $row) {
    $row.State='STARTED_ONCE';$row.InvocationCount=1;Flush
    $result = Join-Path $root $scenario.Id
    [IO.Directory]::CreateDirectory($result) | Out-Null
    $row.Trx=Join-Path $result 'result.trx'
    & dotnet vstest $plan.TestAssembly "/Settings:$($plan.RunSettings)" "/TestCaseFilter:$($scenario.Filter)" "/ResultsDirectory:$result" '/Logger:trx;LogFileName=result.trx' *> (Join-Path $result 'console.log')
    $testExitCode = $LASTEXITCODE
    if (-not (Test-Path -LiteralPath $row.Trx)) { $row.State='NO_TERMINAL_REPORT';Flush;return }
    [xml]$trx = Get-Content -LiteralPath $row.Trx -Raw -Encoding UTF8
    $c = $trx.TestRun.ResultSummary.Counters
    $row.State = if ($testExitCode -eq 0 -and [int]$c.total -gt 0 -and [int]$c.total -eq [int]$c.passed) { 'PASS' } else { 'FAILED_OR_SKIPPED' }
    Flush
}
try {
    # Scheduled GUI tasks do not inherit Codex's process-local Node runtime path.
    # Freeze the exact tools in Prepare, prepend only those directories in this worker.
    $env:PATH = ((@($plan.ToolExecutables.Path | ForEach-Object {Split-Path $_}) | Select-Object -Unique) -join ';')+';'+$env:PATH
    # Strip inherited qualification opt-ins; keep only the private desktop's identity/sentinel.
    foreach ($entry in @(Get-ChildItem Env: | Where-Object {$_.Name -match '^VBAi_RUN_|^VBAi_TEST_|^VBAi_OLLAMA_'})) {
        if ($entry.Name -notmatch '^VBAi_TEST_DESKTOP_') { Remove-Item -LiteralPath ('Env:\'+$entry.Name) }
    }
    Run-Case $plan.Scenarios[0] $rows[0]
    if ($rows[0].State -ne 'PASS') { throw 'Managed gate failed; all real provider/native cases remain NOT_RUN.' }
    $start = [Diagnostics.ProcessStartInfo]::new($plan.OllamaExe,'serve')
    $start.UseShellExecute=$false;$start.CreateNoWindow=$true;$start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true
    foreach ($key in @($start.EnvironmentVariables.Keys)) {
        if ($key -match '(?i)(_KEY|_TOKEN|_SECRET|^CODEX_|^OLLAMA_)') { $start.EnvironmentVariables.Remove($key) }
    }
    $start.EnvironmentVariables['OLLAMA_HOST']='127.0.0.1:'+$plan.BackendPort
    $start.EnvironmentVariables['OLLAMA_MODELS']=$plan.ModelRoot
    $start.EnvironmentVariables['OLLAMA_CONTEXT_LENGTH']='8192';$start.EnvironmentVariables['OLLAMA_NUM_PARALLEL']='1'
    $start.EnvironmentVariables['OLLAMA_NO_CLOUD']='1';$start.EnvironmentVariables['OLLAMA_NOPRUNE']='1'
    $start.EnvironmentVariables['OLLAMA_LLM_LIBRARY']='cpu';$start.EnvironmentVariables['OLLAMA_VULKAN']='false'
    $start.EnvironmentVariables['CUDA_VISIBLE_DEVICES']='-1';$start.EnvironmentVariables['ROCR_VISIBLE_DEVICES']='-1'
    $server=[Diagnostics.Process]::new();$server.StartInfo=$start
    Write-Json (Join-Path $root 'backend-start-intent.json') @{InvocationLimit=1;Exe=$plan.OllamaExe;Port=$plan.BackendPort;ModelDigest=$plan.ModelDigest;Profile='CPU, context8192, parallel1, cloud disabled'}
    if (-not $server.Start()) { throw 'Backend creation returned false; no retry.' }
    $originalBackendHandle=$server.Handle
    $stdout=$server.StandardOutput.ReadToEndAsync();$stderr=$server.StandardError.ReadToEndAsync()
    Write-Json (Join-Path $root 'backend-started.json') @{ProcessId=$server.Id;StartUtc=$server.StartTime.ToUniversalTime().ToString('o');OriginalHandleHeld=($originalBackendHandle -ne [IntPtr]::Zero)}
    $base='http://127.0.0.1:'+$plan.BackendPort
    $ready=[Diagnostics.Stopwatch]::StartNew();$tags=$null
    while ($ready.Elapsed.TotalSeconds -lt 30 -and $null -eq $tags) {
        if ($server.HasExited) { throw 'Owned Ollama exited during readiness.' }
        try { $tags=Invoke-RestMethod ($base+'/api/tags') -TimeoutSec 2 } catch { Start-Sleep -Milliseconds 200 }
    }
    if (@($tags.models | Where-Object {$_.name -ceq $plan.Model -and $_.digest -ieq $plan.ModelDigest}).Count -ne 1) { throw 'The exact frozen model was not advertised; no download/substitution.' }
    Write-Json (Join-Path $root 'backend-attestation.json') @{Version=(Invoke-RestMethod ($base+'/api/version'));Models=$tags;CloudDisabled=$true;RequestedDevice='CPU'}
    $env:VBAi_TEST_OLLAMA_ENDPOINT=$base+'/v1/chat/completions';$env:VBAi_TEST_OLLAMA_MODEL=$plan.Model
    $env:VBAi_TEST_OLLAMA_TEMPERATURE='0';$env:VBAi_TEST_OLLAMA_TOP_P='0.8'
    $env:VBAi_RUN_OLLAMA_TESTS='1';$env:VBAi_OLLAMA_HEADLESS_CAPTURE_WIRE='1';$env:VBAi_OLLAMA_HEADLESS_RESULTS=Join-Path $root 'headless-wire'
    foreach ($i in @(1,2)) { Run-Case $plan.Scenarios[$i] $rows[$i]; if ($rows[$i].State -ne 'PASS') { throw ('Provider gate failed: '+$rows[$i].Id) } }
    $env:VBAi_RUN_OLLAMA_TESTS=$null;$env:VBAi_RUN_OLLAMA_UI_TESTS='1';$env:VBAi_OLLAMA_UI_CAPTURE_WIRE='1';$env:VBAi_OLLAMA_UI_RESULTS=Join-Path $root 'detached-wire'
    Run-Case $plan.Scenarios[3] $rows[3]
    if ($rows[3].State -ne 'PASS') { throw 'Detached chat gate failed; native banks remain NOT_RUN.' }
    $env:VBAi_RUN_OLLAMA_UI_TESTS=$null
    $loadedModels=Invoke-RestMethod ($base+'/api/ps')
    Write-Json (Join-Path $root 'loaded-model-device.json') $loadedModels
    if (@($loadedModels.models | Where-Object {$_.name -ceq $plan.Model -and $_.size_vram -eq 0}).Count -ne 1) { throw 'CPU inference was not independently observed.' }
    foreach ($hostRow in $plan.Hosts) {
        if (@(Get-Process $hostRow.ProcessName -ErrorAction SilentlyContinue).Count -ne 0) { throw ('Existing '+$hostRow.Name+' must be preserved; no native launch.') }
    }
    [Reflection.Assembly]::LoadFrom($plan.TestAssembly) | Out-Null
    $proxy=New-Object VBAi.Tests.Integration.OllamaQualificationProxy($env:VBAi_TEST_OLLAMA_ENDPOINT,[int]$plan.ProxyPort,(Join-Path $root 'embedded-wire'))
    $env:VBAi_TEST_OLLAMA_ENDPOINT=$proxy.Endpoint
    $productAssembly=[Reflection.Assembly]::LoadFrom($plan.Product)
    $settingsType=$productAssembly.GetType('VBAi.LlmSettings',$true)
    $load=$settingsType.GetMethod('Load');$save=$settingsType.GetMethod('Save')
    $settings=$load.Invoke($null,@());$settingsBaseline=@{}
    foreach ($field in $settingsFields) { $settingsBaseline[$field]=$settingsType.GetProperty($field).GetValue($settings,$null) }
    $settingsApplied=@{ProviderName='Ollama';OllamaEndpoint=$proxy.Endpoint;OllamaModel=$plan.Model;OllamaTemperature=[double]0;OllamaTopP=[double]0.8;VbeEditApproval='ReadOnly'}
    Write-Json (Join-Path $root 'settings-intent.json') @{Baseline=$settingsBaseline;Applied=$settingsApplied;SecretsDecryptedOrCaptured=$false;OnlyChangedFields=$settingsFields}
    foreach ($field in $settingsFields) { $settingsType.GetProperty($field).SetValue($settings,$settingsApplied[$field],$null) }
    $save.Invoke($settings,@()) | Out-Null
    $registered=$true
    $ledger.RegistrationApply=& $register -CandidateAssemblyPath $plan.Product -ExpectedMvid ([Guid]$plan.ProductMvid) -Apply -ReportPath $backup
    $env:VBAi_QUALIFICATION_DESKTOP=$env:VBAi_TEST_DESKTOP_NAME
    $env:VBAi_RUN_OLLAMA_OFFICE_TESTS='1';$env:VBAi_RUN_EXCEL_TESTS='1';$env:VBAi_RUN_OFFICE_TESTS='1';$env:VBAi_RUN_OUTLOOK_TESTS='1'
    $env:VBAi_Q028_RESULTS=Join-Path $root 'host-results';$env:VBAi_EXCEL_RESULTS=Join-Path $root 'hosts/Excel';$env:VBAi_OFFICE_RESULTS=Join-Path $root 'hosts/Office';$env:VBAi_OUTLOOK_RESULTS=Join-Path $root 'hosts/Outlook'
    foreach ($hostRow in $plan.Hosts) {
        [Environment]::SetEnvironmentVariable('VBAi_TEST_'+$hostRow.Name.ToUpperInvariant()+'_EXE',$hostRow.Executable,'Process')
        [Environment]::SetEnvironmentVariable('VBAi_TEST_'+$hostRow.Name.ToUpperInvariant()+'_EXE_SHA256',$hostRow.Sha256,'Process')
    }
    for ($i=4;$i -lt $plan.Scenarios.Count;$i++) {
        foreach ($hostRow in $plan.Hosts) { if (@(Get-Process $hostRow.ProcessName -ErrorAction SilentlyContinue).Count -ne 0) { throw 'A preceding/foreign Office host is live; no next bank.' } }
        $env:VBAi_VBE_INSPECTION_TRACE=Join-Path $root ('hosts/'+$plan.Scenarios[$i].Host+'/inspection.jsonl')
        [IO.Directory]::CreateDirectory((Split-Path $env:VBAi_VBE_INSPECTION_TRACE)) | Out-Null
        Run-Case $plan.Scenarios[$i] $rows[$i]
        if ($rows[$i].State -ne 'PASS') { throw ('Native bank failed: '+$rows[$i].Id+'; remaining banks are NOT_RUN.') }
    }
    $ledger.State='FUNCTIONAL_MATRIX_PASS_PENDING_OFFLINE_WIRE_REVIEW'
} catch { $ledger.State='FAILED_OR_BLOCKED';$ledger.Error=$_.Exception.ToString() }
finally {
    $ownedOfficeLive=@($plan.Hosts | Where-Object {@(Get-Process $_.ProcessName -ErrorAction SilentlyContinue).Count -ne 0})
    if ($ownedOfficeLive.Count -eq 0) {
        if ($settingsBaseline -and $settingsApplied) {
            try {
                $current=$load.Invoke($null,@())
                foreach ($field in $settingsFields) {
                    if ($settingsType.GetProperty($field).GetValue($current,$null) -cne $settingsApplied[$field]) { throw ('Concurrent settings change; restore refused: '+$field) }
                }
                foreach ($field in $settingsFields) { $settingsType.GetProperty($field).SetValue($current,$settingsBaseline[$field],$null) }
                $save.Invoke($current,@()) | Out-Null
                $readback=$load.Invoke($null,@())
                foreach ($field in $settingsFields) { if ($settingsType.GetProperty($field).GetValue($readback,$null) -cne $settingsBaseline[$field]) { throw 'Settings restoration readback differs.' } }
                $ledger.SettingsRestored=$true
            } catch { $ledger.SettingsRestoreError=$_.Exception.ToString();$ledger.State='FAILED_OR_BLOCKED' }
        }
        if ($registered -and (Test-Path -LiteralPath ($backup+'.after.clixml'))) {
            try { $ledger.RegistrationRestore=& $register -Restore -BackupPath $backup }
            catch { $ledger.RegistrationRestoreError=$_.Exception.ToString();$ledger.State='FAILED_OR_BLOCKED' }
        }
        if ($proxy) {
            try {
                $settle=[Diagnostics.Stopwatch]::StartNew()
                while (-not $proxy.AllRequestsSettled -and $settle.Elapsed.TotalSeconds -lt 30) { Start-Sleep -Milliseconds 100 }
                $proxy.Dispose();$ledger.ProxyClosed=$true
            } catch { $ledger.ProxyCloseError=$_.Exception.ToString();$ledger.State='FAILED_OR_BLOCKED' }
        }
        if ($server -and -not $server.HasExited -and (-not $proxy -or $ledger.ProxyClosed)) {
            # This ephemeral server has no user requests, documents or shutdown API. Its held
            # Process object identifies the one process created above; no PID/name sweep is used.
            Write-Json (Join-Path $root 'backend-stop-intent.json') @{ProcessId=$server.Id;OriginalHandleHeld=($originalBackendHandle -ne [IntPtr]::Zero);OutstandingProxyRequests=0;NoOfficeTermination=$true}
            $server.Kill()
            if (-not $server.WaitForExit(15000)) { $ledger.BackendRetainedPid=$server.Id;$ledger.State='FAILED_OR_UNCERTAIN' }
            else {
                Write-Json (Join-Path $root 'backend-exit.json') @{ProcessId=$server.Id;ExitCode=$server.ExitCode;Scope='Owned headless test server only; not normal Office exit'}
                if ($stdout.Wait(5000)) { [IO.File]::WriteAllText((Join-Path $root 'backend.stdout.log'),$stdout.GetAwaiter().GetResult()) }
                if ($stderr.Wait(5000)) { [IO.File]::WriteAllText((Join-Path $root 'backend.stderr.log'),$stderr.GetAwaiter().GetResult()) }
                $server.Dispose();$ledger.BackendStopped=$true
            }
        }
    } else { $ledger.RetainedOffice=@($ownedOfficeLive.Name);$ledger.State='FAILED_OR_UNCERTAIN';$ledger.NoTeardownAfterUncertainty=$true }
    $ledger.CompletedUtc=[DateTime]::UtcNow.ToString('o');Flush
}
if ($ledger.State -ne 'FUNCTIONAL_MATRIX_PASS_PENDING_OFFLINE_WIRE_REVIEW') { throw $ledger.Error }

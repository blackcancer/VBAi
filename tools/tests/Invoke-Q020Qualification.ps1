#requires -Version 5.1
param(
    [switch]$Prepare,
    [string]$EvidenceRoot,
    [string]$InstalledDirectory,
    [string]$BuildOutputRoot,
    [string]$ManagedEvidenceTrxPath,
    [string]$Q030ManagedEvidenceTrxPath,
    [ValidateSet(2019,2025)][int]$SolidWorksYear = 2019,
    [string]$ExpectedNativeRevision = '27.5.0',
    [string]$SolidWorksExecutable = 'D:\Program Files\SOLIDWORKS 2019 Corp\SOLIDWORKS\SLDWORKS.exe',
    [string]$VisualStudioExecutable = 'C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\devenv.exe'
)
$ErrorActionPreference = 'Stop'
function Write-Json($path, $value) { $value | ConvertTo-Json -Depth 24 | Set-Content -LiteralPath $path -Encoding UTF8 }
function Read-OwnedFileHash([string]$path) {
    $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    $sha=[Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','') }
    finally { $sha.Dispose();$stream.Dispose() }
}
function Read-Q020ManagedEvidence([string]$Path,[string]$TestAssembly) {
    [xml]$trx=Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $c=$trx.TestRun.ResultSummary.Counters
    if([int]$c.total -le 0 -or [int]$c.total -ne [int]$c.passed -or [int]$c.total -ne [int]$c.executed -or @($trx.TestRun.Results.UnitTestResult|Where-Object outcome -ne 'Passed').Count -ne 0) { throw 'Managed evidence missing/skipped/failed.' }
    $names='VbeProjectLifecycleTests|Standalone|VbeSolidWorksPersistenceTests|ListProjectsPreservesOriginalFileNameGetterDiagnostics|ListProjectsRetainsUnsavedProjectWhenFileNameIsUnavailable'
    foreach($method in $trx.TestRun.TestDefinitions.UnitTest.TestMethod) {
        if($method.codeBase -ine $TestAssembly -or ($method.className+'.'+$method.name) -notmatch $names -or $method.className -notmatch '^VBAi\.Tests\.') { throw 'Managed evidence is outside the original scoped filter/assembly.' }
    }
    return [int]$c.total
}
function Read-Q030ManagedEvidence([string]$Path,[string]$TestAssembly) {
    [xml]$trx=Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $c=$trx.TestRun.ResultSummary.Counters
    if([int]$c.total -le 0 -or [int]$c.total -ne [int]$c.passed -or [int]$c.total -ne [int]$c.executed -or @($trx.TestRun.Results.UnitTestResult|Where-Object outcome -ne 'Passed').Count -ne 0) { throw 'Q030 managed evidence missing/skipped/failed.' }
    foreach($method in $trx.TestRun.TestDefinitions.UnitTest.TestMethod) {
        if($method.codeBase -ine $TestAssembly -or $method.className -notmatch '^VBAi\.Tests\.Unit\.' -or ($method.className+'.'+$method.name) -notmatch 'ToolCatalogTests|CatalogBoundaryTests|ProjectMetadataAuthorizationRefusesRevokedDirectAndCatalogWrites|ScopeAndProjectBindingApplyToAsyncInvocationBeforeHostAccess') { throw 'Q030 managed evidence is outside the scoped Unit filter/assembly.' }
    }
    return [int]$c.total
}
if ($Prepare) {
    if (-not [IO.Path]::IsPathRooted($EvidenceRoot) -or -not (Test-Path -LiteralPath $EvidenceRoot -PathType Container)) { throw 'An existing absolute build/evidence root is required.' }
    $planFile = Join-Path $EvidenceRoot 'q020-plan.json'
    if (Test-Path -LiteralPath $planFile) { throw 'Preparation already exists; do not overwrite a campaign.' }
    $repository = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    $installed = Join-Path $InstalledDirectory 'VBAi.dll'
    if(-not $BuildOutputRoot) { $BuildOutputRoot=Join-Path $EvidenceRoot 'build' }
    if(-not [IO.Path]::IsPathRooted($BuildOutputRoot) -or -not [IO.Directory]::Exists($BuildOutputRoot)) { throw 'Existing absolute BuildOutputRoot required.' }
    $testAssembly = Join-Path $BuildOutputRoot 'VBAi.Tests/Debug/net48/VBAi.Tests.dll'
    $helper = Join-Path $BuildOutputRoot 'VBAi.Desktop.Helper/Debug/net48/VBAi.Desktop.Helper.exe'
    foreach ($file in @($installed,$testAssembly,$helper,$SolidWorksExecutable,$VisualStudioExecutable)) {
        if (-not [IO.Path]::IsPathRooted($file) -or -not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Required absolute file missing: $file" }
    }
    $nativeFileVersion=(Get-Item -LiteralPath $SolidWorksExecutable).VersionInfo.FileVersion
    if($nativeFileVersion.Split('.')[0] -ne $(if($SolidWorksYear -eq 2019){'27'}else{'33'}) -or
        $nativeFileVersion.Split('.')[0] -ne $ExpectedNativeRevision.Split('.')[0]) { throw 'Selected SOLIDWORKS year/revision and executable differ.' }
    $hash = (Get-FileHash -LiteralPath $installed).Hash
    if ((Get-FileHash (Join-Path (Split-Path $testAssembly) 'VBAi.dll')).Hash -cne $hash) { throw 'The harness must reference the exact installed candidate.' }
    $managedEvidence=$null
    if($ManagedEvidenceTrxPath) {
        if(-not [IO.Path]::IsPathRooted($ManagedEvidenceTrxPath)) { throw 'Absolute managed TRX evidence required.' }
        $count=Read-Q020ManagedEvidence $ManagedEvidenceTrxPath $testAssembly
        $originalPlanPath=Join-Path (Split-Path (Split-Path $ManagedEvidenceTrxPath)) 'q020-plan.json'
        $originalPlan=Get-Content -LiteralPath $originalPlanPath -Raw -Encoding UTF8|ConvertFrom-Json
        if($originalPlan.ProductSha256 -cne $hash -or $originalPlan.TestAssembly -ine $testAssembly -or $originalPlan.SourceCommit -cne (& git -C $repository rev-parse HEAD)) { throw 'Managed evidence candidate/source revision mismatch.' }
        $sourceFiles=@($originalPlan.FrozenFiles|Where-Object {$_.Path.StartsWith((Join-Path $repository 'src')+'\',[StringComparison]::OrdinalIgnoreCase) -or $_.Path.StartsWith((Join-Path $repository 'tests')+'\',[StringComparison]::OrdinalIgnoreCase)})
        if($sourceFiles.Count -eq 0) { throw 'Original managed source manifest missing.' }
        foreach($entry in @($sourceFiles)+@($originalPlan.FrozenFiles|Where-Object {$_.Path -ieq $testAssembly -or $_.Path -ieq $helper})) { if((Get-FileHash -LiteralPath $entry.Path).Hash -cne $entry.Sha256) { throw 'Managed evidence source/test/helper bytes changed.' } }
        $managedEvidence=@{TrxPath=$ManagedEvidenceTrxPath;TrxSha256=(Get-FileHash -LiteralPath $ManagedEvidenceTrxPath).Hash;OriginalPlanPath=$originalPlanPath;OriginalPlanSha256=(Get-FileHash -LiteralPath $originalPlanPath).Hash;TestAssemblySha256=(Get-FileHash -LiteralPath $testAssembly).Hash;HelperSha256=(Get-FileHash -LiteralPath $helper).Hash;ProductSha256=$hash;SourceCommit=$originalPlan.SourceCommit;Total=$count;SourceFiles=$sourceFiles}
    }
    if(-not $Q030ManagedEvidenceTrxPath -or -not $managedEvidence) { throw 'The Q020/Q030 campaign requires pinned existing managed evidence for both gates.' }
    $q030Evidence=@{TrxPath=$Q030ManagedEvidenceTrxPath;TrxSha256=(Get-FileHash -LiteralPath $Q030ManagedEvidenceTrxPath).Hash;Total=(Read-Q030ManagedEvidence $Q030ManagedEvidenceTrxPath $testAssembly);TestAssemblySha256=(Get-FileHash -LiteralPath $testAssembly).Hash;ProductSha256=$hash;SourceProof=$managedEvidence.OriginalPlanPath}
    $profile = Join-Path $EvidenceRoot 'debug-profile'
    [IO.Directory]::CreateDirectory($profile) | Out-Null
    $project = @'
<?xml version="1.0" encoding="utf-8"?>
<Project DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <ItemGroup Label="ProjectConfigurations"><ProjectConfiguration Include="Debug|x64"><Configuration>Debug</Configuration><Platform>x64</Platform></ProjectConfiguration></ItemGroup>
  <PropertyGroup Label="Globals"><ProjectGuid>{D284FC87-71EA-467D-A1F0-D6B5FF917109}</ProjectGuid><RootNamespace>Q014SolidWorks</RootNamespace></PropertyGroup>
  <Import Project="$(VCTargetsPath)\Microsoft.Cpp.Default.props" />
  <PropertyGroup Condition="'$(Configuration)|$(Platform)'=='Debug|x64'" Label="Configuration"><ConfigurationType>Utility</ConfigurationType><PlatformToolset>v145</PlatformToolset></PropertyGroup>
  <Import Project="$(VCTargetsPath)\Microsoft.Cpp.props" />
  <PropertyGroup Condition="'$(Configuration)|$(Platform)'=='Debug|x64'">
    <LocalDebuggerCommand>__EXECUTABLE__</LocalDebuggerCommand><LocalDebuggerWorkingDirectory>__DIRECTORY__</LocalDebuggerWorkingDirectory>
    <LocalDebuggerDebuggerType>NativeOnly</LocalDebuggerDebuggerType><DebuggerFlavor>WindowsLocalDebugger</DebuggerFlavor>
    <OutDir>$(ProjectDir)out\</OutDir><IntDir>$(ProjectDir)obj\</IntDir>
  </PropertyGroup>
  <Import Project="$(VCTargetsPath)\Microsoft.Cpp.targets" />
</Project>
'@
    $project = $project.Replace('__EXECUTABLE__',[Security.SecurityElement]::Escape($SolidWorksExecutable)).Replace('__DIRECTORY__',[Security.SecurityElement]::Escape((Split-Path $SolidWorksExecutable)))
    Set-Content (Join-Path $profile 'Q014SolidWorks.vcxproj') $project -Encoding UTF8
    $solution = @'
Microsoft Visual Studio Solution File, Format Version 12.00
# Visual Studio Version 18
VisualStudioVersion = 18.0.0.0
MinimumVisualStudioVersion = 10.0.40219.1
Project("{8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}") = "Q014SolidWorks", "Q014SolidWorks.vcxproj", "{D284FC87-71EA-467D-A1F0-D6B5FF917109}"
EndProject
Global
	GlobalSection(SolutionConfigurationPlatforms) = preSolution
		Debug|x64 = Debug|x64
	EndGlobalSection
	GlobalSection(ProjectConfigurationPlatforms) = postSolution
		{D284FC87-71EA-467D-A1F0-D6B5FF917109}.Debug|x64.ActiveCfg = Debug|x64
		{D284FC87-71EA-467D-A1F0-D6B5FF917109}.Debug|x64.Build.0 = Debug|x64
	EndGlobalSection
	GlobalSection(SolutionProperties) = preSolution
		HideSolutionNode = FALSE
	EndGlobalSection
EndGlobal
'@
    [IO.File]::WriteAllText((Join-Path $profile 'Q014SolidWorks.sln'),("`r`n"+($solution -replace "\r?\n","`r`n")+"`r`n"),[Text.UTF8Encoding]::new($true))
    Copy-Item -LiteralPath $PSCommandPath -Destination (Join-Path $EvidenceRoot 'Invoke-FrozenQ020.ps1')
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Q014Native.cs') -Destination $EvidenceRoot
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Q014VisualStudio.cs') -Destination $EvidenceRoot
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Q014SolidWorks.cs') -Destination $EvidenceRoot
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Q020Startup.cs') -Destination $EvidenceRoot
    $scenarios = @(
        @{Id='managed';Oracle='Real scoped unit TRX; fake COM guards remain managed-only evidence'}
        @{Id='managed-q030';Oracle='Pinned actual catalogue/authorization/project-binding Unit TRX; managed only, no native catalogue claim'}
        @{Id='launch';Oracle='One private owned VS utility Debug.Start; exact retained PID/birth/image/revision/native window desktop'}
        @{Id='bootstrap';Oracle='One owned native NewMacro573 Type100 bootstrap solely to initialize VBE; excluded from Add101 and firstSaveAs evidence'}
        @{Id='load';Oracle='Exact connected bridge PID/path/MVID/SHA and AddIn ProgId; no Connect forcing'}
        @{Id='add101';Oracle='One VBProjects.Add101; unique new unsaved unprotected standalone, all preexisting identities preserved'}
        @{Id='unsaved';Oracle='Exact first FileName getter nullable HRESULT/type from list_projects plus separate descriptor and normalized persistence; no simulated COM observation'}
        @{Id='first-save-as';Oracle='One explicit product SaveAs fresh original .swp, path/bytes/Saved readback; no replay after failure'}
        @{Id='components';Oracle='Unique synthetic module/class/form source and native Label geometry/caption; all baseline components preserved'}
        @{Id='compile';Oracle='One native compile, no diagnostic, design mode; no procedure execution'}
        @{Id='save';Oracle='One product Save after changes, exact requested path, Saved and file hash; full code/reference/designer snapshot'}
        @{Id='q030-refuse-unsafe-open';Oracle='Exact product SOLIDWORKS VBProjects.Open refusal; unchanged collection/source/references/form/original file and retained host; no native call-count instrumentation claim'}
        @{Id='close-standalone';Oracle='One VBProjects.Remove of saved Type101 with revision/path; original closed-file SHA unchanged'}
        @{Id='reload';Oracle='Q020/Q030: one native EditMacro84 opens SAME original file, full code/ref/form Label readback and unchanged disk hash; product standalone Open remains unavailable'}
        @{Id='cleanup';Oracle='One normal owned SOLIDWORKS ExitApp and original-handle exit0; owned idle VS Quit/exit0; no force or retry'}
    )
    $files = @($installed,$testAssembly,$helper,(Join-Path $EvidenceRoot 'Invoke-FrozenQ020.ps1'),(Join-Path $EvidenceRoot 'Q014Native.cs'),(Join-Path $EvidenceRoot 'Q014VisualStudio.cs'),(Join-Path $repository 'tools/Invoke-VBAi.ps1'),(Join-Path $repository 'tools/tests/Invoke-IsolatedDesktopWorker.ps1'),(Join-Path $profile 'Q014SolidWorks.vcxproj'),(Join-Path $profile 'Q014SolidWorks.sln'))
    $files += @(Get-ChildItem (Split-Path $testAssembly) -File | Where-Object Extension -in '.dll','.exe','.config' | Select-Object -ExpandProperty FullName)
    $files += Join-Path $EvidenceRoot 'Q014SolidWorks.cs'
    $files += Join-Path $EvidenceRoot 'Q020Startup.cs'
    if($managedEvidence) { $files += $managedEvidence.TrxPath,$managedEvidence.OriginalPlanPath }
    $files += $q030Evidence.TrxPath
    $files += @(Get-ChildItem (Join-Path $repository 'src'),(Join-Path $repository 'tests') -Recurse -File | Where-Object { $_.Extension -in '.cs','.csproj','.props','.targets' -and $_.FullName -notmatch '\\(bin|obj)\\' } | Select-Object -ExpandProperty FullName)
    Write-Json $planFile @{Scope="Q-020 Add101 / first SaveAs and Q-030 unsafe Open refusal / original-file native Edit Macro SOLIDWORKS $SolidWorksYear revision $ExpectedNativeRevision only; other versions separate";SolidWorksYear=$SolidWorksYear;ExpectedNativeRevision=$ExpectedNativeRevision;NativeFramePattern="^SOLIDWORKS.*$SolidWorksYear";SourceCommit=(& git -C $repository rev-parse HEAD);SourceStatus=@(& git -C $repository status --porcelain);
        Repository=$repository;EvidenceRoot=$EvidenceRoot;InstalledProduct=$installed;ProductSha256=$hash;
        ProductMvid=([Reflection.Assembly]::ReflectionOnlyLoadFrom($installed).ManifestModule.ModuleVersionId.ToString('D'));
        BuildOutputRoot=$BuildOutputRoot;ManagedEvidence=$managedEvidence;Q030ManagedEvidence=$q030Evidence;HelperAssembly=$helper;TestAssembly=$testAssembly;SolidWorksExecutable=$SolidWorksExecutable;VisualStudioExecutable=$VisualStudioExecutable;
        Solution=(Join-Path $profile 'Q014SolidWorks.sln');Scenarios=$scenarios;FrozenFiles=@($files | Select-Object -Unique | ForEach-Object {@{Path=$_;Sha256=(Get-FileHash -LiteralPath $_).Hash}});
        LaunchAuthorization="Maintainer explicitly authorized autonomous SOLIDWORKS $SolidWorksYear launch in this chat";CloseAuthorization='Maintainer authorized automatic normal closure of qualification instances';NoRetries=$true;NoDesktopSwitch=$true;NoForceTermination=$true;PreparedUtc=[DateTime]::UtcNow.ToString('o')}
    Get-Content $planFile -Raw -Encoding UTF8
    exit 0
}

$plan = Get-Content (Join-Path $PSScriptRoot 'q020-plan.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA' -or -not $env:VBAi_TEST_DESKTOP_NAME) { throw 'Only the reviewed private desktop STA worker may execute this campaign.' }
foreach ($file in $plan.FrozenFiles) {
    if ((Get-FileHash -LiteralPath $file.Path).Hash -cne $file.Sha256) { throw "Frozen file changed: $($file.Path)" }
}
$ledger = Join-Path $plan.EvidenceRoot 'campaign.json'
if (Test-Path -LiteralPath $ledger) { throw 'Campaign already claimed; no replay.' }
$records = @($plan.Scenarios | ForEach-Object { [pscustomobject]@{Id=$_.Id;Oracle=$_.Oracle;State='NOT_RUN';Error=$null} })
$script:blocked = $false; $script:wire = 0; $script:hostProcess = $null
function Record { Write-Json $ledger @{Scenarios=$records;NoRetries=$true;Desktop=$env:VBAi_TEST_DESKTOP_NAME;Utc=[DateTime]::UtcNow.ToString('o')} }
function Check($condition,[string]$message) { if (-not $condition) { throw $message } }
$desktopType=[Reflection.Assembly]::LoadFrom($plan.HelperAssembly).GetType('VBAi.Tests.Integration.IsolatedTestDesktop',$true)
$flags=[Reflection.BindingFlags]'Static,NonPublic'
function Desktop-Check { $desktopType.GetMethod('RequireCurrent',$flags).Invoke($null,@($env:VBAi_TEST_DESKTOP_NAME)) | Out-Null }
function Window-Check($window) {
    Desktop-Check
    $desktopType.GetMethod('RequireOfficeWindowInventory',$flags).Invoke($null,@([string]$env:VBAi_TEST_DESKTOP_NAME,[uint32]$window.Pid,[bool]$true,[IntPtr]::new([long]$window.Handle)))|Out-Null
}
function Stage([string]$id,[scriptblock]$body) {
    $record=@($records | Where-Object Id -eq $id)[0]
    if ($script:blocked) { return }
    $record.State='STARTED_ONCE'; Record
    try {
        Desktop-Check; & $body
        if($script:baselineSources) {
            foreach($name in $script:baselineSources.Keys) {
                Check ((Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$name}).Sha256 -ceq $script:baselineSources[$name]) 'Original disposable-macro source changed.'
            }
        }
        $record.State='PASS'; Record
    }
    catch { $record.State='FAILED_OR_UNCERTAIN';$record.Error=$_.Exception.ToString()+[Environment]::NewLine+$_.ScriptStackTrace;$script:blocked=$true;Record }
}
function Send-Bridge([hashtable]$request,[switch]$AllowRefusal) {
    Desktop-Check; $script:hostProcess.Refresh(); Check (-not $script:hostProcess.HasExited) 'Owned original host exited.'
    $stem=Join-Path $plan.EvidenceRoot ('wire-'+(++$script:wire).ToString('D4'))
    Write-Json ($stem+'-intent.json') @{Pid=$script:hostProcess.Id;Request=$request;Utc=[DateTime]::UtcNow.ToString('o')}
    $response=& (Join-Path $plan.Repository 'tools/Invoke-VBAi.ps1') -HostProcessId $script:hostProcess.Id -ResponseTimeoutSeconds 20 -RequestJson ($request|ConvertTo-Json -Compress -Depth 12) | ConvertFrom-Json
    Write-Json ($stem+'-response.json') $response
    if ($AllowRefusal) { return $response }
    Check ($response.Ok -eq $true) ($request.Command+': '+$response.Error)
    return $response.Data
}
function File-Dialog([int]$command,[string]$path,[string]$action) {
    $windows=@([Q014Native]::Windows($script:hostProcess.Id)); Write-Json (Join-Path $plan.EvidenceRoot ('dialog-'+$command+'-before-'+$script:wire+'.json')) $windows
    Check (@($windows | Where-Object {$_.Class -eq '#32770' -and $_.Visible}).Count -eq 0) 'Existing modal dialog prevents native command.'
    $root=@($windows | Where-Object {$_.Class -match '^Afx:' -and $_.Caption -match $plan.NativeFramePattern -and $_.Visible})
    Check ($root.Count -eq 1) 'Unique visible SOLIDWORKS frame required.'
    Window-Check $root[0]
    $intent=Join-Path $plan.EvidenceRoot ('native-command-'+$command+'-intent.json')
    Check (-not(Test-Path $intent)) 'Native command already claimed.'
    Write-Json $intent @{Command=$command;Path=$path;Action=$action;Window=$root[0]}
    $script:hostProcess.Refresh();Check (-not $script:hostProcess.HasExited) 'Original host exited before native file command.'
    $script:pendingFileCommand=[Q014SolidWorks]::BeginFileCommand($script:hostProcess.Id,$plan.ExpectedNativeRevision,$command)
    $deadline=[DateTime]::UtcNow.AddSeconds(20)
    do { $dialogs=@([Q014Native]::Windows($script:hostProcess.Id) | Where-Object {$_.Class -eq '#32770' -and $_.Visible});if($dialogs.Count -ne 0){break};Start-Sleep -Milliseconds 200 } while([DateTime]::UtcNow -lt $deadline)
    Check ($dialogs.Count -eq 1) 'The one native command did not expose a unique file dialog.'
    Window-Check $dialogs[0];$controls=[Q014Native]::FileControls($dialogs[0].Handle,$script:hostProcess.Id)
    $caption=[Q014Native]::Text($controls[1]);Check ($caption.Replace('&','') -in $(if($action -eq 'Save'){@('Save','Enregistrer')}else{@('Open','Ouvrir')})) 'Unexpected dialog action.'
    if($action -eq 'Save'){Check (-not(Test-Path $path)) 'Fresh output required.'}else{Check (Test-Path $path) 'Owned input missing.'}
    [Q014Native]::Filename($controls[0],$path)
    $again=[Q014Native]::FileControls($dialogs[0].Handle,$script:hostProcess.Id);Check (($again -join ',') -ceq ($controls -join ',')) 'File dialog controls changed.'
    Window-Check $dialogs[0];Write-Json ($intent.Replace('-intent','-button-intent')) @{Controls=$controls;Caption=$caption;Path=$path;Delivery='WM_SETTEXT/readback/BM_CLICK; no global input'}
    [Q014Native]::Click($controls[1])
    $deadline=[DateTime]::UtcNow.AddSeconds(10)
    while(-not $script:pendingFileCommand.Completed -and [DateTime]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 100}
    Write-Json ($intent.Replace('-intent','-terminal')) @{Completed=$script:pendingFileCommand.Completed;Returned=$script:pendingFileCommand.Returned;Error=$script:pendingFileCommand.Error}
    Check ($script:pendingFileCommand.Completed -and $script:pendingFileCommand.Returned -and -not $script:pendingFileCommand.Error) 'Native file command did not reach verified terminal success; no retry.'
}
Record
Stage 'managed' {
    if($plan.ManagedEvidence) {
        $e=$plan.ManagedEvidence
        Check ((Get-FileHash -LiteralPath $e.TrxPath).Hash -ceq $e.TrxSha256 -and (Get-FileHash -LiteralPath $plan.TestAssembly).Hash -ceq $e.TestAssemblySha256 -and (Get-FileHash -LiteralPath $plan.InstalledProduct).Hash -ceq $e.ProductSha256 -and (Get-FileHash -LiteralPath $plan.HelperAssembly).Hash -ceq $e.HelperSha256) 'Pinned managed evidence binary/TRX changed.'
        foreach($file in $e.SourceFiles) { Check ((Get-FileHash -LiteralPath $file.Path).Hash -ceq $file.Sha256) 'Managed evidence source changed.' }
        Check ((Read-Q020ManagedEvidence $e.TrxPath $plan.TestAssembly) -eq $e.Total) 'Managed evidence counter changed.'
        Write-Json (Join-Path $plan.EvidenceRoot 'managed-evidence-reused.json') @{State='MANAGED_EVIDENCE_REUSED';Evidence=$e;FreshTestExecution=$false;Utc=[DateTime]::UtcNow.ToString('o')}
        return
    }
    $results=Join-Path $plan.EvidenceRoot 'managed';[IO.Directory]::CreateDirectory($results)|Out-Null
    & dotnet vstest $plan.TestAssembly '/TestCaseFilter:TestCategory=Unit&(FullyQualifiedName~VbeProjectLifecycleTests|FullyQualifiedName~Standalone|FullyQualifiedName~VbeSolidWorksPersistenceTests|FullyQualifiedName~ListProjectsPreservesOriginalFileNameGetterDiagnostics|FullyQualifiedName~ListProjectsRetainsUnsavedProjectWhenFileNameIsUnavailable)' '/Logger:trx;LogFileName=managed.trx' "/ResultsDirectory:$results" *> (Join-Path $results 'test.log')
    Check ($LASTEXITCODE -eq 0) 'Scoped managed tests failed; native launch refused.'
    [xml]$trx=Get-Content (Join-Path $results 'managed.trx');$c=$trx.TestRun.ResultSummary.Counters
    Check ([int]$c.total -gt 0 -and [int]$c.total -eq [int]$c.passed) 'Managed tests missing/skipped/failed.'
}
Stage 'managed-q030' {
    $e=$plan.Q030ManagedEvidence
    Check ((Get-FileHash -LiteralPath $e.TrxPath).Hash -ceq $e.TrxSha256 -and (Get-FileHash -LiteralPath $plan.TestAssembly).Hash -ceq $e.TestAssemblySha256 -and (Get-FileHash -LiteralPath $plan.InstalledProduct).Hash -ceq $e.ProductSha256) 'Q030 pinned managed evidence changed.'
    Check ((Read-Q030ManagedEvidence $e.TrxPath $plan.TestAssembly) -eq $e.Total) 'Q030 managed evidence counters changed.'
    Write-Json (Join-Path $plan.EvidenceRoot 'q030-managed-evidence-reused.json') @{State='MANAGED_EVIDENCE_REUSED';Evidence=$e;FreshTestExecution=$false;Scope='Actual managed catalogue/approval/privacy/project-binding guards; native catalogue is not claimed'}
}
Add-Type -Path (Join-Path $PSScriptRoot 'Q014Native.cs')
Add-Type -Path (Join-Path $PSScriptRoot 'Q020Startup.cs') -ReferencedAssemblies UIAutomationClient,UIAutomationTypes,WindowsBase
$swInterop=Join-Path (Split-Path $plan.SolidWorksExecutable) 'api/redist/SolidWorks.Interop.sldworks.dll'
[Reflection.Assembly]::LoadFrom($swInterop)|Out-Null
Add-Type -Path (Join-Path $PSScriptRoot 'Q014SolidWorks.cs') -ReferencedAssemblies $swInterop
Stage 'launch' {
    Check (@(Get-Process SLDWORKS -ErrorAction SilentlyContinue).Count -eq 0) 'Existing SOLIDWORKS prevents another launch.'
    $launch=$desktopType.GetMethod('Launch',$flags)
    Write-Json (Join-Path $plan.EvidenceRoot 'vs-launch-intent.json') @{Executable=$plan.VisualStudioExecutable;Solution=$plan.Solution;Command='Debug.Start';Desktop=$env:VBAi_TEST_DESKTOP_NAME}
    $script:vsChild=$launch.Invoke($null,@([string]$plan.VisualStudioExecutable,[string[]]@($plan.Solution),[string]$plan.EvidenceRoot,[string]$env:VBAi_TEST_DESKTOP_NAME))
    $instanceFlags=[Reflection.BindingFlags]'Instance,NonPublic'
    $vsPid=$script:vsChild.GetType().GetProperty('ProcessId',$instanceFlags).GetValue($script:vsChild,$null)
    $vsThread=$script:vsChild.GetType().GetProperty('ThreadId',$instanceFlags).GetValue($script:vsChild,$null)
    $vsHandle=$script:vsChild.GetType().GetProperty('ProcessHandle',$instanceFlags).GetValue($script:vsChild,$null)
    $vsObserved=[Diagnostics.Process]::GetProcessById($vsPid)
    $script:vsBirth=$vsObserved.StartTime.ToUniversalTime().ToString('o');$vsObserved.Dispose()
    Write-Json (Join-Path $plan.EvidenceRoot 'vs-original-process.json') @{Pid=$vsPid;ThreadId=$vsThread;Handle=$vsHandle.ToInt64();StartUtc=$script:vsBirth;ImagePath=$plan.VisualStudioExecutable;ImageSha256=(Get-FileHash -LiteralPath $plan.VisualStudioExecutable).Hash;HandleSource='OriginalCreateProcessHandle';WorkerPid=$PID}
    $public=Join-Path (Split-Path $plan.VisualStudioExecutable) 'PublicAssemblies'
    $interop=Join-Path $public 'Microsoft.VisualStudio.Interop.dll';$envdte=Join-Path $public 'EnvDTE.dll'
    [Reflection.Assembly]::LoadFrom($interop)|Out-Null;[Reflection.Assembly]::LoadFrom($envdte)|Out-Null
    Add-Type -Path (Join-Path $PSScriptRoot 'Q014VisualStudio.cs') -ReferencedAssemblies $interop,$envdte
    $deadline=[DateTime]::UtcNow.AddSeconds(90)
    do {$dte=[Q014Native]::ExactRot('!VisualStudio.DTE.18.0:'+$vsPid);if($dte){break};Start-Sleep -Milliseconds 500}while([DateTime]::UtcNow -lt $deadline)
    Check ($null -ne $dte) 'Exact owned Visual Studio ROT unavailable.'
    $deadline=[DateTime]::UtcNow.AddSeconds(90);$ready=$false
    do {
        try {$ready=[Q014VisualStudio]::IsReady($dte,$plan.Solution)}
        catch {
            $cause=$_.Exception;while($cause.InnerException){$cause=$cause.InnerException}
            if($cause.HResult -ne -2147418111){throw} # RPC_E_CALL_REJECTED during read-only IDE loading.
        }
        if($ready){break};Start-Sleep -Milliseconds 500
    }while([DateTime]::UtcNow -lt $deadline)
    Check $ready 'Owned IDE project/document readiness unproved; no Debug.Start emitted.'
    $vsWindows=@([Q014Native]::Windows($vsPid));Check ($vsWindows.Count -gt 0) 'Owned IDE desktop unproved.'
    foreach($window in $vsWindows){Window-Check $window}
    $deadline=[DateTime]::UtcNow.AddSeconds(30);$beforeStart=$null
    do {
        try { $beforeStart=[Q014VisualStudio]::Read($dte);break }
        catch {
            $cause=$_.Exception;while($cause.InnerException){$cause=$cause.InnerException}
            if($cause.HResult -ne -2147418111){throw} # Read-only readiness observation; never repeat Debug.Start.
        }
        Start-Sleep -Milliseconds 200
    }while([DateTime]::UtcNow -lt $deadline)
    Check ($null -ne $beforeStart -and $beforeStart.Mode -eq 1 -and $beforeStart.Targets.Count -eq 0) 'Idle prelaunch debugger state unproved; no Debug.Start emitted.'
    Write-Json (Join-Path $plan.EvidenceRoot 'debug-start-intent.json') @{VsPid=$vsPid;Solution=$plan.Solution;Before=$beforeStart;Invocations=1}
    [Q014VisualStudio]::Start($dte,$plan.Solution)
    $deadline=[DateTime]::UtcNow.AddSeconds(120)
    do {$hosts=@(Get-Process SLDWORKS -ErrorAction SilentlyContinue);if($hosts.Count -gt 0){break};Start-Sleep -Milliseconds 500}while([DateTime]::UtcNow -lt $deadline)
    Check ($hosts.Count -eq 1) 'One Debug.Start did not expose a unique host within the launch bound; no retry.'
    $script:hostProcess=$hosts[0];$null=$script:hostProcess.Handle
    $deadline=[DateTime]::UtcNow.AddSeconds(15)
    do {$script:hostProcess.Refresh();$actualPath=$script:hostProcess.Path;if($actualPath){break};Start-Sleep -Milliseconds 200}while([DateTime]::UtcNow -lt $deadline)
    Check ($actualPath -ieq $plan.SolidWorksExecutable) 'Wrong or unavailable SOLIDWORKS executable.'
    $script:hostBirth=$script:hostProcess.StartTime.ToUniversalTime().ToString('o')
    Write-Json (Join-Path $plan.EvidenceRoot 'host-original-process.json') @{Pid=$script:hostProcess.Id;StartUtc=$script:hostBirth;ImagePath=$actualPath;ImageSha256=(Get-FileHash -LiteralPath $actualPath).Hash;OriginalHandle=$script:hostProcess.Handle.ToInt64();HandleSource='WorkerFirstQueryHandle';CreateProcessHandleProven=$false;WorkerPid=$PID;ParentVsPid=$vsPid;Desktop=$env:VBAi_TEST_DESKTOP_NAME;FrameVerified=$false;RevisionVerified=$false}
    $deadline=[DateTime]::UtcNow.AddSeconds(90)
    $toolbarAcknowledged=$false;$toolbarClosed=$false;$startupObservation=0;$toolbarModal=0L
    do {
        $desktopType.GetMethod('RequireOfficeWindowInventory',$flags).Invoke($null,@([string]$env:VBAi_TEST_DESKTOP_NAME,[uint32]$script:hostProcess.Id,[bool]$false,[IntPtr]::Zero))|Out-Null
        $windows=@([Q020Startup]::Snapshot($script:hostProcess.Id,$deadline))
        Write-Json (Join-Path $plan.EvidenceRoot ('startup-windows-'+(++$startupObservation).ToString('D3')+'.json')) @{Utc=[DateTime]::UtcNow.ToString('o');Windows=$windows;Deadline=$deadline}
        $button=[Q020Startup]::KnownToolbarButton($windows)
        if($button) {
            if($toolbarAcknowledged) { Check (-not $toolbarClosed -and $button.Root -eq $toolbarModal) 'Toolbar warning reappeared/changed after one acknowledgment; no retry.' }
            else {
                $toolbarModal=$button.Root;$toolbarAcknowledged=$true
                $privateGuard=[Action]{Desktop-Check;$script:hostProcess.Refresh();Check (-not $script:hostProcess.HasExited -and $script:hostProcess.StartTime.ToUniversalTime().ToString('o') -ceq $script:hostBirth -and $script:hostProcess.Path -ieq $plan.SolidWorksExecutable) 'Original host changed before toolbar acknowledgment.';$desktopType.GetMethod('RequireOfficeWindowInventory',$flags).Invoke($null,@([string]$env:VBAi_TEST_DESKTOP_NAME,[uint32]$script:hostProcess.Id,[bool]$true,[IntPtr]::new($toolbarModal)))|Out-Null}
                [Q020Startup]::Acknowledge($windows,(Join-Path $plan.EvidenceRoot 'toolbar-warning-intent.json'),$deadline,$privateGuard)
            }
        } elseif($toolbarAcknowledged) {
            Check (@($windows|Where-Object Handle -eq $toolbarModal).Count -eq 0) 'Original toolbar warning closure unproved.'
            $toolbarClosed=$true
        }
        $frames=@($windows|Where-Object {$_.Class -match '^Afx:' -and $_.Caption -match $plan.NativeFramePattern -and $_.Visible})
        if($frames.Count -gt 0 -and -not $button){break}
        Start-Sleep -Milliseconds 500
    }while([DateTime]::UtcNow -lt $deadline)
    Check ($frames.Count -eq 1) 'Actual SOLIDWORKS frame desktop unproved.';Window-Check $frames[0]
    $revision=[Q014SolidWorks]::Verify($script:hostProcess.Id);Check ($revision -ceq $plan.ExpectedNativeRevision) ("Expected native SOLIDWORKS revision "+$plan.ExpectedNativeRevision)
    Write-Json (Join-Path $plan.EvidenceRoot 'host-identity.json') @{Pid=$script:hostProcess.Id;Path=$script:hostProcess.Path;StartUtc=$script:hostProcess.StartTime.ToUniversalTime().ToString('o');OriginalHandle=$script:hostProcess.Handle.ToInt64();Revision=$revision;Windows=$windows;Desktop=$env:VBAi_TEST_DESKTOP_NAME}
    if($toolbarAcknowledged){Write-Json (Join-Path $plan.EvidenceRoot 'toolbar-warning-terminal.json') @{AcknowledgedOnce=$true;OriginalModalClosed=$true;NativeFrameVerified=$true;Utc=[DateTime]::UtcNow.ToString('o')}}
}
$script:bootstrap=Join-Path $plan.EvidenceRoot ('Q020Bootstrap'+$plan.SolidWorksYear+'.swp')
$script:fixture=Join-Path $plan.EvidenceRoot ('Q020Standalone'+$plan.SolidWorksYear+'.swp')
Stage 'bootstrap' { File-Dialog 573 $script:bootstrap 'Save'; Check ((Test-Path $script:bootstrap) -and (Get-Item $script:bootstrap).Length -gt 0) 'Owned bootstrap absent.' }
Stage 'load' {
    $status=Send-Bridge @{Command='status'}
    Check ($status.Connected -and $status.HostProcessId -eq $script:hostProcess.Id -and $status.ProcessBitness -eq 64 -and $status.AssemblyModuleVersionId -ceq $plan.ProductMvid -and $status.AssemblyPath -ieq $plan.InstalledProduct) 'Loaded candidate identity differs.'
    Check ((Get-FileHash $status.AssemblyPath).Hash -ceq $plan.ProductSha256) 'Loaded bytes differ.'
    $addins=Send-Bridge @{Command='list_addins'}
    Check (@($addins.AddIns|Where-Object {$_.Properties.ProgId -eq 'VBAi.AddIn' -and $_.Properties.Connect}).Count -eq 1) 'Addin connection unproved.'
    Write-Json (Join-Path $plan.EvidenceRoot 'loaded-candidate.json') $status
}
Stage 'add101' {
    $before=Send-Bridge @{Command='project_collection_state'}
    $added=Send-Bridge @{Command='create_standalone_project';ExpectedProjectVersion=$before.Version}
    Write-Json (Join-Path $plan.EvidenceRoot 'add101-result.json') @{Before=$before;Result=$added}
    Check ($added.Verified -and $added.MutationInvoked -and -not $added.Uncertain) 'Add101 did not verify; never replay.'
    $extra=@($added.CollectionState.Projects|Where-Object {$_.Identity -notin @($before.Projects.Identity)})
    Check ($extra.Count -eq 1 -and $extra[0].Type -eq 101 -and $extra[0].Mode -eq 2 -and $extra[0].Protection -eq 0 -and -not $extra[0].Saved -and -not $extra[0].Path) 'Native Type101 unsaved state differs.'
    $script:selector=$extra[0].Name
    $script:preexisting=$before
}
Stage 'unsaved' {
    $metadata=Send-Bridge @{Command='project_properties';Project=$script:selector}
    $filename=@($metadata.Properties|Where-Object Name -eq 'FileName')
    Check ($filename.Count -eq 1) 'FileName descriptor absent.'
    $persistence=Send-Bridge @{Command='project_persistence_status';Project=$script:selector}
    Check (-not $persistence.ProjectSaved -and -not $persistence.HostHasPath -and -not $persistence.FileExists) 'Unsaved normalized persistence differs.'
    $rows=@(Send-Bridge @{Command='list_projects'})
    $raw=@($rows|Where-Object {$_.Name -ceq $script:selector})
    Check ($raw.Count -eq 1 -and $raw[0].PSObject.Properties.Name -contains 'FileNameErrorHResult' -and $raw[0].PSObject.Properties.Name -contains 'FileNameErrorType') 'Raw native getter diagnostic unavailable; do not substitute simulated evidence.'
    if($null -ne $raw[0].FileNameErrorHResult){
        $hr=[BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$raw[0].FileNameErrorHResult),0)
        Check ($hr -in @(0x800A004CL,0x80070003L) -and $raw[0].FileNameErrorType -and $null -eq $raw[0].FileName) 'Unsaved raw getter exception differs from supported path-not-found cases.'
        $getterState='NATIVE_GETTER_EXCEPTION_OBSERVED'
    }else{
        Check ($null -eq $raw[0].FileNameErrorType -and [string]::IsNullOrWhiteSpace($raw[0].FileName)) 'Successful unsaved getter returned an unexpected path/error type.'
        $getterState='NATIVE_GETTER_SUCCESS_OBSERVED'
    }
    Write-Json (Join-Path $plan.EvidenceRoot 'unsaved-getter-observation.json') @{NativeFirstGetter=$raw[0];RawGetterState=$getterState;Descriptor=$filename[0];Persistence=$persistence;Scope='Diagnostics describe the first list_projects FileName getter. HostPath independently rereads it; this is not a one-getter count claim.'}
}
Stage 'first-save-as' {
    Check (-not(Test-Path $script:fixture)) 'Fresh original output required.'
    $metadata=Send-Bridge @{Command='project_properties';Project=$script:selector}
    $saved=Send-Bridge @{Command='save_host_document_as';Project=$script:selector;Path=$script:fixture;ExpectedProjectVersion=$metadata.Version}
    Check ($saved.SaveAsInvoked -and $saved.SaveInvoked -and $saved.ProjectSaved -and $saved.HostPath -ieq $script:fixture -and $saved.Bytes -gt 0 -and -not $saved.Uncertain) 'First SaveAs state unproved; no replay.'
    $script:selector=$script:fixture
    $collection=Send-Bridge @{Command='project_collection_state'}
    Check (@($collection.Projects|Where-Object {$_.Path -ieq $script:fixture -and $_.Type -eq 101 -and $_.Saved}).Count -eq 1) 'First saved project Type101 identity differs.'
    Write-Json (Join-Path $plan.EvidenceRoot 'first-save-as.json') @{Result=$saved;Collection=$collection;Sha256=(Read-OwnedFileHash $script:fixture)}
    $script:baseline=@(Send-Bridge @{Command='list_modules';Project=$script:fixture});$script:baselineSources=@{}
    foreach($m in $script:baseline){$script:baselineSources[$m.Name]=(Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$m.Name}).Sha256}
}
function Snapshot {
    $state=Send-Bridge @{Command='project_properties';Project=$script:fixture}
    $code=@{}
    foreach($m in $state.Components){$code[$m.Name]=Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$m.Name}}
    return @{Components=$state.Components;References=$state.References;Code=$code;Form=(Send-Bridge @{Command='form_state';Project=$script:fixture;Form='Q020Form'})}
}
Stage 'components' {
    foreach($entry in @(@{Name='Q020Module';Command='create_module';Text="Option Explicit`r`nPublic Function SyntheticValue() As Long`r`n    SyntheticValue = 42`r`nEnd Function"},@{Name='Q020Class';Command='create_class';Text="Option Explicit`r`nPublic Function SyntheticText() As String`r`n    SyntheticText = `"Q020 owned`"`r`nEnd Function"},@{Name='Q020Form';Command='create_form';Text="Option Explicit`r`nPrivate Sub Q020Marker()`r`n    Dim owned As Long`r`n    owned = 20`r`nEnd Sub"})) {
        Check ($entry.Name -notin @($script:baseline.Name)) 'Synthetic name collision.'
        $req=@{Command=$entry.Command;Project=$script:fixture;ExpectedMode=2}
        if($entry.Command -eq 'create_form'){$req.Form=$entry.Name}else{$req.Module=$entry.Name}
        Send-Bridge $req|Out-Null
        $before=Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$entry.Name}
        $component=Send-Bridge @{Command='component_properties';Project=$script:fixture;Module=$entry.Name}
        Send-Bridge @{Command='replace_lines';Project=$script:fixture;Module=$entry.Name;ExpectedSha256=$before.Sha256;StartLine=1;Count=$component.CodeLines;Text=$entry.Text}|Out-Null
        Check ((Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$entry.Name}).Code.Trim() -ceq $entry.Text.Trim()) 'Synthetic source readback differs.'
    }
    $form=Send-Bridge @{Command='form_state';Project=$script:fixture;Form='Q020Form'}
    Send-Bridge @{Command='add_form_control';Project=$script:fixture;Form='Q020Form';ExpectedFormVersion=$form.Version;Control='Q020Label';ControlType='Forms.Label.1';Left=12;Top=12;Width=160;Height=24;Caption='Q020 owned standalone'}|Out-Null
    $script:expected=Snapshot
    Check (@($script:expected.Form.Controls|Where-Object {$_.Name -eq 'Q020Label' -and $_.Caption -ceq 'Q020 owned standalone' -and $_.Left -eq 12 -and $_.Top -eq 12 -and $_.Width -eq 160 -and $_.Height -eq 24}).Count -eq 1) 'Native Label state differs.'
    Write-Json (Join-Path $plan.EvidenceRoot 'pending-snapshot.json') $script:expected
}
Stage 'compile' {
    $selected=Send-Bridge @{Command='select_code';Project=$script:fixture;Module='Q020Module';ExpectedSha256=$script:expected.Code['Q020Module'].Sha256;StartLine=2;ExpectedMode=2}
    Check ($selected.Project -ieq $script:fixture -and $selected.Module -ceq 'Q020Module' -and $selected.Line -eq 2 -and $selected.Mode -eq 2) 'Owned project code selection not verified before compile.'
    Write-Json (Join-Path $plan.EvidenceRoot 'compile-owned-selection.json') $selected
    $compiled=Send-Bridge @{Command='compile_project';Project=$script:fixture;ExpectedMode=2}
    Check ($compiled.Compiled -and -not (Send-Bridge @{Command='debug_dialog'}).Visible -and (Send-Bridge @{Command='debug_state';Project=$script:fixture}).Mode -eq 2) 'Compile/design state unproved.'
}
Stage 'save' {
    $metadata=Send-Bridge @{Command='project_properties';Project=$script:fixture}
    $saved=Send-Bridge @{Command='save_host_document';Project=$script:fixture;ExpectedHostPath=$script:fixture;ExpectedProjectVersion=$metadata.Version}
    Check ($saved.SaveInvoked -and -not $saved.SaveAsInvoked -and $saved.ProjectSaved -and -not $saved.Uncertain -and $saved.HostPath -ieq $script:fixture) 'One product Save unverified.'
    $script:savedHash=Read-OwnedFileHash $script:fixture
    Write-Json (Join-Path $plan.EvidenceRoot 'saved-file.json') @{Result=$saved;Sha256=$script:savedHash;Path=$script:fixture}
}
Stage 'q030-refuse-unsafe-open' {
    $beforeCollection=Send-Bridge @{Command='project_collection_state'}
    $beforeSnapshot=Snapshot
    $beforeBytes=Read-OwnedFileHash $script:fixture
    $beforeIdentity=@{Pid=$script:hostProcess.Id;StartUtc=$script:hostProcess.StartTime.ToUniversalTime().ToString('o');Handle=$script:hostProcess.Handle.ToInt64()}
    $response=Send-Bridge @{Command='open_standalone_project';Path=$script:fixture;ExpectedProjectVersion=$beforeCollection.Version} -AllowRefusal
    $expectedError='Opening a standalone SWP through VBProjects.Open is disabled in SOLIDWORKS after an observed host termination. Use SOLIDWORKS Tools > Macro > Edit instead.'
    Check ($response.Ok -eq $false -and $null -eq $response.Data -and $response.Error -ceq $expectedError) 'The exact SOLIDWORKS unsafe-Open refusal was not observed.'
    $afterCollection=Send-Bridge @{Command='project_collection_state'}
    $afterSnapshot=Snapshot
    Check ($afterCollection.Version -ceq $beforeCollection.Version -and ($afterCollection.Projects|ConvertTo-Json -Depth 30 -Compress) -ceq ($beforeCollection.Projects|ConvertTo-Json -Depth 30 -Compress)) 'Refusal changed the native project collection.'
    Check (($afterSnapshot.Components|ConvertTo-Json -Depth 30 -Compress) -ceq ($beforeSnapshot.Components|ConvertTo-Json -Depth 30 -Compress)) 'Refusal changed components.'
    Check (($afterSnapshot.References|ConvertTo-Json -Depth 30 -Compress) -ceq ($beforeSnapshot.References|ConvertTo-Json -Depth 30 -Compress)) 'Refusal changed references.'
    foreach($name in $beforeSnapshot.Code.Keys){Check ($afterSnapshot.Code[$name].Sha256 -ceq $beforeSnapshot.Code[$name].Sha256 -and $afterSnapshot.Code[$name].Code -ceq $beforeSnapshot.Code[$name].Code) 'Refusal changed full module/class/form source.'}
    Check ($afterSnapshot.Form.Version -ceq $beforeSnapshot.Form.Version -and ($afterSnapshot.Form.Controls|ConvertTo-Json -Depth 30 -Compress) -ceq ($beforeSnapshot.Form.Controls|ConvertTo-Json -Depth 30 -Compress)) 'Refusal changed form state.'
    Check ((Read-OwnedFileHash $script:fixture) -ceq $beforeBytes) 'Refusal changed original file bytes.'
    $status=Send-Bridge @{Command='status'}
    Check ($status.Connected -and $status.HostProcessId -eq $beforeIdentity.Pid -and $status.AssemblyModuleVersionId -ceq $plan.ProductMvid -and $status.AssemblyPath -ieq $plan.InstalledProduct -and (Get-FileHash $status.AssemblyPath).Hash -ceq $plan.ProductSha256) 'Refusal host or loaded candidate changed.'
    Check (-not $script:hostProcess.HasExited -and $script:hostProcess.StartTime.ToUniversalTime().ToString('o') -ceq $beforeIdentity.StartUtc -and $script:hostProcess.Handle.ToInt64() -eq $beforeIdentity.Handle) 'Original owned process identity changed.'
    Write-Json (Join-Path $plan.EvidenceRoot 'q030-native-refusal.json') @{Response=$response;BeforeCollection=$beforeCollection;AfterCollection=$afterCollection;BeforeSnapshot=$beforeSnapshot;AfterSnapshot=$afterSnapshot;OriginalFileSha256=$beforeBytes;HostIdentity=$beforeIdentity;LoadedCandidate=$status;Scope='Native direct product refusal and unchanged owned state. Zero collection access is a separate managed assertion; native API call count is not instrumented. Catalogue approval/privacy guards have separate managed TRX evidence. Product standalone Open remains unavailable.'}
}
Stage 'close-standalone' {
    $metadata=Send-Bridge @{Command='project_properties';Project=$script:fixture}
    $closed=Send-Bridge @{Command='close_standalone_project';Project=$script:fixture;ExpectedHostPath=$script:fixture;ExpectedProjectVersion=$metadata.Version}
    Check ($closed.Verified -and $closed.MutationInvoked -and -not $closed.Uncertain) 'Close standalone unverified; no replay.'
    Check (@($closed.CollectionState.Projects|Where-Object {$_.Path -ieq $script:fixture}).Count -eq 0) 'Original standalone remains loaded.'
    Check ((Get-FileHash $script:fixture).Hash -ceq $script:savedHash) 'Close changed original bytes.'
    $script:baselineSources=$null
    Write-Json (Join-Path $plan.EvidenceRoot 'original-closed-file.json') @{Result=$closed;Path=$script:fixture;Sha256=$script:savedHash;ReloadInput='SameOriginalFile'}
}
Stage 'reload' {
    File-Dialog 84 $script:fixture 'Open'
    $collection=Send-Bridge @{Command='project_collection_state'}
    $reloaded=@($collection.Projects|Where-Object {$_.Path -ieq $script:fixture -and $_.Type -in @(100,101) -and $_.Mode -eq 2 -and $_.Saved})
    Check ($reloaded.Count -eq 1) 'Native original reload path/design/saved identity differs.'
    Write-Json (Join-Path $plan.EvidenceRoot 'native-reload-hosting.json') @{NativeApi='SOLIDWORKS EditMacro84';Path=$script:fixture;ObservedType=$reloaded[0].Type;Scope='Add101 / first SaveAs / Remove require standalone Type101; native Edit Macro may expose the saved file as a host Type100 project.'}
    $actual=Snapshot
    Check (($actual.Components|ConvertTo-Json -Compress) -ceq ($script:expected.Components|ConvertTo-Json -Compress)) 'Reload component inventory differs.'
    Check (($actual.References|ConvertTo-Json -Compress) -ceq ($script:expected.References|ConvertTo-Json -Compress)) 'Reload reference manifest differs.'
    foreach($name in $script:expected.Code.Keys){Check ($actual.Code[$name].Sha256 -ceq $script:expected.Code[$name].Sha256 -and $actual.Code[$name].Code -ceq $script:expected.Code[$name].Code) 'Full source reload differs.'}
    Check ($actual.Form.Version -ceq $script:expected.Form.Version) 'Reload full designer state differs.'
    Check ((Read-OwnedFileHash $script:fixture) -ceq $script:savedHash) 'Native Edit changed original bytes.'
    Write-Json (Join-Path $plan.EvidenceRoot 'original-reload-snapshot.json') $actual
}
Stage 'cleanup' {
    $exit=[Q014SolidWorks]::BeginNormalExit($script:hostProcess.Id,$plan.ExpectedNativeRevision)
    $bound=[DateTime]::UtcNow.AddSeconds(30)
    while(-not $exit.Completed -and [DateTime]::UtcNow -lt $bound){Start-Sleep -Milliseconds 100}
    Check ($exit.Completed -and $exit.Returned -and -not $exit.Error) 'Normal owned host ExitApp did not return.'
    Check ($script:hostProcess.WaitForExit(30000) -and $script:hostProcess.ExitCode -eq 0) 'Original owned host normal exit unproved.'
    Write-Json (Join-Path $plan.EvidenceRoot 'host-normal-exit.json') @{Pid=$script:hostProcess.Id;StartUtc=$script:hostProcess.StartTime.ToUniversalTime().ToString('o');ImagePath=$plan.SolidWorksExecutable;OriginalHandle=$script:hostProcess.Handle.ToInt64();HandleSource='WorkerFirstQueryHandle';CreateProcessHandleProven=$false;ExitObserved=$true;ExitCode=$script:hostProcess.ExitCode;NormalExit=$true;ForcedTermination=$false;CandidateMvid=$plan.ProductMvid;CandidateSha256=$plan.ProductSha256}
    [Q014VisualStudio]::Quit($dte,$plan.Solution)
    $wait=$script:vsChild.GetType().GetMethod('Wait',$instanceFlags).Invoke($script:vsChild,@([int]30000))
    Check $wait 'Owned VS original exit unproved.'
    $vsExit=$script:vsChild.GetType().GetMethod('ExitCode',$instanceFlags).Invoke($script:vsChild,@())
    Check ($vsExit -eq 0) 'Owned VS did not exit normally.'
    Write-Json (Join-Path $plan.EvidenceRoot 'vs-normal-exit.json') @{Pid=$vsPid;StartUtc=$script:vsBirth;ImagePath=$plan.VisualStudioExecutable;OriginalHandle=$vsHandle.ToInt64();HandleSource='OriginalCreateProcessHandle';ExitObserved=$true;ExitCode=$vsExit;ForcedTermination=$false}
}
if($script:blocked -and $script:hostProcess -and -not $script:hostProcess.HasExited){
    Write-Json (Join-Path $plan.EvidenceRoot 'retained-owner.json') @{Pid=$script:hostProcess.Id;Desktop=$env:VBAi_TEST_DESKTOP_NAME;Reason='Failed or uncertain stage; original owner retained without retry or force.'}
    while(-not $script:hostProcess.HasExited){Start-Sleep -Milliseconds 1000;$script:hostProcess.Refresh()}
}
if($script:blocked -and $script:hostProcess -and $script:hostProcess.HasExited) {
    Write-Json (Join-Path $plan.EvidenceRoot 'failed-original-host-exit.json') @{State='FAILED_ORIGINAL_HOST_EXIT_OBSERVED';Pid=$script:hostProcess.Id;StartUtc=$script:hostBirth;ImagePath=$plan.SolidWorksExecutable;OriginalHandle=$script:hostProcess.Handle.ToInt64();HandleSource='WorkerFirstQueryHandle';CreateProcessHandleProven=$false;ExitObserved=$true;ExitCode=$script:hostProcess.ExitCode;NormalExitProven=$false;NativeQualificationPromoted=$false;Utc=[DateTime]::UtcNow.ToString('o')}
}
if($script:blocked -and $script:vsChild) {
    # Retain the original CreateProcess handle until actual IDE exit, even if SW exited first.
    while(-not $script:vsChild.GetType().GetMethod('Wait',$instanceFlags).Invoke($script:vsChild,@([int]1000))) { }
    $observed=$script:vsChild.GetType().GetMethod('Wait',$instanceFlags).Invoke($script:vsChild,@([int]0))
    if($observed) { $code=$script:vsChild.GetType().GetMethod('ExitCode',$instanceFlags).Invoke($script:vsChild,@());Write-Json (Join-Path $plan.EvidenceRoot 'failed-original-vs-exit.json') @{State='FAILED_ORIGINAL_IDE_EXIT_OBSERVED';Pid=$vsPid;StartUtc=$script:vsBirth;ImagePath=$plan.VisualStudioExecutable;OriginalHandle=$vsHandle.ToInt64();HandleSource='OriginalCreateProcessHandle';ExitObserved=$true;ExitCode=$code;NormalExitProven=$false;NativeQualificationPromoted=$false;Utc=[DateTime]::UtcNow.ToString('o')} }
}
if($script:blocked){exit 1}else{exit 0}

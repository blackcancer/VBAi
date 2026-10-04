#requires -Version 5.1
param(
    [switch]$Prepare,
    [switch]$ContinueAfterLaunchReceiptError,
    [string]$EvidenceRoot,
    [string]$InstalledDirectory,
    [string]$RetainedManagedRoot,
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
if ($Prepare) {
    if (-not [IO.Path]::IsPathRooted($EvidenceRoot) -or -not (Test-Path -LiteralPath $EvidenceRoot -PathType Container)) { throw 'An existing absolute build/evidence root is required.' }
    $planFile = Join-Path $EvidenceRoot 'q014-plan.json'
    if (Test-Path -LiteralPath $planFile) { throw 'Preparation already exists; do not overwrite a campaign.' }
    $repository = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    $installed = Join-Path $InstalledDirectory 'VBAi.dll'
    $testAssembly = Join-Path $EvidenceRoot 'build/VBAi.Q014.Tests/Debug/net48/VBAi.Tests.dll'
    $helper = Join-Path $EvidenceRoot 'build/VBAi.Desktop.Helper/Debug/net48/VBAi.Desktop.Helper.exe'
    foreach ($file in @($installed,$testAssembly,$helper,$SolidWorksExecutable,$VisualStudioExecutable)) {
        if (-not [IO.Path]::IsPathRooted($file) -or -not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Required absolute file missing: $file" }
    }
    $nativeFileVersion=(Get-Item -LiteralPath $SolidWorksExecutable).VersionInfo.FileVersion
    if($nativeFileVersion.Split('.')[0] -ne $(if($SolidWorksYear -eq 2019){'27'}else{'33'}) -or
        $nativeFileVersion.Split('.')[0] -ne $ExpectedNativeRevision.Split('.')[0]) { throw 'Selected SOLIDWORKS year/revision and executable differ.' }
    $hash = (Get-FileHash -LiteralPath $installed).Hash
    if ($RetainedManagedRoot) {
        $previousPlan=Get-Content (Join-Path $RetainedManagedRoot 'q014-plan.json') -Raw | ConvertFrom-Json
        if ($previousPlan.ProductSha256 -cne $hash -or
            (Get-FileHash $previousPlan.TestAssembly).Hash -cne (Get-FileHash $testAssembly).Hash) { throw 'Retained managed candidate differs.' }
        foreach ($old in $previousPlan.FrozenFiles | Where-Object {$_.Path.StartsWith((Split-Path $previousPlan.TestAssembly),[StringComparison]::OrdinalIgnoreCase)}) {
            $current=Join-Path (Split-Path $testAssembly) ([IO.Path]::GetFileName($old.Path))
            if ((Get-FileHash $old.Path).Hash -cne $old.Sha256 -or (Get-FileHash $current).Hash -cne $old.Sha256) { throw 'Retained managed dependencies differ.' }
        }
        $retainedTrx=Join-Path $RetainedManagedRoot 'managed/managed.trx'
        [xml]$retained=Get-Content $retainedTrx
        $counts=$retained.TestRun.ResultSummary.Counters
        if ([int]$counts.total -le 0 -or [int]$counts.total -ne [int]$counts.passed) { throw 'Retained managed tests not all passed.' }
    }
    if ((Get-FileHash (Join-Path (Split-Path $testAssembly) 'VBAi.dll')).Hash -cne $hash) { throw 'The harness must reference the exact installed candidate.' }
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
EndGlobal
'@
    Set-Content (Join-Path $profile 'Q014SolidWorks.sln') $solution -Encoding UTF8
    Copy-Item -LiteralPath $PSCommandPath -Destination (Join-Path $EvidenceRoot 'Invoke-FrozenQ014.ps1')
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Q014Native.cs') -Destination $EvidenceRoot
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Q014VisualStudio.cs') -Destination $EvidenceRoot
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Q014SolidWorks.cs') -Destination $EvidenceRoot
    $scenarios = @(
        @{Id='managed'; Oracle='Exact scoped TRX; persistence identity/revision/owner guards and one-mutation uncertainty; desktop argument guards'}
        @{Id='launch'; Oracle="One owned VS utility Debug.Start; exact $SolidWorksYear executable, PID, start time, native UI desktop and exact PID ROT; revision $ExpectedNativeRevision"}
        @{Id='new-macro'; Oracle='One native New Macro, one addressed Save; fresh owned SWP and Type100 project; no existing macro execution'}
        @{Id='load'; Oracle='Installed path, SHA256, MVID, 64-bit bridge and connected VBAi.AddIn in exact native host'}
        @{Id='module-class'; Oracle='Unique synthetic module/class code and recognized procedure; unchanged original inventory/source'}
        @{Id='form'; Oracle='Synthetic UserForm/Label caption, geometry and code with independent native readback'}
        @{Id='stale-guards'; Oracle='Stale source/project/form requests refused with unchanged native source/designer'}
        @{Id='export-import'; Oracle='Retained BAS/CLS/FRM/FRX, guarded remove/import once each, full source/designer equality'}
        @{Id='compile'; Oracle='Native compilation success, no diagnostic and design mode'}
        @{Id='debug'; Oracle='One breakpoint/run/step/continue; observed break lines, scalar value and unique synthetic output; no shortcut'}
        @{Id='save'; Oracle='One product Save; Verified=true, Uncertain=false, Saved=true; unchanged source and retained SWP hash'}
        @{Id='reopen'; Oracle='Separate retained byte copy opened by native Edit Macro; source/class/designer equality and disk hash unchanged; no execution to unload'}
        @{Id='ui'; Oracle='Actual Monaco module/class rendering, form designer, resize/restore, close/reopen and stale-project warning absence; reviewed real captures'}
        @{Id='assistant'; Oracle='Synthetic local assistant tool dispatch, permission/revision refusal, cancellation and recovery; no credentials or paid calls'}
        @{Id='cleanup'; Oracle='Backups retained, exact owned host normal exit observed on original handle; no force termination; independent debugger state'}
    )
    $files = @($installed,$testAssembly,$helper,(Join-Path $EvidenceRoot 'Invoke-FrozenQ014.ps1'),(Join-Path $EvidenceRoot 'Q014Native.cs'),(Join-Path $EvidenceRoot 'Q014VisualStudio.cs'),(Join-Path $repository 'tools/Invoke-VBAi.ps1'),(Join-Path $repository 'tools/tests/Invoke-IsolatedDesktopWorker.ps1'),(Join-Path $profile 'Q014SolidWorks.vcxproj'),(Join-Path $profile 'Q014SolidWorks.sln'))
    $files += @(Get-ChildItem (Split-Path $testAssembly) -File | Where-Object Extension -in '.dll','.exe','.config' | Select-Object -ExpandProperty FullName)
    $files += Join-Path $EvidenceRoot 'Q014SolidWorks.cs'
    Write-Json $planFile @{Scope="Q-014 SOLIDWORKS $SolidWorksYear revision $ExpectedNativeRevision only; other versions separate";SolidWorksYear=$SolidWorksYear;ExpectedNativeRevision=$ExpectedNativeRevision;NativeFramePattern="^SOLIDWORKS.*$SolidWorksYear";SourceCommit=(& git -C $repository rev-parse HEAD);SourceStatus=@(& git -C $repository status --porcelain);
        Repository=$repository;EvidenceRoot=$EvidenceRoot;InstalledProduct=$installed;ProductSha256=$hash;
        ProductMvid=([Reflection.Assembly]::ReflectionOnlyLoadFrom($installed).ManifestModule.ModuleVersionId.ToString('D'));
        HelperAssembly=$helper;TestAssembly=$testAssembly;SolidWorksExecutable=$SolidWorksExecutable;VisualStudioExecutable=$VisualStudioExecutable;
        Solution=(Join-Path $profile 'Q014SolidWorks.sln');Scenarios=$scenarios;FrozenFiles=@($files | Select-Object -Unique | ForEach-Object {@{Path=$_;Sha256=(Get-FileHash -LiteralPath $_).Hash}});
        RetainedManagedTrx=$(if($RetainedManagedRoot){$retainedTrx}else{$null});RetainedManagedSha256=$(if($RetainedManagedRoot){(Get-FileHash $retainedTrx).Hash}else{$null});
        LaunchAuthorization="Maintainer explicitly authorized autonomous SOLIDWORKS $SolidWorksYear launch in this chat";CloseAuthorization='Maintainer authorized automatic normal closure of qualification instances';NoRetries=$true;NoDesktopSwitch=$true;NoForceTermination=$true;PreparedUtc=[DateTime]::UtcNow.ToString('o')}
    Get-Content $planFile -Raw -Encoding UTF8
    exit 0
}

$plan = Get-Content (Join-Path $PSScriptRoot 'q014-plan.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA' -or -not $env:VBAi_TEST_DESKTOP_NAME) { throw 'Only the reviewed private desktop STA worker may execute this campaign.' }
foreach ($file in $plan.FrozenFiles) {
    $actualHash=(Get-FileHash -LiteralPath $file.Path).Hash
    if ($actualHash -cne $file.Sha256) {
        $accepted=$false
        if($ContinueAfterLaunchReceiptError -and $file.Path -ceq $plan.Solution) {
            $review=Get-Content (Join-Path $plan.EvidenceRoot 'reviewed-solution-rewrite.json') -Raw -Encoding UTF8|ConvertFrom-Json
            $accepted=$review.Path -ceq $file.Path -and $review.OriginalSha256 -ceq $file.Sha256 -and $review.ReviewedSha256 -ceq $actualHash
        }
        if(-not $accepted){throw "Frozen file changed: $($file.Path)"}
    }
}
$ledger = Join-Path $plan.EvidenceRoot 'campaign.json'
if ($ContinueAfterLaunchReceiptError) {
    $previous=Get-Content $ledger -Raw -Encoding UTF8|ConvertFrom-Json
    if (@($previous.Scenarios|Where-Object {$_.Id -eq 'managed' -and $_.State -eq 'PASS'}).Count -ne 1 -or
        @($previous.Scenarios|Where-Object {$_.Id -notin @('managed','launch') -and $_.State -ne 'NOT_RUN'}).Count -ne 0 -or
        @(Get-ChildItem $plan.EvidenceRoot -Filter 'wire-*-intent.json').Count -ne 0) { throw 'Continuation requires completed managed proof and no native fixture action.' }
    $ledger=Join-Path $plan.EvidenceRoot 'campaign-continuation.json'
}
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
    $actual=$desktopType.GetMethod('DesktopName',$flags).Invoke($null,@([uint32]$window.Thread))
    Check ($actual -ceq $env:VBAi_TEST_DESKTOP_NAME) 'Native UI thread is not on the inactive desktop.'
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
if (-not $ContinueAfterLaunchReceiptError) { Stage 'managed' {
    if ($plan.RetainedManagedTrx) {
        Check ((Get-FileHash $plan.RetainedManagedTrx).Hash -ceq $plan.RetainedManagedSha256) 'Retained TRX changed.'
        Write-Json (Join-Path $plan.EvidenceRoot 'retained-managed.json') @{Trx=$plan.RetainedManagedTrx;Sha256=$plan.RetainedManagedSha256;Candidate=$plan.ProductSha256;ExecutedAgain=$false}
        return
    }
    $results=Join-Path $plan.EvidenceRoot 'managed';[IO.Directory]::CreateDirectory($results)|Out-Null
    & dotnet vstest $plan.TestAssembly '/TestCaseFilter:TestCategory=Unit' '/Logger:trx;LogFileName=managed.trx' "/ResultsDirectory:$results" *> (Join-Path $results 'test.log')
    Check ($LASTEXITCODE -eq 0) 'Scoped managed tests failed; native launch refused.'
    [xml]$trx=Get-Content (Join-Path $results 'managed.trx');$c=$trx.TestRun.ResultSummary.Counters
    Check ([int]$c.total -gt 0 -and [int]$c.total -eq [int]$c.passed) 'Managed tests missing/skipped/failed.'
} }
Add-Type -Path (Join-Path $PSScriptRoot 'Q014Native.cs')
$swInterop=Join-Path (Split-Path $plan.SolidWorksExecutable) 'api/redist/SolidWorks.Interop.sldworks.dll'
[Reflection.Assembly]::LoadFrom($swInterop)|Out-Null
Add-Type -Path (Join-Path $PSScriptRoot 'Q014SolidWorks.cs') -ReferencedAssemblies $swInterop
if (-not $ContinueAfterLaunchReceiptError) { Stage 'launch' {
    Check (@(Get-Process SLDWORKS -ErrorAction SilentlyContinue).Count -eq 0) 'Existing SOLIDWORKS prevents another launch.'
    $launch=$desktopType.GetMethod('Launch',$flags)
    Write-Json (Join-Path $plan.EvidenceRoot 'vs-launch-intent.json') @{Executable=$plan.VisualStudioExecutable;Solution=$plan.Solution;Command='Debug.Start';Desktop=$env:VBAi_TEST_DESKTOP_NAME}
    $script:vsChild=$launch.Invoke($null,@([string]$plan.VisualStudioExecutable,[string[]]@($plan.Solution),[string]$plan.EvidenceRoot,[string]$env:VBAi_TEST_DESKTOP_NAME))
    $instanceFlags=[Reflection.BindingFlags]'Instance,NonPublic'
    $vsPid=$script:vsChild.GetType().GetProperty('ProcessId',$instanceFlags).GetValue($script:vsChild,$null)
    $vsThread=$script:vsChild.GetType().GetProperty('ThreadId',$instanceFlags).GetValue($script:vsChild,$null)
    $vsHandle=$script:vsChild.GetType().GetProperty('ProcessHandle',$instanceFlags).GetValue($script:vsChild,$null)
    Write-Json (Join-Path $plan.EvidenceRoot 'vs-original-process.json') @{Pid=$vsPid;ThreadId=$vsThread;Handle=$vsHandle.ToInt64()}
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
    $deadline=[DateTime]::UtcNow.AddSeconds(90)
    do {$windows=@([Q014Native]::Windows($script:hostProcess.Id));$frames=@($windows|Where-Object {$_.Class -match '^Afx:' -and $_.Caption -match $plan.NativeFramePattern -and $_.Visible});if($frames.Count -gt 0){break};Start-Sleep -Milliseconds 500}while([DateTime]::UtcNow -lt $deadline)
    Check ($frames.Count -eq 1) 'Actual SOLIDWORKS frame desktop unproved.';Window-Check $frames[0]
    $revision=[Q014SolidWorks]::Verify($script:hostProcess.Id);Check ($revision -ceq $plan.ExpectedNativeRevision) ("Expected native SOLIDWORKS revision "+$plan.ExpectedNativeRevision)
    Write-Json (Join-Path $plan.EvidenceRoot 'host-identity.json') @{Pid=$script:hostProcess.Id;Path=$script:hostProcess.Path;StartUtc=$script:hostProcess.StartTime.ToUniversalTime().ToString('o');OriginalHandle=$script:hostProcess.Handle.ToInt64();Revision=$revision;Windows=$windows;Desktop=$env:VBAi_TEST_DESKTOP_NAME}
} } else {
    $recovered=Get-Content (Join-Path $plan.EvidenceRoot 'host-identity-recovered.json') -Raw -Encoding UTF8|ConvertFrom-Json
    Check ($recovered.State -eq 'READY' -and $recovered.NativeDebugStartAttempts -eq 1 -and $recovered.Desktop -ceq $env:VBAi_TEST_DESKTOP_NAME) 'Verified recovered launch receipt required.'
    $script:hostProcess=Get-Process -Id $recovered.Pid;$null=$script:hostProcess.Handle
    Check ($script:hostProcess.Path -ieq $plan.SolidWorksExecutable -and $script:hostProcess.StartTime.ToUniversalTime().ToString('o') -ceq $recovered.StartUtc) 'Recovered host identity changed.'
    Check ([Q014SolidWorks]::Verify($recovered.Pid) -ceq $plan.ExpectedNativeRevision) 'Recovered exact PID ROT differs.'
    $frames=@([Q014Native]::Windows($recovered.Pid)|Where-Object {$_.Class -match '^Afx:' -and $_.Caption -match $plan.NativeFramePattern -and $_.Visible});Check ($frames.Count -eq 1) 'Recovered native frame ambiguous.';Window-Check $frames[0]
    @($records|Where-Object Id -eq 'managed')[0].State='PASS_RETAINED_TRX'
    @($records|Where-Object Id -eq 'launch')[0].State='PASS_WITH_RETAINED_INITIAL_HARNESS_FAILURE'
    Record
}
$script:fixture=Join-Path $plan.EvidenceRoot ("Q014Owned"+$plan.SolidWorksYear+".swp")
Stage 'new-macro' { File-Dialog 573 $script:fixture 'Save';Start-Sleep -Milliseconds 800;Check ((Test-Path $script:fixture) -and (Get-Item $script:fixture).Length -gt 0) 'Native-created SWP absent.' }
Stage 'load' {
    $status=Send-Bridge @{Command='status'};Write-Json (Join-Path $plan.EvidenceRoot 'loaded-candidate.json') $status
    Check ($status.HostProcessId -eq $script:hostProcess.Id -and $status.ProcessBitness -eq 64 -and $status.AssemblyModuleVersionId -ceq $plan.ProductMvid -and $status.AssemblyPath -ieq $plan.InstalledProduct) 'Loaded candidate identity differs.'
    Check ((Get-FileHash $status.AssemblyPath).Hash -ceq $plan.ProductSha256) 'Loaded candidate disk bytes differ.'
    $addins=Send-Bridge @{Command='list_addins'}
    Check (@($addins.AddIns | Where-Object {$_.Properties.ProgId -eq 'VBAi.AddIn' -and $_.Properties.Connect}).Count -eq 1) 'VBAi.AddIn connection not verified.'
    $projects=@(Send-Bridge @{Command='list_projects'});Check (@($projects|Where-Object {$_.FileName -ieq $script:fixture -and $_.Mode -eq 2}).Count -eq 1) 'Exact disposable project unavailable in design mode.'
    $script:baseline=@(Send-Bridge @{Command='list_modules';Project=$script:fixture})
    $script:baselineSources=@{}
    foreach($module in $script:baseline){$script:baselineSources[$module.Name]=(Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$module.Name}).Sha256}
    Write-Json (Join-Path $plan.EvidenceRoot 'baseline.json') $script:baseline
    Write-Json (Join-Path $plan.EvidenceRoot 'baseline-source-hashes.json') $script:baselineSources
}
Stage 'module-class' {
    $script:module='Q014Module';$script:class='Q014Class';$script:form='Q014Form'
    foreach($entry in @(@{Name=$script:module;Command='create_module';Text="Option Explicit`r`nPublic Sub Q014Debug()`r`n    Dim value As Long`r`n    value = 41`r`n    value = value + 1`r`n    Open `"$(Join-Path $plan.EvidenceRoot 'debug-marker.txt')`" For Output As #1`r`n    Print #1, CStr(value)`r`n    Close #1`r`nEnd Sub"},@{Name=$script:class;Command='create_class';Text="Option Explicit`r`nPublic Function Value() As Long`r`n    Value = 42`r`nEnd Function"})) {
        Send-Bridge @{Command=$entry.Command;Project=$script:fixture;Module=$entry.Name;ExpectedMode=2}|Out-Null
        $before=Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$entry.Name}
        $component=Send-Bridge @{Command='component_properties';Project=$script:fixture;Module=$entry.Name}
        Send-Bridge @{Command='replace_lines';Project=$script:fixture;Module=$entry.Name;ExpectedSha256=$before.Sha256;StartLine=1;Count=$component.CodeLines;Text=$entry.Text}|Out-Null
        $after=Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$entry.Name}
        Check ($after.Code.Trim() -ceq $entry.Text.Trim()) 'Synthetic code readback differs.'
    }
    $script:moduleSource=Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$script:module}
    $script:classSource=Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$script:class}
    $procedures=Send-Bridge @{Command='list_procedures';Project=$script:fixture;Module=$script:module};Check (@($procedures.Procedures|Where-Object Name -eq 'Q014Debug').Count -eq 1) 'Native procedure unrecognized.'
}
Stage 'form' {
    Send-Bridge @{Command='create_form';Project=$script:fixture;Form=$script:form}|Out-Null
    $state=Send-Bridge @{Command='form_state';Project=$script:fixture;Form=$script:form}
    Send-Bridge @{Command='add_form_control';Project=$script:fixture;Form=$script:form;ExpectedFormVersion=$state.Version;Control='Q014Label';ControlType='Forms.Label.1';Left=12;Top=12;Width=160;Height=24;Caption=("Q014 "+$plan.SolidWorksYear+" synthetic")}|Out-Null
    $script:formState=Send-Bridge @{Command='form_state';Project=$script:fixture;Form=$script:form}
    Check (@($script:formState.Controls|Where-Object {$_.Name -eq 'Q014Label' -and ($_.Caption -ceq ("Q014 "+$plan.SolidWorksYear+" synthetic"))}).Count -eq 1) 'Native Label differs.'
}
Stage 'stale-guards' {
    $refusal=Send-Bridge @{Command='replace_lines';Project=$script:fixture;Module=$script:module;ExpectedSha256=('0'*64);StartLine=1;Count=1;Text="' must never apply"} -AllowRefusal
    Check (-not $refusal.Ok -and $refusal.Error -match 'changed|SHA|hash|modifi') 'Stale source request not explicitly refused.'
    Check ((Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$script:module}).Sha256 -ceq $script:moduleSource.Sha256) 'Refused request changed code.'
    $refusal=Send-Bridge @{Command='remove_component';Project=$script:fixture;Module=$script:class;ExpectedProjectVersion='stale';ExpectedComponentVersion='stale'} -AllowRefusal
    Check (-not $refusal.Ok) 'Stale component removal accepted.'
    Check ((Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$script:class}).Sha256 -ceq $script:classSource.Sha256) 'Refused removal changed class.'
    $refusal=Send-Bridge @{Command='set_form_property';Project=$script:fixture;Form=$script:form;ExpectedFormVersion='stale';Property='Caption';Value='must never apply'} -AllowRefusal
    Check (-not $refusal.Ok) 'Stale form request accepted.'
    Check ((Send-Bridge @{Command='form_state';Project=$script:fixture;Form=$script:form}).Version -ceq $script:formState.Version) 'Refused request changed designer.'
}
Stage 'export-import' {
    foreach($entry in @(@{Name=$script:module;Extension='.bas'},@{Name=$script:class;Extension='.cls'},@{Name=$script:form;Extension='.frm'})) {
        $path=Join-Path $plan.EvidenceRoot ($entry.Name+$entry.Extension);$before=Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$entry.Name}
        $component=Send-Bridge @{Command='component_properties';Project=$script:fixture;Module=$entry.Name}
        Send-Bridge @{Command='export_component';Project=$script:fixture;Module=$entry.Name;Path=$path;ExpectedComponentVersion=$component.Version}|Out-Null
        Check ((Test-Path $path) -and (Get-Item $path).Length -gt 0) 'Native export absent.'
        $project=Send-Bridge @{Command='project_properties';Project=$script:fixture}
        $removed=Send-Bridge @{Command='remove_component';Project=$script:fixture;Module=$entry.Name;ExpectedProjectVersion=$project.Version;ExpectedComponentVersion=$component.Version}
        $imported=Send-Bridge @{Command='import_component';Project=$script:fixture;Path=$path;ExpectedProjectVersion=$removed.Version}
        Check ($imported.Applied -and $imported.ImportedName -ceq $entry.Name) 'Import result differs.'
        Check ((Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$entry.Name}).Sha256 -ceq $before.Sha256) 'Imported source differs.'
        if($entry.Extension -eq '.frm') { Check ((Send-Bridge @{Command='form_state';Project=$script:fixture;Form=$entry.Name}).Version -ceq $script:formState.Version) 'Imported form differs.' }
    }
}
Stage 'compile' {
    Send-Bridge @{Command='select_code';Project=$script:fixture;Module=$script:module;ExpectedSha256=$script:moduleSource.Sha256;StartLine=2}|Out-Null
    $compiled=Send-Bridge @{Command='compile_project';Project=$script:fixture;ExpectedMode=2};Check ($compiled.Compiled -eq $true) 'Compile not verified.'
    Check (-not (Send-Bridge @{Command='debug_dialog'}).Visible) 'Native compile diagnostic visible.'
    Check ((Send-Bridge @{Command='debug_state';Project=$script:fixture}).Mode -eq 2) 'Compile left design mode.'
}
Stage 'debug' {
    function Debug-Action([int]$id,[string]$action,[int]$line,[int]$mode) {
        $commands=@(Send-Bridge @{Command='list_commands';Query=''})
        $command=@($commands|Where-Object {$_.Id -eq $id -and $_.Enabled}|Group-Object Caption|ForEach-Object {$_.Group[0]});Check ($command.Count -eq 1) 'Unique enabled native debugger command identity required.'
        Send-Bridge @{Command='invoke_debug';Project=$script:fixture;Module=$script:module;ExpectedSha256=$script:moduleSource.Sha256;StartLine=$line;ExpectedMode=$mode;Action=$action;ControlId=$id;ControlCaption=$command[0].Caption}|Out-Null
    }
    function Debug-Observe([int]$mode,[int]$line) {
        $deadline=[DateTime]::UtcNow.AddSeconds(8)
        do {$state=Send-Bridge @{Command='debug_state';Project=$script:fixture};if($state.Mode -eq $mode -and ($line -eq 0 -or $state.Selection.StartLine -eq $line)){return};Start-Sleep -Milliseconds 100}while([DateTime]::UtcNow -lt $deadline)
        throw 'Native debugger state/line unproved; no replay.'
    }
    Debug-Action 51 'toggle_breakpoint' 5 2;Debug-Action 186 'run' 2 2;Debug-Observe 1 5
    Debug-Action 194 'step_over' 5 1;Debug-Observe 1 6
    Debug-Action 186 'continue' 6 1;Debug-Observe 2 0
    Check ((Get-Content (Join-Path $plan.EvidenceRoot 'debug-marker.txt') -Raw).Trim() -ceq '42') 'Synthetic result differs.'
    Debug-Action 51 'toggle_breakpoint' 5 2
}
Stage 'save' {
    $project=Send-Bridge @{Command='project_properties';Project=$script:fixture}
    $save=Send-Bridge @{Command='save_host_document';Project=$script:fixture;ExpectedHostPath=$script:fixture;ExpectedProjectVersion=$project.Version}
    Check ($save.Verified -eq $true -and $save.Uncertain -eq $false -and $save.SaveInvoked -eq $true -and $save.ProjectSaved -eq $true) 'Product save uncertain or failed.'
    Check ((Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$script:module}).Sha256 -ceq $script:moduleSource.Sha256) 'Save changed module.'
    Check ((Send-Bridge @{Command='read_module';Project=$script:fixture;Module=$script:class}).Sha256 -ceq $script:classSource.Sha256) 'Save changed class.'
    Check ((Send-Bridge @{Command='form_state';Project=$script:fixture;Form=$script:form}).Version -ceq $script:formState.Version) 'Save changed designer.'
    $script:savedHash=Read-OwnedFileHash $script:fixture
    Write-Json (Join-Path $plan.EvidenceRoot 'saved-file.json') @{Path=$script:fixture;Sha256=$script:savedHash;Result=$save}
}
Stage 'reopen' {
    $script:reopened=Join-Path $plan.EvidenceRoot 'Q014SavedCopy.swp'
    $input=[IO.File]::Open($script:fixture,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    try {$output=[IO.File]::Open($script:reopened,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None);try {$input.CopyTo($output)}finally {$output.Dispose()}}finally {$input.Dispose()}
    Check ((Get-FileHash $script:reopened).Hash -ceq $script:savedHash) 'Saved copy bytes differ.'
    File-Dialog 84 $script:reopened 'Open';Start-Sleep -Milliseconds 500
    Check ((Send-Bridge @{Command='read_module';Project=$script:reopened;Module=$script:module}).Sha256 -ceq $script:moduleSource.Sha256) 'Reopened module differs.'
    Check ((Send-Bridge @{Command='read_module';Project=$script:reopened;Module=$script:class}).Sha256 -ceq $script:classSource.Sha256) 'Reopened class differs.'
    Check ((Send-Bridge @{Command='form_state';Project=$script:reopened;Form=$script:form}).Version -ceq $script:formState.Version) 'Reopened form differs.'
    Check ((Read-OwnedFileHash $script:reopened) -ceq $script:savedHash) 'Native Edit Macro changed file bytes.'
}
# UI and assistant require actual observed embedded controls, rendering and approved
# local provider capability. A bridge response alone cannot satisfy these oracles.
if(-not $script:blocked) {
    foreach($id in @('ui','assistant')) {$r=@($records|Where-Object Id -eq $id)[0];$r.State='NOT_RUN';$r.Error='Separate native embedded observation/dispatch contract required; no simulated acceptance.'}
    Record
}
# The maintainer authorized normal automatic closure for this qualification.
# Preserve ownership while the separate embedded UI/assistant observers complete;
# their recovery close records must retain backups and independent exit evidence.
$cleanup=@($records|Where-Object Id -eq 'cleanup')[0]
$cleanup.Error='Normal automatic closure authorized; retain ownership until embedded UI qualification and source preservation finish.'
Record
if($script:hostProcess -and -not $script:hostProcess.HasExited) {
    Write-Json (Join-Path $plan.EvidenceRoot 'retained-owner.json') @{Pid=$script:hostProcess.Id;VsPid=$(if($ContinueAfterLaunchReceiptError){$recovered.VsPid}else{$vsPid});Desktop=$env:VBAi_TEST_DESKTOP_NAME;Reason='Retain original ownership; native outcomes and cleanup reported separately';NoRetries=$true}
    while(-not $script:hostProcess.HasExited) {Start-Sleep -Milliseconds 1000;$script:hostProcess.Refresh()}
}
exit 1

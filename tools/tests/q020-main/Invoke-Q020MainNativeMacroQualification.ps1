#requires -Version 5.1
param(
    [switch]$Prepare,[switch]$MainDesktopAuthorized,
    [string]$EvidenceRoot,
    [string]$RepositoryRoot,
    [string]$OwnerSid,
    [int]$ProtectedVisualStudioPid = 0,
    [string]$BuildOutputRoot,
    [string]$CandidatePath,
    [string]$ProfileSolution,
    [string[]]$ManagedEvidenceTrxPaths,
    [ValidateSet(2019,2025)][int]$SolidWorksYear = 2019,
    [string]$ExpectedNativeRevision = '27.5.0',
    [string]$SolidWorksExecutable = 'D:\Program Files\SOLIDWORKS 2019 Corp\SOLIDWORKS\SLDWORKS.exe',
    [string]$VisualStudioExecutable = 'C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\devenv.exe'
)
$ErrorActionPreference = 'Stop'
$env:PSModulePath=(Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules')+';'+$env:PSModulePath
Import-Module Microsoft.PowerShell.Utility
. (Join-Path $PSScriptRoot 'FormOracle.ps1')
function Write-Json($Path,$Value) {
    $s=[IO.File]::Open($Path,[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::Read)
    try{$b=[Text.UTF8Encoding]::new($false).GetBytes((ConvertTo-Json -InputObject $Value -Depth 100));$s.Write($b,0,$b.Length);$s.Flush($true)}finally{$s.Dispose()}
}
function Write-Claim($Path,$Value) {
    $s=[IO.File]::Open($Path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
    try{$b=[Text.UTF8Encoding]::new($false).GetBytes((ConvertTo-Json -InputObject $Value -Depth 100));$s.Write($b,0,$b.Length);$s.Flush($true)}finally{$s.Dispose()}
}
function Check($Condition,[string]$Message) { if(-not $Condition) { throw $Message } }
function Hash([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }
function Read-Managed([string]$Path,[string]$Assembly) {
    [xml]$trx=Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $c=$trx.TestRun.ResultSummary.Counters
    Check ([int]$c.total -gt 0 -and [int]$c.total -eq [int]$c.executed -and [int]$c.total -eq [int]$c.passed) 'Pinned managed gate has missing/failed/skipped tests.'
    foreach($m in $trx.TestRun.TestDefinitions.UnitTest.TestMethod) {
        $ownedScratchContract=$m.className -ceq 'VBAi.Tests.GitScratchDirectoryTests' -and $m.name -ceq 'StandaloneOutputUsesAnOwnedTemporaryGuidDirectory'
        Check ($m.codeBase -ieq $Assembly -and ($m.className -match '^VBAi\.Tests\.Unit\.' -or $ownedScratchContract)) 'Managed evidence is not from the selected test assembly or declared scoped contracts.'
    }
    return @{Path=$Path;Sha256=(Hash $Path);Total=[int]$c.total;Methods=@($trx.TestRun.TestDefinitions.UnitTest.TestMethod|ForEach-Object {$_.className+'.'+$_.name})}
}
if($Prepare) {
    Check ($MainDesktopAuthorized) 'Explicit selected main-desktop preparation authorization required.'
    Check ([IO.Path]::IsPathRooted($EvidenceRoot) -and [IO.Directory]::Exists($EvidenceRoot)) 'Existing absolute fresh evidence directory required.'
    $planPath=Join-Path $EvidenceRoot 'native-macro-plan.json'
    Check (-not [IO.File]::Exists($planPath)) 'Campaign preparation already exists; never overwrite.'
    if([string]::IsNullOrWhiteSpace($RepositoryRoot)){$RepositoryRoot=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path}
    Check ([IO.Path]::IsPathRooted($RepositoryRoot) -and [IO.File]::Exists((Join-Path $RepositoryRoot 'VBAi.sln'))) 'Explicit checked-out repository root with VBAi.sln required.'
    $repository=(Resolve-Path -LiteralPath $RepositoryRoot).Path
    if([string]::IsNullOrWhiteSpace($OwnerSid)){$OwnerSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value}
    $parsedSid=[Security.Principal.SecurityIdentifier]::new($OwnerSid)
    Check ($parsedSid.Value -ceq [Security.Principal.WindowsIdentity]::GetCurrent().User.Value -and $ProtectedVisualStudioPid -gt 0) 'Prepared owner SID must be the current user; protected IDE PID must be positive (the configured gate contract).'
    $tests=Join-Path $BuildOutputRoot 'VBAi.Tests/Debug/net48/VBAi.Tests.dll'
    $helper=Join-Path $BuildOutputRoot 'VBAi.Desktop.Helper/Debug/net48/VBAi.Desktop.Helper.exe'
    $base=Join-Path $repository 'tools/tests/Invoke-Q020Qualification.ps1'
    foreach($p in @($CandidatePath,$tests,$helper,$ProfileSolution,$base,$SolidWorksExecutable,$VisualStudioExecutable)) { Check ([IO.Path]::IsPathRooted($p) -and [IO.File]::Exists($p)) ('Required pinned file absent: '+$p) }
    Check ((Hash (Join-Path (Split-Path $tests) 'VBAi.dll')) -ceq (Hash $CandidatePath)) 'Product and test-copy candidate bytes differ.'
    Check (@($ManagedEvidenceTrxPaths).Count -gt 0) 'Actual focused managed evidence required; this harness never reruns tests.'
    $managed=@($ManagedEvidenceTrxPaths|ForEach-Object {Read-Managed $_ $tests})
    $methods=@($managed|ForEach-Object {$_.Methods})
    foreach($pattern in @('VbeSolidWorksMacroCreationTests','SolidWorksMacroPublication','SolidWorksMacro','SolidWorksOpenRefuses')) { Check (@($methods|Where-Object {$_ -match $pattern}).Count -gt 0) ('Missing current managed contract coverage: '+$pattern) }
    $major=$(if($SolidWorksYear -eq 2019){'27'}else{'33'})
    Check ((Get-Item -LiteralPath $SolidWorksExecutable).VersionInfo.FileVersion.Split('.')[0] -ceq $major -and $ExpectedNativeRevision.Split('.')[0] -ceq $major) 'Selected native version differs.'
    Copy-Item -LiteralPath $PSCommandPath -Destination (Join-Path $EvidenceRoot 'Invoke-FrozenNativeMacro.ps1')
    foreach($name in @('Q014Native.cs','Q014SolidWorks.cs','Q014VisualStudio.cs')) { Copy-Item -LiteralPath (Join-Path $repository ('tools/tests/'+$name)) -Destination (Join-Path $EvidenceRoot $name) }
    foreach($name in @('Q020Startup.cs','Q020NativeTeardown.cs','Q020NativeCapture.cs','Q020MainDesktopGate.cs','OwnedLayout.cs','FormOracle.ps1')) { Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $EvidenceRoot $name) }
    # A real owned bitmap exercises binary FRX transport. It is never displayed/executed as a macro.
    Add-Type -AssemblyName System.Drawing
    $bitmap=[Drawing.Bitmap]::new(8,8)
    try { for($x=0;$x -lt 8;$x++){for($y=0;$y -lt 8;$y++){$bitmap.SetPixel($x,$y,[Drawing.Color]::FromArgb(255,($x*30),($y*30),120))}};$bitmap.Save((Join-Path $EvidenceRoot 'owned-picture.bmp'),[Drawing.Imaging.ImageFormat]::Bmp) }
    finally {$bitmap.Dispose()}
    $scenarios=@(
        @{Id='managed';Oracle='Reuse pinned current focused TRX and binary/source manifest; no test replay'},
        @{Id='launch-1';Oracle='One selected visible main-desktop SOLIDWORKS CreateProcessW; original birth/image/handle and bounded exact warning acknowledgment; no IDE launched'},
        @{Id='bootstrap-1';Oracle='Read-only exact owned frame/dialog readiness up to30s before one bootstrap573 to expose VBE; persistent visible32770 fails untouched; strict FileDialog precondition remains'},
        @{Id='connected-1';Oracle='Synchronous exact loaded candidate and live AddIn connection observations before product scenarios; no autostart or Connect-setter acceptance'},
        @{Id='negative';Oracle='One unsupported source publication prewrite refusal; source/collection unchanged, destination absent'},
        @{Id='create';Oracle='A: one create_solidworks_macro fresh .swp -> unique Type100/ThisLibrary; edit module/class/Q020CreatedForm/picture, compile once, nativeSave once; one visible designer PNG pending manual review'},
        @{Id='publish';Oracle='B: unsaved101 populated and compiled draft with Q020PublishedForm -> explicit new100 mapping; unchanged source full snapshot/version; code/hiddenattributes/designer/picture/references. Publication visual UI is assessed only after fresh reopen with source101 absent; no B-save screenshot'},
        @{Id='q030';Oracle='Known unsafe VBProjects.Open refusal with full state and disk bytes unchanged'},
        @{Id='normal-close-1';Oracle='Normal original SW exit; user IDE preserved; after source-preservation proof only exact named dirty-source predeclared Yes6/No7/Cancel2 or native French Yes6/No7/Help9 prompt permits one No, no save; independent handle exit receipts'},
        @{Id='closed-bytes';Oracle='Archive actual files after normal original exit; record saved/closed byte hashes separately'},
        @{Id='launch-2';Oracle='One fresh selected host; previous owned originals exit proven'},
        @{Id='bootstrap-2';Oracle='Read-only exact fresh owned frame/dialog readiness up to30s before one separate bootstrap573; persistent visible32770 fails untouched; strict FileDialog precondition remains'},
        @{Id='connected-2';Oracle='Fresh synchronous exact loaded candidate and live AddIn connection observations; no Connect setter or autostart claim'},
        @{Id='reload-create';Oracle='NativeEdit84 A same original path once; full code/attributes/FRX semantic designer/picture/refs/General and closed-file hash; Q020CreatedForm designer PNG pending manual review'},
        @{Id='reload-publish';Oracle='NativeEdit84 B same original path once; one explicit designer materialization before component inspection; full persisted readback; source101 absent in fresh host; Q020PublishedForm designer PNG pending manual review; Type100 Saved diagnostic'},
        @{Id='normal-close-2';Oracle='Normal fresh SW exit; original retained creation handle verified; user IDE preserved'}
    )
    $frozen=@($CandidatePath,$tests,$helper,$base,$ProfileSolution,(Join-Path (Split-Path $ProfileSolution) 'Q014SolidWorks.vcxproj'),(Join-Path (Split-Path $ProfileSolution) 'Q014SolidWorks.vcxproj.user'),$SolidWorksExecutable,$VisualStudioExecutable)
    $frozen+=@(Get-ChildItem $EvidenceRoot -File|Select-Object -ExpandProperty FullName)
    $frozen+=@(Get-ChildItem (Split-Path $tests) -File|Where-Object Extension -in '.dll','.exe','.config'|Select-Object -ExpandProperty FullName)
    $frozen+=@(Get-ChildItem (Join-Path $repository 'src'),(Join-Path $repository 'tests') -Recurse -File|Where-Object {$_.Extension -in '.cs','.csproj','.props','.targets' -and $_.FullName -notmatch '\\(bin|obj)\\'}|Select-Object -ExpandProperty FullName)
    $frozen+=@(Join-Path $repository 'tools/Invoke-VBAi.ps1')
    $frozen+=@($ManagedEvidenceTrxPaths)
    Write-Json $planPath @{Schema='Q020NativeMacroMainOrderedV1';MainDesktopAuthorized=$true;LayoutHelper=(Join-Path $EvidenceRoot 'OwnedLayout.cs');CaptureLayoutScope='Native Toolbox fresh after last designer focus; one Hide0/show8 sameHWND; exact root placement and toolbox restoration; no General while hidden';DesktopScope='MainDefaultExplicitUserAuthorization';Authorization='Explicit maintainer authorization supplied through MainDesktopAuthorized for this prepared campaign';SourceCommit=(& git -c ('safe.directory='+$repository.Replace('\','/')) -C $repository rev-parse HEAD);SourceStatus=@(& git -c ('safe.directory='+$repository.Replace('\','/')) -C $repository status --porcelain);Repository=$repository;OwnerSid=$OwnerSid;ProtectedVisualStudioPid=$ProtectedVisualStudioPid;EvidenceRoot=$EvidenceRoot;BaseHarness=$base;BaseHarnessSha256=(Hash $base);InstalledProduct=$CandidatePath;ProductSha256=(Hash $CandidatePath);ProductMvid=[Reflection.Assembly]::ReflectionOnlyLoadFrom($CandidatePath).ManifestModule.ModuleVersionId.ToString('D');TestAssembly=$tests;HelperAssembly=$helper;BuildOutputRoot=$BuildOutputRoot;ManagedEvidence=$managed;SolidWorksYear=$SolidWorksYear;ExpectedNativeRevision=$ExpectedNativeRevision;NativeFramePattern="^SOLIDWORKS.*$SolidWorksYear";SolidWorksExecutable=$SolidWorksExecutable;VisualStudioExecutable=$VisualStudioExecutable;Solution=$ProfileSolution;Picture=(Join-Path $EvidenceRoot 'owned-picture.bmp');Scenarios=$scenarios;FrozenFiles=@($frozen|Select-Object -Unique|ForEach-Object {@{Path=$_;Sha256=(Hash $_)}});NoRetries=$true;NoDesktopSwitch=$true;NoForceTermination=$true;LegacyAdd101SaveAsAcceptance=$false;ConnectionPrerequisite='Synchronous exact-owner status and live AddIn GUID/Connect receipts in each generation before product actions; no Connect setter or autostart qualification';ClosedBytesOracle='Capture both saved-before-exit and closed-after-exit hashes; do not erase an observed difference. Reload must preserve closed input bytes and exact complete persisted contents.';PreparedUtc=[DateTime]::UtcNow.ToString('o')}
    Get-Content -LiteralPath $planPath -Raw -Encoding UTF8
    exit 0
}

$plan=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'native-macro-plan.json') -Raw -Encoding UTF8|ConvertFrom-Json
Check ($plan.MainDesktopAuthorized) 'Explicit frozen main authorization missing.'
Check ([Threading.Thread]::CurrentThread.ApartmentState -eq 'STA' -and $plan.DesktopScope -ceq 'MainDefaultExplicitUserAuthorization') 'Explicitly authorized main-desktop STA worker required.'
$env:VBAi_TEST_DESKTOP_NAME='Default'
foreach($file in $plan.FrozenFiles) {Check ((Hash $file.Path) -ceq $file.Sha256) ('Frozen file changed: '+$file.Path)}
$claim=Join-Path $plan.EvidenceRoot 'execution-claim.json'
$stream=[IO.File]::Open($claim,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
try {$bytes=[Text.Encoding]::UTF8.GetBytes((@{Pid=$PID;Utc=[DateTime]::UtcNow.ToString('o');PlanSha256=(Hash (Join-Path $PSScriptRoot 'native-macro-plan.json'))}|ConvertTo-Json));$stream.Write($bytes,0,$bytes.Length);$stream.Flush($true)}finally{$stream.Dispose()}
$campaignRoot=$plan.EvidenceRoot;$ledger=Join-Path $campaignRoot 'campaign.json';$script:blocked=$false;$script:wire=0
# The registration utility deliberately enables StrictMode. Initialize retained
# ownership before the first launch; direct host original handles remain authoritative.
$script:hostProcess=$null;$script:swChild=$null;$script:hostOriginalHandle=[IntPtr]::Zero
$script:hostBirth=$null
$script:disposableSourceName=$null;$script:disposableSourceVersion=$null
$script:instanceFlags=[Reflection.BindingFlags]'Instance,NonPublic'
$script:qualificationScriptRoot=$PSScriptRoot
$records=@($plan.Scenarios|ForEach-Object {[pscustomobject]@{Id=$_.Id;Oracle=$_.Oracle;State='NOT_RUN';Error=$null}})
function Record {Write-Json $ledger @{Scenarios=$records;NoRetries=$true;Desktop=$env:VBAi_TEST_DESKTOP_NAME;Utc=[DateTime]::UtcNow.ToString('o')}}
function Stage([string]$id,[scriptblock]$body) {
    if($script:blocked){return};$r=@($records|Where-Object Id -ceq $id)[0];$r.State='STARTED_ONCE';Record
    try{Desktop-Check;& $body;$r.State='PASS';Record}catch{$r.State='FAILED_OR_UNCERTAIN';$r.Error=$_.Exception.ToString()+[Environment]::NewLine+$_.ScriptStackTrace;$script:blocked=$true;Record}
}
# Import only reviewed definitions, never execute the legacy campaign/Prepare or its Add101 SaveAs path.
$tokens=$null;$parseErrors=$null;$baseAst=[Management.Automation.Language.Parser]::ParseFile($plan.BaseHarness,[ref]$tokens,[ref]$parseErrors)
Check ($parseErrors.Count -eq 0 -and (Hash $plan.BaseHarness) -ceq $plan.BaseHarnessSha256) 'Legacy reusable definitions drifted.'
foreach($name in @('Read-OwnedFileHash','Desktop-Check','Window-Check','Send-Bridge','File-Dialog')) {
    $definition=@($baseAst.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq $name},$true));Check ($definition.Count -eq 1) 'Exact unique reusable function required.'
    $text=$definition[0].Extent.Text
    $text=$text.Replace('$PSScriptRoot','$script:qualificationScriptRoot')
    if($name -ceq 'File-Dialog'){$text=$text.Replace("('native-command-'+`$command+'-intent.json')","('native-command-'+`$command+'-'+`$script:wire+'-intent.json')")}
    if($name -ceq 'Send-Bridge') {
        Check ($text.Contains('-ResponseTimeoutSeconds 20')) 'Reviewed bridge timeout binding changed.'
        $text=$text.Replace('-ResponseTimeoutSeconds 20','-ResponseTimeoutSeconds 60')
    }
    . ([scriptblock]::Create($text))
}
function Launch-Owned {
    Check (-not $script:swChild -and -not $script:hostProcess) 'Previous original host must be released before one direct launch.'
    Check (@(Get-Process SLDWORKS -ErrorAction SilentlyContinue).Count -eq 0) 'Existing SOLIDWORKS prevents another launch.'
    Write-Json (Join-Path $plan.EvidenceRoot 'host-launch-intent.json') @{Executable=$plan.SolidWorksExecutable;Desktop='Default';Delivery='One CreateProcessW WinSta0\\Default';Attempts=1;VisualStudio='NOT_APPLICABLE; existing user IDE preserved'}
    $script:swChild=$desktopType.GetMethod('LaunchSolidWorks',$flags).Invoke($null,@([string]$plan.SolidWorksExecutable,[string]$plan.EvidenceRoot,[string]'Default'))
    $hostPid=$script:swChild.GetType().GetProperty('ProcessId',$script:instanceFlags).GetValue($script:swChild,$null)
    $hostThread=$script:swChild.GetType().GetProperty('ThreadId',$script:instanceFlags).GetValue($script:swChild,$null)
    $script:hostOriginalHandle=$script:swChild.GetType().GetProperty('ProcessHandle',$script:instanceFlags).GetValue($script:swChild,$null)
    $hostOriginalThread=$script:swChild.GetType().GetProperty('ThreadHandle',$script:instanceFlags).GetValue($script:swChild,$null)
    $script:hostBirth=$null
    # The original child is assigned before receipt/SDK reads. A refused identity never loses ownership.
    Write-Claim (Join-Path $plan.EvidenceRoot 'host-original-create.json') @{Pid=$hostPid;ThreadId=$hostThread;OriginalHandle=$script:hostOriginalHandle.ToInt64();OriginalThreadHandle=$hostOriginalThread.ToInt64();HandleSource='OriginalCreateProcessHandle';CreateProcessHandleProven=$true;WorkerPid=$PID;ExpectedImagePath=$plan.SolidWorksExecutable;IdentityVerified=$false;Desktop='Default';Utc=[DateTime]::UtcNow.ToString('o')}
    $desktopType.GetMethod('ValidateSolidWorksChild',$flags).Invoke($null,@($script:swChild))|Out-Null
    $script:hostBirth=$script:swChild.GetType().GetProperty('BirthUtc',$script:instanceFlags).GetValue($script:swChild,$null)
    $actualPath=$script:swChild.GetType().GetProperty('ImagePath',$script:instanceFlags).GetValue($script:swChild,$null)
    Check ($script:swChild.GetType().GetProperty('IdentityVerified',$script:instanceFlags).GetValue($script:swChild,$null) -and $actualPath -ieq $plan.SolidWorksExecutable) 'Selected original native SOLIDWORKS identity differs.'
    $script:hostProcess=[Diagnostics.Process]::GetProcessById($hostPid)
    Write-Json (Join-Path $plan.EvidenceRoot 'host-original-process.json') @{Pid=$hostPid;ThreadId=$hostThread;StartUtc=$script:hostBirth;ImagePath=$actualPath;ImageSha256=(Hash $actualPath);OriginalHandle=$script:hostOriginalHandle.ToInt64();HandleSource='OriginalCreateProcessHandle';CreateProcessHandleProven=$true;IdentityVerified=$true;IdentitySource='GetProcessTimes+QueryFullProcessImageNameW on original creation handle';WorkerPid=$PID;Desktop='Default';FrameVerified=$false;RevisionVerified=$false}
    $deadline=[DateTime]::UtcNow.AddSeconds(90)
    $toolbarAcknowledged=$false;$toolbarClosed=$false;$startupObservation=0;$toolbarModal=0L
    $rotReady=$false;$revision=$null;$rotObservation=0
    do {
        $desktopType.GetMethod('RequireOfficeWindowInventory',$flags).Invoke($null,@([string]$env:VBAi_TEST_DESKTOP_NAME,[uint32]$script:hostProcess.Id,[bool]$false,[IntPtr]::Zero))|Out-Null
        $windows=@([Q020Startup]::Snapshot($script:hostProcess.Id,$deadline))
        Write-Json (Join-Path $plan.EvidenceRoot ('startup-windows-'+(++$startupObservation).ToString('D3')+'.json')) @{Utc=[DateTime]::UtcNow.ToString('o');Windows=$windows;Deadline=$deadline}
        $button=[Q020Startup]::KnownToolbarButton($windows)
        if($button) {
            if($toolbarAcknowledged) { Check (-not $toolbarClosed -and $button.Root -eq $toolbarModal) 'Toolbar warning reappeared/changed after one acknowledgment; no retry.' }
            else {
                $toolbarModal=$button.Root;$toolbarAcknowledged=$true
                $ownedGuard=[Action]{Desktop-Check;$desktopType.GetMethod('RequireOfficeWindowInventory',$flags).Invoke($null,@([string]$env:VBAi_TEST_DESKTOP_NAME,[uint32]$hostPid,[bool]$true,[IntPtr]::new($toolbarModal)))|Out-Null}
                [Q020Startup]::Acknowledge($windows,(Join-Path $plan.EvidenceRoot 'toolbar-warning-intent.json'),$deadline,$ownedGuard)
            }
        } elseif($toolbarAcknowledged) {
            Check (@($windows|Where-Object Handle -eq $toolbarModal).Count -eq 0) 'Original toolbar warning closure unproved.'
            $toolbarClosed=$true
        }
        $frames=@($windows|Where-Object {$_.Class -match '^Afx:' -and $_.Caption -match $plan.NativeFramePattern -and $_.Visible})
        Check ($frames.Count -le 1) 'Actual SOLIDWORKS frame identity is ambiguous.'
        if($frames.Count -eq 1 -and -not $button){
            Window-Check $frames[0]
            Check ([DateTime]::UtcNow -lt $deadline) 'Original startup deadline expired before read-only ROT observation.'
            $rotError=$null
            try {
                $revision=[Q014SolidWorks]::Verify($script:hostProcess.Id)
                Check ($revision -ceq $plan.ExpectedNativeRevision) ('Expected native SOLIDWORKS revision '+$plan.ExpectedNativeRevision)
                $rotReady=$true
            } catch {
                $cause=$_.Exception
                while($cause.InnerException){$cause=$cause.InnerException}
                $rotError=$cause.ToString()
                Write-Json (Join-Path $plan.EvidenceRoot ('startup-rot-'+(++$rotObservation).ToString('D3')+'.json')) @{Pid=$script:hostProcess.Id;Birth=$script:hostBirth;ReadOnly=$true;Ready=$false;Error=$rotError;Deadline=$deadline;Utc=[DateTime]::UtcNow.ToString('o')}
                if($cause -isnot [InvalidOperationException] -or $cause.Message -cne 'Exact SOLIDWORKS PID ROT unavailable; no activation attempted.'){throw}
            }
            if($rotReady){
                Write-Json (Join-Path $plan.EvidenceRoot ('startup-rot-'+(++$rotObservation).ToString('D3')+'.json')) @{Pid=$script:hostProcess.Id;Birth=$script:hostBirth;ReadOnly=$true;Ready=$true;Revision=$revision;Deadline=$deadline;Utc=[DateTime]::UtcNow.ToString('o')}
                break
            }
        }
        Start-Sleep -Milliseconds 500
    }while([DateTime]::UtcNow -lt $deadline)
    Check ($frames.Count -eq 1) 'Actual SOLIDWORKS frame desktop unproved.';Window-Check $frames[0]
    Check ($rotReady -and $revision -ceq $plan.ExpectedNativeRevision -and [DateTime]::UtcNow -lt $deadline) 'Exact owned SOLIDWORKS ROT/revision readiness unproved within original startup deadline.'
    Write-Json (Join-Path $plan.EvidenceRoot 'host-identity.json') @{Pid=$hostPid;Path=$actualPath;StartUtc=$script:hostBirth;OriginalHandle=$script:hostOriginalHandle.ToInt64();HandleSource='OriginalCreateProcessHandle';IdentitySource='Validated original native handle';Revision=$revision;Windows=$windows;Desktop=$env:VBAi_TEST_DESKTOP_NAME}
    if($toolbarAcknowledged){Write-Json (Join-Path $plan.EvidenceRoot 'toolbar-warning-terminal.json') @{AcknowledgedOnce=$true;OriginalModalClosed=$true;NativeFrameVerified=$true;Utc=[DateTime]::UtcNow.ToString('o')}}
}
$launchBody={Launch-Owned}
Add-Type -Path (Join-Path $PSScriptRoot 'Q020MainDesktopGate.cs')
$desktopType=[Q020MainDesktopGate];$flags=[Reflection.BindingFlags]'Static,NonPublic'
$desktopType.GetMethod('Configure',$flags).Invoke($null,@([string]$plan.OwnerSid,[string]$plan.VisualStudioExecutable,[string]$plan.Solution,[string]$campaignRoot,[int]$plan.ProtectedVisualStudioPid))|Out-Null
$desktopType.GetMethod('ConfigureSolidWorksExecutable',$flags).Invoke($null,@([string]$plan.SolidWorksExecutable))|Out-Null
Add-Type -Path (Join-Path $PSScriptRoot 'Q014Native.cs')
Add-Type -Path (Join-Path $PSScriptRoot 'Q020Startup.cs') -ReferencedAssemblies UIAutomationClient,UIAutomationTypes,WindowsBase
Add-Type -Path (Join-Path $PSScriptRoot 'Q020NativeTeardown.cs')
Add-Type -Path (Join-Path $PSScriptRoot 'Q020NativeCapture.cs') -ReferencedAssemblies System.Drawing
Add-Type -Path $plan.LayoutHelper
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class Q020MacroNativeReceipt {
    public sealed class Bar { public long Handle; public uint Pid, Thread; public string Class; public bool Visible; }
    private delegate bool Visitor(IntPtr window, IntPtr state);
    public static Bar[] ReadBars(long rootValue, uint pid, uint thread) {
        var root=new IntPtr(rootValue); RequireRoot(root,pid,thread);
        var rows=new List<Bar>(); Exception failure=null; int count=0;
        Visitor visit=(w,s)=> { try {
            if(++count>512) throw new InvalidOperationException("VBE child inventory bound exceeded.");
            var name=new StringBuilder(256);
            if(GetClassName(w,name,name.Capacity)==0) throw new InvalidOperationException("VBE child class unavailable.");
            if(name.ToString()=="MsoCommandBar") {
                uint p; uint t=GetWindowThreadProcessId(w,out p);
                if(p!=pid || t!=thread || p==0 || t==0 || GetAncestor(w,2)!=root || !IsChild(root,w))
                    throw new InvalidOperationException("Exact VBE command-bar descendant identity unavailable.");
                if(rows.Count>=16) throw new InvalidOperationException("VBE command-bar bound exceeded.");
                rows.Add(new Bar {Handle=w.ToInt64(),Pid=p,Thread=t,Class=name.ToString(),Visible=IsWindowVisible(w)});
            }
            return true;
        } catch(Exception error) {failure=error;return false;} };
        EnumChildWindows(root,visit,IntPtr.Zero); GC.KeepAlive(visit);
        if(failure!=null) throw failure;
        RequireRoot(root,pid,thread);
        if(rows.Count==0) throw new InvalidOperationException("No exact VBE native command-bar candidate.");
        return rows.ToArray();
    }
    private static void RequireRoot(IntPtr root,uint pid,uint thread) {
        uint p; uint t=GetWindowThreadProcessId(root,out p); var name=new StringBuilder(256);
        if(p!=pid || t!=thread || !IsWindowVisible(root) || GetClassName(root,name,name.Capacity)==0 ||
            name.ToString()!="wndclass_desked_gsk" || GetAncestor(root,2)!=root)
            throw new InvalidOperationException("Exact visible VBE root identity changed.");
    }
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr root,Visitor visitor,IntPtr state);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window,out uint pid);
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window,uint flag);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent,IntPtr child);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr window,StringBuilder name,int count);
}
'@
$swInterop=Join-Path (Split-Path $plan.SolidWorksExecutable) 'api/redist/SolidWorks.Interop.sldworks.dll';[Reflection.Assembly]::LoadFrom($swInterop)|Out-Null
Add-Type -Path (Join-Path $PSScriptRoot 'Q014SolidWorks.cs') -ReferencedAssemblies $swInterop
function Write-LoaderContext {
    Desktop-Check;$script:hostProcess.Refresh()
    Check (-not $script:hostProcess.HasExited -and $script:hostProcess.StartTime.ToUniversalTime().ToString('o') -ceq $script:hostBirth -and
        $script:hostProcess.Path -ieq $plan.SolidWorksExecutable) 'Original host generation changed before loader receipt.'
    $roots=@([Q014Native]::Windows($script:hostProcess.Id)|Where-Object {$_.Class -ceq 'wndclass_desked_gsk' -and $_.Visible})
    Check ($roots.Count -eq 1) 'Unique visible owned VBE root required for loader context.'
    Window-Check $roots[0]
    $bars=[Q020MacroNativeReceipt]::ReadBars($roots[0].Handle,$roots[0].Pid,$roots[0].Thread)
    Write-Json (Join-Path $plan.EvidenceRoot 'native-vbe.json') @{HostPid=$script:hostProcess.Id;HostBirthUtc=$script:hostBirth;ImagePath=$plan.SolidWorksExecutable;Desktop=$env:VBAi_TEST_DESKTOP_NAME;RootHwnd=$roots[0].Handle;OwnerThreadId=$roots[0].Thread;RootClass=$roots[0].Class;Visible=$roots[0].Visible;CandidatesMsoBars=@($bars);UiActions=0;Utc=[DateTime]::UtcNow.ToString('o')}
    Write-Json (Join-Path $plan.EvidenceRoot 'main-desktop-context.json') @{ActorPid=$PID;Desktop=$env:VBAi_TEST_DESKTOP_NAME;Scope='Explicitly authorized visible main desktop';NativeWindowOwnerVerified=$true;NoPrivateDesktopClaim=$true;Utc=[DateTime]::UtcNow.ToString('o')}
}
function Require-Connected {
    Desktop-Check
    $status=Send-Bridge @{Command='status'}
    Check ($status.Connected -and $status.HostProcessId -eq $script:hostProcess.Id -and $status.ProcessBitness -eq 64 -and $status.AssemblyModuleVersionId -ceq $plan.ProductMvid -and $status.AssemblyPath -ieq $plan.InstalledProduct -and (Hash $status.AssemblyPath) -ceq $plan.ProductSha256) 'Actual loaded main-desktop candidate differs; no automatic Connect or mutation.'
    $addins=Send-Bridge @{Command='list_addins'}
    $candidate=@($addins.AddIns|Where-Object {$_.Properties.ProgId -ceq 'VBAi.AddIn'})
    Check ($candidate.Count -eq 1 -and $candidate[0].Properties.Connect -eq $true -and ([Guid]$candidate[0].Properties.Guid) -eq [Guid]'8E854243-087F-4D6C-9E0E-8622B0E50883' -and $null -ne $candidate[0].Errors -and @($candidate[0].Errors.PSObject.Properties).Count -eq 0) 'Exact live AddIn GUID/connection/errors differ.'
    Write-Json (Join-Path $plan.EvidenceRoot 'independent-connected-prerequisite.json') @{Connected=$true;HostProcessId=$status.HostProcessId;AssemblyModuleVersionId=$status.AssemblyModuleVersionId;AssemblyPath=$status.AssemblyPath;AssemblySha256=(Hash $status.AssemblyPath);Status=$status;AddIns=$addins;ConnectionObservedBeforeProductScenarios=$true;ConnectSetterEntries=0;AutoStartQualified=$false;Source='Two owned-host read-only bridge responses and exact native window ownership'}
    Write-Json (Join-Path $plan.EvidenceRoot 'loaded-candidate.json') $status
}
function Read-Q020BootstrapReadiness {
    Desktop-Check
    $desktopType.GetMethod('RequireOfficeWindowInventory',$flags).Invoke($null,@([string]'Default',[uint32]$script:hostProcess.Id,[bool]$false,[IntPtr]::Zero))|Out-Null
    $windows=@([Q014Native]::Windows($script:hostProcess.Id))
    $frames=@($windows|Where-Object {$_.Class -match '^Afx:' -and $_.Caption -match $plan.NativeFramePattern -and $_.Visible})
    foreach($frame in $frames){Window-Check $frame}
    $desktopType.GetMethod('RequireOfficeWindowInventory',$flags).Invoke($null,@([string]'Default',[uint32]$script:hostProcess.Id,[bool]$false,[IntPtr]::Zero))|Out-Null
    return @{Frames=$frames;Dialogs=@($windows|Where-Object {$_.Class -ceq '#32770' -and $_.Visible})}
}
function Wait-Q020BootstrapReadiness([ValidateRange(1,30000)][int]$TimeoutMilliseconds=30000) {
    $clock=[Diagnostics.Stopwatch]::StartNew();$observations=@()
    $receipt=Join-Path $plan.EvidenceRoot 'bootstrap-readiness.json'
    for($sample=0;$sample -lt 128;$sample++) {
        $state=Read-Q020BootstrapReadiness
        Check (@($state.Frames).Count -le 1) 'Ambiguous owned SOLIDWORKS frame; no bootstrap command.'
        $ready=@($state.Frames).Count -eq 1 -and @($state.Dialogs).Count -eq 0 -and $clock.ElapsedMilliseconds -lt $TimeoutMilliseconds
        $observations+=@{Sample=$sample;ElapsedMilliseconds=$clock.ElapsedMilliseconds;Frames=@($state.Frames);VisibleDialogs=@($state.Dialogs);Ready=$ready}
        Write-Json $receipt @{State='OBSERVING';TimeoutMilliseconds=$TimeoutMilliseconds;Observations=$observations;Native573Entered=$false;UiActions=0;UnknownDialogsDismissed=0}
        if($ready){Write-Json $receipt @{State='READY';TimeoutMilliseconds=$TimeoutMilliseconds;Observations=$observations;Native573Entered=$false;UiActions=0;UnknownDialogsDismissed=0};return}
        if($clock.ElapsedMilliseconds -ge $TimeoutMilliseconds){break}
        Start-Sleep -Milliseconds 250
    }
    Write-Json $receipt @{State='REFUSED';TimeoutMilliseconds=$TimeoutMilliseconds;Observations=$observations;Native573Entered=$false;UiActions=0;UnknownDialogsDismissed=0}
    throw 'Owned bootstrap frame/dialog readiness deadline; no native command or dialog dismissal.'
}
function Capture([string]$selector,[string]$label,[switch]$WithoutForm) {
    $formName=$script:formNames[$selector]
    $state=Send-Bridge @{Command='project_properties';Project=$selector};$code=@{};$attributes=@{};$resources=@{}
    $exports=Join-Path $plan.EvidenceRoot ('exports-'+$label);[IO.Directory]::CreateDirectory($exports)|Out-Null
    foreach($m in $state.Components) {
        $code[$m.Name]=Send-Bridge @{Command='read_module';Project=$selector;Module=$m.Name}
        if($m.Type -in @(1,2,3)) {
            $ext=$(if($m.Type -eq 1){'.bas'}elseif($m.Type -eq 2){'.cls'}else{'.frm'});$out=Join-Path $exports ($m.Name+$ext)
            $component=Send-Bridge @{Command='component_properties';Project=$selector;Module=$m.Name}
            Send-Bridge @{Command='export_component';Project=$selector;Module=$m.Name;Path=$out;ExpectedComponentVersion=$component.Version}|Out-Null
            $attributes[$m.Name]=@([IO.File]::ReadAllLines($out)|Where-Object {$_ -match '^Attribute '}|ForEach-Object {$_.Trim()}|Sort-Object -CaseSensitive) -join "`n"
            if($m.Type -eq 3){$frx=[IO.Path]::ChangeExtension($out,'.frx');Check ([IO.File]::Exists($frx) -and (Get-Item $frx).Length -gt 0) 'Real FRX absent.';$resources[$m.Name]=@{Path=$frx;Sha256=(Hash $frx);Bytes=(Get-Item $frx).Length}}
        }
    }
    $form=$null;$formTree=$null;$pictureDigest=$null;$formExport=$null
    if(-not $WithoutForm){
        Check ($formName) 'Owned form mapping required.'
        $form=Send-Bridge @{Command='form_state';Project=$selector;Form=$formName}
        $formTree=Send-Bridge @{Command='form_tree';Project=$selector;Form=$formName}
        $picture=@($formTree.Properties|Where-Object Name -ceq 'Picture')
        Check ($picture.Count -eq 1 -and -not $picture[0].Error -and -not [string]::IsNullOrWhiteSpace([string]$picture[0].Digest)) 'Actual UserForm Picture.Digest missing or unreadable.'
        $pictureDigest=[string]$picture[0].Digest
        $formExport=Get-CanonicalExportedForm (Join-Path $exports ($formName+'.frm')) $formName
        Assert-FormReadable $formTree
        $null=Get-IndependentFormRoot $formTree $formExport
        Check ($form.Version -ceq $formTree.TreeVersion) 'Form state and full designer tree revisions differ.'
    }
    # Conditional compilation is native General text, not a VBIDE COM property.
    Check ($code.ContainsKey('Q020Module') -and -not [string]::IsNullOrWhiteSpace($code['Q020Module'].Sha256)) 'Captured owned code revision required before General selection.'
    $selected=Send-Bridge @{Command='select_code';Project=$selector;Module='Q020Module';ExpectedSha256=$code['Q020Module'].Sha256;StartLine=2;ExpectedMode=2}
    Check ($selected.Project -ieq $selector -and $selected.Module -ceq 'Q020Module' -and $selected.Line -eq 2 -and $selected.Mode -eq 2) 'Owned project selection for General differs.'
    $commands=@();$complete=$false
    for($offset=0;$offset -lt 2000;$offset+=200){
        $page=@(Send-Bridge @{Command='list_commands';Offset=$offset;Limit=200});$commands+=$page
        if($page.Count -lt 200){$complete=$true;break}
    }
    Check $complete 'Finite General command inventory incomplete.'
    $captions=@($commands|Where-Object {$_.Id -eq 2578 -and $_.Enabled -and -not [string]::IsNullOrWhiteSpace($_.Caption)}|Select-Object -ExpandProperty Caption -Unique)
    Check ($captions.Count -eq 1) 'Observed General command caption missing or ambiguous.'
    $nativeGeneral=Send-Bridge @{Command='read_project_general';Project=$selector;ExpectedProjectVersion=$state.Version;ExpectedMode=2;ControlCaption=$captions[0]}
    Write-Json (Join-Path $plan.EvidenceRoot ($label+'-native-general.json')) $nativeGeneral
    Check ($nativeGeneral.Available -and $nativeGeneral.Terminal -and $nativeGeneral.OriginalExecuteReturned -and $nativeGeneral.DialogClosed -and
        -not $nativeGeneral.Uncertain -and -not $nativeGeneral.Error -and $nativeGeneral.OpenAttempts -eq 1 -and $nativeGeneral.CancelAttempts -eq 1 -and
        $nativeGeneral.FieldAttempts -eq 0 -and $nativeGeneral.OkAttempts -eq 0 -and -not $nativeGeneral.MutationInvoked -and -not $nativeGeneral.CommittedRequested) 'One readonly native General inspection did not settle.'
    $general=@(@{Name='Name';Value=$nativeGeneral.Name;Error=$null},@{Name='Description';Value=$nativeGeneral.Description;Error=$null},
        @{Name='HelpFile';Value=$nativeGeneral.HelpFile;Error=$null},@{Name='HelpContextID';Value=[int]$nativeGeneral.HelpContextText;Error=$null},
        @{Name='ConditionalCompilation';Value=$nativeGeneral.ConditionalCompilation;Error=$null})
    Check ((Send-Bridge @{Command='project_properties';Project=$selector}).Version -ceq $state.Version) 'Readonly General inspection changed canonical project metadata.'
    return @{General=$general;Components=$state.Components;References=$state.References;Code=$code;Attributes=$attributes;Resources=$resources;Form=$form;FormTree=$formTree;PictureDigest=$pictureDigest;FormExport=$formExport}
}
function Assert-Q020SnapshotEquivalent($before,$after,[switch]$Subset) {
    foreach($name in $before.Code.Keys){Check ($after.Code.ContainsKey($name) -and $after.Code[$name].Code -ceq $before.Code[$name].Code -and $after.Code[$name].Sha256 -ceq $before.Code[$name].Sha256) 'Exact full source differs.'}
    if(-not $Subset){
        $leftComponents=@($before.Components|ForEach-Object {$_.Name+'|'+$_.Type}|Sort-Object)
        $rightComponents=@($after.Components|ForEach-Object {$_.Name+'|'+$_.Type}|Sort-Object)
        Check (($leftComponents -join ';') -ceq ($rightComponents -join ';')) 'Exact typed component manifest differs.'
    }
    $leftReferences=@($before.References|ForEach-Object {([Guid]$_.Guid).ToString('D')+'|'+$_.Major+'|'+$_.Minor+'|'+$_.BuiltIn+'|'+$_.IsBroken}|Sort-Object)
    $rightReferences=@($after.References|ForEach-Object {([Guid]$_.Guid).ToString('D')+'|'+$_.Major+'|'+$_.Minor+'|'+$_.BuiltIn+'|'+$_.IsBroken}|Sort-Object)
    Check (($leftReferences -join ';') -ceq ($rightReferences -join ';')) 'Exact typed reference manifest differs.'
    $left=@($before.General|Where-Object {-not $Subset -or $_.Name -cne 'Name'});$right=@($after.General|Where-Object {-not $Subset -or $_.Name -cne 'Name'})
    foreach($property in $left){
        $actual=@($right|Where-Object Name -ceq $property.Name)
        Check ($actual.Count -eq 1 -and -not $property.Error -and -not $actual[0].Error) 'Exact readable General value missing.'
        if($property.Name -ceq 'HelpContextID'){Check ([int]$property.Value -eq [int]$actual[0].Value) 'General HelpContextID differs.'}
        else{Check ([string]$property.Value -ceq [string]$actual[0].Value) 'General string value differs.'}
    }
    foreach($name in $before.Attributes.Keys){Check ($before.Attributes[$name] -ceq $after.Attributes[$name]) 'Hidden class/form attributes differ.'}
    if($before.Form){Assert-IndependentFormContent $before $after}
    # FRX export byte hashes are recorded, not an oracle: VBE can regenerate nondeterministic resource streams.
    foreach($name in $before.Resources.Keys){Check ($after.Resources.ContainsKey($name) -and $after.Resources[$name].Bytes -gt 0) 'Persisted FRX resource companion absent.'}
}
function Compile-Owned([string]$selector,[string]$label) {
    $code=Send-Bridge @{Command='read_module';Project=$selector;Module='Q020Module'}
    $selected=Send-Bridge @{Command='select_code';Project=$selector;Module='Q020Module';ExpectedSha256=$code.Sha256;StartLine=2;ExpectedMode=2}
    Check ($selected.Project -ieq $selector -and $selected.Module -ceq 'Q020Module' -and $selected.Line -eq 2 -and $selected.Mode -eq 2) 'Exact owned project selection before compile differs.'
    Write-Json (Join-Path $plan.EvidenceRoot ($label+'-compile-owned-selection.json')) $selected
    $compiled=Send-Bridge @{Command='compile_project';Project=$selector;ExpectedMode=2}
    $dialog=Send-Bridge @{Command='debug_dialog'}
    Check ($null -ne $dialog.PSObject.Properties['Available'] -and $dialog.Available -eq $false -and
        $null -eq $dialog.Diagnostic -and -not $dialog.Error -and @($dialog.Buttons).Count -eq 0) 'Native debug_dialog must explicitly report no available dialog or diagnostic/error.'
    Check ($compiled.Compiled -and
        (Send-Bridge @{Command='debug_state';Project=$selector}).Mode -eq 2) 'One owned native compile/design state not verified.'
    Write-Json (Join-Path $plan.EvidenceRoot ($label+'-compiled.json')) $compiled
}
function Capture-Designer([string]$selector,[string]$label,$settledOpened=$null) {
 Desktop-Check;$formName=$script:formNames[$selector];Check ($formName) 'Owned mapped form required.'
 $before=Capture $selector ($label+'-before-ui');$beforeProject=Send-Bridge @{Command='project_properties';Project=$selector};$beforeBytes=Read-OwnedFileHash $selector
 if($null -eq $settledOpened){$opened=Send-Bridge @{Command='open_form';Project=$selector;Form=$formName}}
 else{Check ($settledOpened.Project -ieq $selector -and $settledOpened.Form -ceq $formName -and -not [string]::IsNullOrWhiteSpace([string]$settledOpened.Version)) 'Reused opened-designer response identity/version differs.';$opened=$settledOpened}
 $windows=Send-Bridge @{Command='vbe_windows'};$designers=@($windows.Windows|Where-Object {$_.Properties.Type -eq 1 -and $_.Properties.Visible -and $_.Properties.Caption -match ([Regex]::Escape($formName))})
 Check ($designers.Count -eq 1 -and @($designers[0].Errors.PSObject.Properties).Count -eq 0) 'Exact mapped visible designer required.'
 $focused=Send-Bridge @{Command='show_vbe_window';WindowCaption=$designers[0].Properties.Caption;WindowType=1}
 Check ($focused.Visible -and $focused.FocusVerified -and $focused.ActiveWindow.Properties.Type -eq 1 -and $focused.ActiveWindow.Properties.Caption -ceq $designers[0].Properties.Caption) 'Exact designer active readback differs.'
 # Discover native Toolbox only after the last designer focus; HWND recreation is valid and must be observed.
 $roots=@([Q014Native]::Windows($script:hostProcess.Id)|Where-Object {$_.Class -ceq 'wndclass_desked_gsk' -and $_.Visible});Check ($roots.Count -eq 1) 'Exact visible VBE root required.';Window-Check $roots[0]
 $owner=[Q020OwnedLayout]::new([int]$script:hostProcess.Id,[string]$script:hostBirth,[string]$plan.SolidWorksExecutable,[long]$roots[0].Handle,[uint32]$roots[0].Thread)
 $toolbox=$owner.DiscoverNativeToolbox()
 Check ($toolbox -and $toolbox.Visible -and $toolbox.Pid -eq $script:hostProcess.Id -and $toolbox.Thread -eq $roots[0].Thread -and $toolbox.Owner -eq $roots[0].Handle -and $toolbox.RootOwner -eq $roots[0].Handle) 'Actual unique native floating Toolbox ownership differs.'
 $layout=$owner.Backup();$capture=$null;$failure=$null;$success=$false
 $guard=[Action]{Desktop-Check;Window-Check $roots[0];$owner.RequireOwner();Check ((Send-Bridge @{Command='project_properties';Project=$selector}).Version -ceq $beforeProject.Version -and (Read-OwnedFileHash $selector) -ceq $beforeBytes) 'Owned source version/bytes changed before capture UI action.'}
 try {
  Write-Json (Join-Path $plan.EvidenceRoot ($label+'-layout-before.json')) @{Layout=$layout;Toolbox=$toolbox;NativeOnlyToolboxMapping=$true;AfterLastDesignerFocus=$true}
  $guard.Invoke();Write-Claim (Join-Path $plan.EvidenceRoot ($label+'-toolbox-hide-intent.json')) @{Attempts=1;Toolbox=$toolbox;NativeCommand='ShowWindow SW_HIDE0';NoVBIDEWindowMapping=$true;NoSave=$true}
  $owner.HideToolbox($toolbox,$guard)
  Write-Json (Join-Path $plan.EvidenceRoot ($label+'-toolbox-hide-return.json')) $toolbox
  Check ($toolbox.HideEntered -and $toolbox.HideReturned -and $toolbox.HiddenVerified) 'One native Toolbox hide unverified; no retry.'
  $guard.Invoke();Write-Claim (Join-Path $plan.EvidenceRoot ($label+'-layout-apply-intent.json')) @{ShowRestoreAttempts=1;SetPositionAttempts=1;WorkArea=$layout.Work;Target=$layout.Target;SetPositionNoActivate=$true;SetPositionNoZOrder=$true;ShowRestoreMayActivate=$true}
  $owner.Apply($layout,$guard)
  $png=Join-Path $plan.EvidenceRoot ($label+'-designer.png');Write-Claim (Join-Path $plan.EvidenceRoot ($label+'-capture-intent.json')) @{Attempts=1;FocusMaximum=1;BarrierMaximum=1;Pid=$script:hostProcess.Id;Birth=$script:hostBirth;Root=$roots[0].Handle;Source=$selector}
  $capture=[Q020NativeCapture]::Begin($script:hostProcess.Id,$script:hostBirth,$plan.SolidWorksExecutable,'Default',[long]$roots[0].Handle,$png)
  $clock=[Diagnostics.Stopwatch]::StartNew();while(-not $capture.Completed -and $clock.Elapsed.TotalSeconds -lt 20){Start-Sleep -Milliseconds 100}
  Write-Json (Join-Path $plan.EvidenceRoot ($label+'-designer-capture.json')) @{Capture=$capture;Opened=$opened;Focused=$focused;Designer=$designers[0];VisualReview='PENDING_MANUAL_REVIEW'}
  Check ($capture.Completed -and $capture.Returned -and -not $capture.Error -and $capture.DpiRestored -and $capture.ForegroundVerifiedBefore -and $capture.ForegroundVerifiedAfter -and $capture.FGReturn -and $capture.FocusAttempts -eq 1 -and $capture.BarrierAttempts -eq 1 -and $capture.Bytes -gt 0 -and (Get-Item -LiteralPath $png).Length -eq $capture.Bytes) 'Owned capture failed or pending; no retry.'
  $success=$true
 }catch{$failure=$_.Exception.ToString()+[Environment]::NewLine+$_.ScriptStackTrace}finally{
  try {
   if($capture -and -not $capture.Completed){throw 'Capture pending; refuse concurrent restoration.'}
   if($layout.ShowRestoreEntered){Write-Claim (Join-Path $plan.EvidenceRoot ($label+'-layout-restore-intent.json')) @{Attempts=1;Before=$layout.Before};$owner.Restore($layout,$guard);Write-Json (Join-Path $plan.EvidenceRoot ($label+'-layout-restored.json')) $layout}
   if($toolbox.HideEntered){$guard.Invoke();Write-Claim (Join-Path $plan.EvidenceRoot ($label+'-toolbox-restore-intent.json')) @{Attempts=1;Toolbox=$toolbox;NativeCommand='ShowWindow SW_SHOWNA8 SAME HWND';NoVBIDEWindowMapping=$true};$owner.RestoreToolbox($toolbox,$guard);Write-Json (Join-Path $plan.EvidenceRoot ($label+'-toolbox-restored.json')) $toolbox;Check ($toolbox.RestoreEntered -and $toolbox.RestoreReturned -and $toolbox.RestoredVerified) 'Original same-HWND Toolbox visibility/geometry restore unverified.'}
   $guard.Invoke();$after=Capture $selector ($label+'-after-ui');Assert-Q020SnapshotEquivalent $before $after
  }catch{$failure=($failure+[Environment]::NewLine+'RESTORATION_OR_SOURCE: '+$_.Exception.ToString());$success=$false}
 }
 Write-Json (Join-Path $plan.EvidenceRoot ($label+'-ui-terminal.json')) @{Success=($success -and -not $failure);Error=$failure;LayoutRestored=$layout.RestoreVerified;ToolboxHideEntered=$toolbox.HideEntered;ToolboxRestoreEntered=$toolbox.RestoreEntered;ToolboxRestored=$toolbox.RestoredVerified;SourceUnchanged=($success -and -not $failure);VisualReview='PENDING_MANUAL_REVIEW';NoMacroExecution=$true}
 Check ($success -and -not $failure) 'Capture/source/layout proof failed; retain without retry.'
}
function Populate([string]$selector) {
    $formName=$script:formNames[$selector];Check ($formName) 'Owned population form mapping required.'
    foreach($entry in @(@{Name='Q020Module';Kind='create_module';Text="Option Explicit`r`nPublic Function OwnedValue() As Long`r`n    OwnedValue = 42`r`nEnd Function"},@{Name='Q020Class';Kind='create_class';Text="Option Explicit`r`nPublic Function OwnedText() As String`r`n    OwnedText = `"Q020 synthetic`"`r`nEnd Function"},@{Name=$formName;Kind='create_form';Text="Option Explicit`r`nPrivate Sub OwnedMarker()`r`n    Dim marker As Long`r`n    marker = 20`r`nEnd Sub"})) {
        $request=@{Command=$entry.Kind;Project=$selector;ExpectedMode=2};if($entry.Kind -ceq 'create_form'){$request.Form=$entry.Name}else{$request.Module=$entry.Name}
        Send-Bridge $request|Out-Null;$read=Send-Bridge @{Command='read_module';Project=$selector;Module=$entry.Name};$c=Send-Bridge @{Command='component_properties';Project=$selector;Module=$entry.Name}
        Send-Bridge @{Command='replace_lines';Project=$selector;Module=$entry.Name;ExpectedSha256=$read.Sha256;StartLine=1;Count=$c.CodeLines;Text=$entry.Text}|Out-Null
    }
    $c=Send-Bridge @{Command='component_properties';Project=$selector;Module='Q020Class'}
    Send-Bridge @{Command='set_class_instancing';Project=$selector;Module='Q020Class';ExpectedComponentVersion=$c.Version;Value=2}|Out-Null
    $f=Send-Bridge @{Command='form_state';Project=$selector;Form=$formName}
    Send-Bridge @{Command='add_form_control';Project=$selector;Form=$formName;ExpectedFormVersion=$f.Version;Control='Q020Label';ControlType='Forms.Label.1';Left=12;Top=12;Width=160;Height=24;Caption='Q020 native owned'}|Out-Null
    $f=Send-Bridge @{Command='form_state';Project=$selector;Form=$formName}
    $picture=Send-Bridge @{Command='set_form_picture';Project=$selector;Form=$formName;ExpectedFormVersion=$f.Version;Path=$plan.Picture}
    Check ($picture.Picture) 'Actual OLE picture fingerprint absent.'
}
function Exit-Originals {
    $frame=$null;$vbeRoot=$null
    if($script:disposableSourceName) {
        $p=Send-Bridge @{Command='project_properties';Project=$script:disposableSourceName}
        Check ($p.Version -ceq $script:disposableSourceVersion) 'Original dirty source changed before teardown; do not discard.'
        $frames=@([Q014Native]::Windows($script:hostProcess.Id)|Where-Object {$_.Class -match '^Afx:' -and $_.Caption -match $plan.NativeFramePattern -and $_.Visible})
        Check ($frames.Count -eq 1) 'Original teardown frame unavailable.';$frame=$frames[0];Window-Check $frame
        $vbeRoots=@([Q014Native]::Windows($script:hostProcess.Id)|Where-Object {$_.Class -ceq 'wndclass_desked_gsk' -and $_.Visible})
        Check ($vbeRoots.Count -eq 1) 'Exact owned VBE teardown root unavailable.';$vbeRoot=$vbeRoots[0];Window-Check $vbeRoot
        Write-Json (Join-Path $plan.EvidenceRoot 'teardown-owned-ui-before.json') @{NativeSwFrame=$frame;VbeRoot=$vbeRoot;Source=$script:disposableSourceName;SourceVersion=$script:disposableSourceVersion;PromptContext='Original owned VBE root; native SW frame lifetime/thread is separate and may end before the VBA prompt'}
    }
    # Before entry retain all original-generation/source/mode/native ownership checks. After entry no PID reopen.
    $ownedPid=[int]$script:hostProcess.Id
    Desktop-Check
    Check (-not $script:hostProcess.HasExited -and $script:hostProcess.StartTime.ToUniversalTime().ToString('o') -ceq $script:hostBirth -and $script:hostProcess.Path -ieq $plan.SolidWorksExecutable) 'Original normal-exit generation changed before entry.'
    Check (-not $script:swChild.GetType().GetMethod('Wait',$script:instanceFlags).Invoke($script:swChild,@([int]0))) 'Original process already exited before the one ExitApp.'
    $exit=[Q014SolidWorks]::BeginNormalExit($ownedPid,$plan.ExpectedNativeRevision);$bound=[DateTime]::UtcNow.AddSeconds(90)
    $discarded=$false;$discardDialog=0L;$discardReturned=$false;$uiStopped=$false;$uiLoss=$null
    $originalWait={ [bool]$script:swChild.GetType().GetMethod('Wait',$script:instanceFlags).Invoke($script:swChild,@([int]0)) }
    while(((-not $exit.Completed) -or (-not (& $originalWait))) -and [DateTime]::UtcNow -lt $bound){
        if($vbeRoot -and -not $uiStopped -and -not (& $originalWait)) {
            try {
                $guard=[Action]{
                    try { $desktopType.GetMethod('RequireCurrent',$flags).Invoke($null,@([string]'Default'))|Out-Null }
                    catch { throw [Q020NativeTeardown+UiContextUnavailableException]::new('Original actor Default context unavailable after ExitApp; no discard.') }
                    if((& $originalWait) -or [DateTime]::UtcNow -ge $bound){throw [Q020NativeTeardown+UiContextUnavailableException]::new('Original exit signaled or delivery deadline ended; no discard.')}
                    [Q020NativeTeardown]::RequireExitingContext($ownedPid,$vbeRoot.Handle,[uint32]$vbeRoot.Thread,$discardDialog)
                }
                $guard.Invoke()
                $prompt=[Q020NativeTeardown]::Observe($ownedPid,$vbeRoot.Handle,$script:disposableSourceName)
                if($prompt) {
                    if($discarded){Check ($prompt.Dialog -eq $discardDialog) 'Named dirty-source prompt reappeared after one discard; no retry.'}
                    else {
                        $discarded=$true;$discardDialog=$prompt.Dialog
                        $intent=Join-Path $plan.EvidenceRoot 'dirty-source-discard-intent.json';Check (-not [IO.File]::Exists($intent)) 'Source discard already claimed.'
                        Write-Claim $intent @{Candidate=$prompt;SourceVersion=$script:disposableSourceVersion;SourcePreservationPreviouslyVerified=$true;Purpose='Owned test teardown only';NoSave=$true;AttemptClaimed=1;OriginalHandle=$script:hostOriginalHandle.ToInt64();LivenessSource='OriginalCreateProcessHandle.Wait(0)'}
                        [Q020NativeTeardown]::Discard($prompt,$guard);$discardReturned=$true
                    }
                }
            } catch {
                $cause=$_.Exception;$contextLoss=$false
                while($cause){if($cause -is [Q020NativeTeardown+UiContextUnavailableException]){$contextLoss=$true;break};$cause=$cause.InnerException}
                if(-not $contextLoss){throw} # Unknown modal, text/shape/control ambiguity remain failures.
                $uiStopped=$true;$uiLoss=$_.Exception.ToString()
                Write-Json (Join-Path $plan.EvidenceRoot 'teardown-ui-context-ended.json') @{Pid=$ownedPid;OriginalHandle=$script:hostOriginalHandle.ToInt64();UiObservationsStopped=$true;NoFurtherDiscard=$true;DiscardClaimed=$discarded;DiscardReturned=$discardReturned;Error=$uiLoss;OnlySameOriginalExitAppAndHandleWaitRemain=$true;Utc=[DateTime]::UtcNow.ToString('o')}
            }
        }
        Start-Sleep -Milliseconds 100
    }
    Check ($exit.Completed -and $exit.Returned -and -not $exit.Error) 'Original ExitApp did not settle; retain without retry.'
    Check ($script:swChild.GetType().GetMethod('Wait',$script:instanceFlags).Invoke($script:swChild,@([int]0))) 'Original CreateProcess host exit unobserved within the single 90-second deadline.'
    $code=$script:swChild.GetType().GetMethod('ExitCode',$script:instanceFlags).Invoke($script:swChild,@())
    Write-Json (Join-Path $plan.EvidenceRoot 'host-original-exit.json') @{Pid=$script:hostProcess.Id;Birth=$script:hostBirth;Handle=$script:hostOriginalHandle.ToInt64();HandleSource='OriginalCreateProcessHandle';ExitObserved=$true;ExitCode=$code;Forced=$false;NormalExitProven=($code -eq 0);NormalApiReturned=$true;DeadlineSeconds=90;UiContextEnded=$uiStopped;DiscardClaimed=$discarded;DiscardReturned=$discardReturned;VisualStudio='NOT_APPLICABLE; no owned IDE launched'}
    if($discarded){Write-Json (Join-Path $plan.EvidenceRoot 'dirty-source-discard-terminal.json') @{Source=$script:disposableSourceName;OriginalDialog=$discardDialog;AttemptClaimed=1;DiscardReturned=$discardReturned;NormalOriginalExitObserved=($code -eq 0);NoSave=$true}}
    Check ($code -eq 0) 'Original host exit abnormal.'
    $script:hostProcess.Dispose();$script:hostProcess=$null;$script:swChild.Dispose();$script:swChild=$null;$script:hostOriginalHandle=[IntPtr]::Zero;$script:disposableSourceName=$null

}
Record
Stage 'managed' {foreach($e in $plan.ManagedEvidence){Check ((Hash $e.Path) -ceq $e.Sha256 -and (Read-Managed $e.Path $plan.TestAssembly).Total -eq $e.Total) 'Current managed evidence changed.'};Write-Json (Join-Path $campaignRoot 'managed-reused.json') @{Evidence=$plan.ManagedEvidence;FreshTestExecution=$false}}
$plan.EvidenceRoot=Join-Path $campaignRoot 'generation-1';[IO.Directory]::CreateDirectory($plan.EvidenceRoot)|Out-Null
Stage 'launch-1' $launchBody
Stage 'bootstrap-1' {Wait-Q020BootstrapReadiness;File-Dialog 573 (Join-Path $plan.EvidenceRoot 'bootstrap.swp') 'Save';Check ([IO.File]::Exists((Join-Path $plan.EvidenceRoot 'bootstrap.swp')) -and (Get-Item -LiteralPath (Join-Path $plan.EvidenceRoot 'bootstrap.swp')).Length -gt 0) 'Owned bootstrap bytes absent.';Write-LoaderContext}
Stage 'connected-1' {Require-Connected}
Stage 'negative' {
    $collection=Send-Bridge @{Command='project_collection_state'};$hosted=@($collection.Projects|Where-Object {$_.Type -eq 100 -and $_.Path -ieq (Join-Path $plan.EvidenceRoot 'bootstrap.swp')});Check ($hosted.Count -eq 1) 'Exact owned bootstrap native100 source required; never select an unrelated macro.'
    $source=$hosted[0].Name;$metadata=Send-Bridge @{Command='project_properties';Project=$source};$dest=Join-Path $plan.EvidenceRoot 'must-not-exist.swp'
    $response=Send-Bridge @{Command='publish_solidworks_macro';Project=$source;Path=$dest;ExpectedMode=2;ExpectedProjectVersion=$metadata.Version} -AllowRefusal
    Check (-not $response.Ok -and $response.Error -match 'standalone|101') 'Expected unsupported-source prewrite refusal not observed.'
    Check (-not [IO.File]::Exists($dest) -and (Send-Bridge @{Command='project_collection_state'}).Version -ceq $collection.Version -and (Send-Bridge @{Command='project_properties';Project=$source}).Version -ceq $metadata.Version) 'Prewrite refusal changed owned state.'
    Write-Json (Join-Path $plan.EvidenceRoot 'negative.json') @{Response=$response;Before=$collection;SetterEntryObservation='NotInstrumented';ExpectedNativeCommandEntries=0;PrewriteRefusalVerified=$true}
}
$aPath=Join-Path $plan.EvidenceRoot 'created-native.swp';$bPath=Join-Path $plan.EvidenceRoot 'published-native.swp';$expected=@{};$savedHashes=@{}
$script:formNames=@{};$script:formNames[$aPath]='Q020CreatedForm';$script:formNames[$bPath]='Q020PublishedForm'
Stage 'create' {
    $c=Send-Bridge @{Command='project_collection_state'};$created=Send-Bridge @{Command='create_solidworks_macro';Path=$aPath;ExpectedMode=2;ExpectedProjectVersion=$c.Version}
    Check ($created.Verified -and $created.Terminal -and -not $created.Uncertain -and $created.CommandAttempts -eq 1 -and $created.FilenameAttempts -eq 1 -and $created.SaveAttempts -eq 1 -and $created.HostPath -ieq $aPath) 'Native creation unverified; never retry.'
    Populate $aPath;Compile-Owned $aPath 'created-native';$p=Send-Bridge @{Command='project_properties';Project=$aPath}
    $save=Send-Bridge @{Command='save_host_document';Project=$aPath;ExpectedMode=2;ExpectedProjectVersion=$p.Version;ExpectedHostPath=$aPath};Check ($save.Verified -and -not $save.Uncertain) 'Native100 save not verified.'
    $expected['create']=Capture $aPath 'create-saved';Capture-Designer $aPath 'create-saved';$savedHashes['create']=Read-OwnedFileHash $aPath;Write-Json (Join-Path $plan.EvidenceRoot 'create-saved.json') @{Creation=$created;Save=$save;Snapshot=$expected['create'];Sha256=$savedHashes['create']}
}
Stage 'publish' {
    $c=Send-Bridge @{Command='project_collection_state'};$draft=Send-Bridge @{Command='create_standalone_project';ExpectedProjectVersion=$c.Version};Check ($draft.Verified -and -not $draft.Uncertain) 'Source101 creation unverified.'
    $source=$draft.Project;$script:formNames[$source]='Q020PublishedForm';Populate $source;Compile-Owned $source 'publication-source';$before=Capture $source 'source-before';$p=Send-Bridge @{Command='project_properties';Project=$source}
    $published=Send-Bridge @{Command='publish_solidworks_macro';Project=$source;Path=$bPath;ExpectedMode=2;ExpectedProjectVersion=$p.Version}
    Check ($published.Verified -and $published.Terminal -and $published.OriginalPreserved -and -not $published.Uncertain -and $published.DestinationProject -cne $source -and $published.HostPath -ieq $bPath) 'Explicit source/new native identity mapping unverified.'
    $after=Capture $source 'source-after';Assert-Q020SnapshotEquivalent $before $after;Check ((Send-Bridge @{Command='project_properties';Project=$source}).Version -ceq $p.Version) 'Publication changed canonical source version.'
    $script:disposableSourceName=$source;$script:disposableSourceVersion=$p.Version
    $expected['publish']=Capture $bPath 'publish-saved'
    $union=@($before.References)+@($published.AddedNativeHostReferences)
    $keys=@($union|ForEach-Object {([Guid]$_.Guid).ToString('D')+'|'+$_.Major+'|'+$_.Minor+'|'+$_.BuiltIn}|Sort-Object)
    $actualKeys=@($expected['publish'].References|ForEach-Object {([Guid]$_.Guid).ToString('D')+'|'+$_.Major+'|'+$_.Minor+'|'+$_.BuiltIn}|Sort-Object)
    Check ($keys.Count -eq @($keys|Select-Object -Unique).Count -and ($keys -join ';') -ceq ($actualKeys -join ';')) 'Destination references must equal source union declared native host references, without loss/duplicate/extra.'
    $mapped=@{};foreach($key in $before.Keys){$mapped[$key]=$before[$key]};$mapped.References=$expected['publish'].References
    Assert-Q020SnapshotEquivalent $mapped $expected['publish'] -Subset;$savedHashes['publish']=Read-OwnedFileHash $bPath
    Write-Json (Join-Path $plan.EvidenceRoot 'publish-saved.json') @{Result=$published;SourceBefore=$before;SourceAfter=$after;Snapshot=$expected['publish'];Sha256=$savedHashes['publish']}
}
Stage 'q030' {
    $c=Send-Bridge @{Command='project_collection_state'};$before=Capture $aPath 'q030-before';$bytes=Read-OwnedFileHash $aPath
    $r=Send-Bridge @{Command='open_standalone_project';Path=$aPath;ExpectedProjectVersion=$c.Version} -AllowRefusal
    Check (-not $r.Ok -and $r.Error -ceq 'Opening a standalone SWP through VBProjects.Open is disabled in SOLIDWORKS after an observed host termination. Use SOLIDWORKS Tools > Macro > Edit instead.') 'Unsafe Open guard differs.'
    Assert-Q020SnapshotEquivalent $before (Capture $aPath 'q030-after');Check ((Read-OwnedFileHash $aPath) -ceq $bytes -and (Send-Bridge @{Command='project_collection_state'}).Version -ceq $c.Version) 'Unsafe Open refusal changed state.'
    Write-Json (Join-Path $plan.EvidenceRoot 'q030-refusal.json') @{Response=$r;Before=$c;SetterEntryObservation='NotInstrumented';NativeApiCallCountClaimed=$false}
}
Stage 'normal-close-1' {Exit-Originals}
$closedHashes=@{}
Stage 'closed-bytes' {
    foreach($entry in @(@{Id='create';Path=$aPath},@{Id='publish';Path=$bPath})){
        $closedHashes[$entry.Id]=Hash $entry.Path;$copy=Join-Path $campaignRoot ($entry.Id+'-closed.swp');Copy-Item -LiteralPath $entry.Path -Destination $copy
        Check ((Hash $copy) -ceq $closedHashes[$entry.Id]) 'Closed archive copy differs.'
        Write-Json (Join-Path $campaignRoot ($entry.Id+'-closed.json')) @{OriginalPath=$entry.Path;SavedSha256=$savedHashes[$entry.Id];ClosedSha256=$closedHashes[$entry.Id];EqualSavedAndClosed=($savedHashes[$entry.Id] -ceq $closedHashes[$entry.Id]);Archive=$copy;ArchiveUsedAsReloadInput=$false}
    }
}
if(-not $script:blocked){$plan.EvidenceRoot=Join-Path $campaignRoot 'generation-2';[IO.Directory]::CreateDirectory($plan.EvidenceRoot)|Out-Null;$script:wire=0}
Stage 'launch-2' $launchBody
Stage 'bootstrap-2' {Wait-Q020BootstrapReadiness;File-Dialog 573 (Join-Path $plan.EvidenceRoot 'bootstrap.swp') 'Save';Check ([IO.File]::Exists((Join-Path $plan.EvidenceRoot 'bootstrap.swp')) -and (Get-Item -LiteralPath (Join-Path $plan.EvidenceRoot 'bootstrap.swp')).Length -gt 0) 'Owned bootstrap bytes absent.';Write-LoaderContext}
Stage 'connected-2' {Require-Connected}
foreach($entry in @(@{Id='create';Path=$aPath},@{Id='publish';Path=$bPath})) {
    Stage ('reload-'+$entry.Id) {
        Check ((Hash $entry.Path) -ceq $closedHashes[$entry.Id]) 'Closed original changed before reload.'
        File-Dialog 84 $entry.Path 'Open';$c=Send-Bridge @{Command='project_collection_state'};$row=@($c.Projects|Where-Object {$_.Path -ieq $entry.Path -and $_.Type -eq 100 -and $_.Mode -eq 2 -and $_.Protection -eq 0})
        Check ($row.Count -eq 1) 'Fresh exact-path native Type100/design identity differs.'
        if($entry.Id -ceq 'publish'){Check (@($c.Projects|Where-Object Type -eq 101).Count -eq 0) 'Publication visual proof requires source101 absent from the fresh host.'}
        $formName=$script:formNames[$entry.Path];Check ($formName) 'Fresh mapped form required before designer materialization.'
        $opened=Send-Bridge @{Command='open_form';Project=$entry.Path;Form=$formName}
        Check ($opened.Project -ieq $entry.Path -and $opened.Form -ceq $formName -and -not [string]::IsNullOrWhiteSpace([string]$opened.Version)) 'One fresh designer materialization response differs; no retry.'
        $actual=Capture $entry.Path ('reload-'+$entry.Id);Assert-Q020SnapshotEquivalent $expected[$entry.Id] $actual;Capture-Designer $entry.Path ('reload-'+$entry.Id) -settledOpened $opened
        Check ((Read-OwnedFileHash $entry.Path) -ceq $closedHashes[$entry.Id]) 'Fresh native Edit changed closed original bytes.'
        Write-Json (Join-Path $plan.EvidenceRoot ($entry.Id+'-reload.json')) @{OriginalPath=$entry.Path;ObservedType=$row[0].Type;SavedDiagnostic=$row[0].Saved;Snapshot=$actual;ClosedInputSha256=$closedHashes[$entry.Id];ReloadSha256=(Read-OwnedFileHash $entry.Path);NoMacroExecution=$true}
    }
}
Stage 'normal-close-2' {Exit-Originals}
if($script:blocked) {
    Write-Json (Join-Path $campaignRoot 'retained-owner.json') @{State='FAILED_OR_UNCERTAIN';HostPid=$(if($script:swChild){$script:swChild.GetType().GetProperty('ProcessId',$script:instanceFlags).GetValue($script:swChild,$null)}else{$null});OriginalHandle=$script:hostOriginalHandle.ToInt64();IdentityVerified=$(if($script:swChild){$script:swChild.GetType().GetProperty('IdentityVerified',$script:instanceFlags).GetValue($script:swChild,$null)}else{$false});NativeQualificationPromoted=$false;NoRetry=$true;VisualStudio='NOT_APPLICABLE; user IDE preserved'}
    # Keep the original native creation handle until actual exit; never substitute a query handle.
    while($script:swChild -and -not $script:swChild.GetType().GetMethod('Wait',$script:instanceFlags).Invoke($script:swChild,@([int]0))){Start-Sleep -Milliseconds 500}
    if($script:swChild){
        $code=$script:swChild.GetType().GetMethod('ExitCode',$script:instanceFlags).Invoke($script:swChild,@())
        if(-not [IO.File]::Exists((Join-Path $plan.EvidenceRoot 'host-original-exit.json'))){Write-Json (Join-Path $plan.EvidenceRoot 'failed-original-host-exit.json') @{Pid=$script:swChild.GetType().GetProperty('ProcessId',$script:instanceFlags).GetValue($script:swChild,$null);Birth=$script:hostBirth;Handle=$script:hostOriginalHandle.ToInt64();HandleSource='OriginalCreateProcessHandle';ExitObserved=$true;ExitCode=$code;NormalExitProven=$false;NativeQualificationPromoted=$false}}
        $script:swChild.Dispose();$script:swChild=$null
    }
    exit 1
}
Write-Json (Join-Path $campaignRoot 'terminal.json') @{State='COMPLETE';AllPlannedStagesPassed=$true;VisualReview='PENDING_MANUAL_REVIEW';RequiredVisualCaptures=@('generation-1/create-saved-designer.png','generation-2/reload-create-designer.png','generation-2/reload-publish-designer.png');PublicationVisualScope='Fresh reopened native100 with source101 absent; no pre-close publication UI acceptance';VisualAcceptanceClaimed=$false;NoMacroExecution=$true;ProductMvid=$plan.ProductMvid;ProductSha256=$plan.ProductSha256;LegacyAdd101SaveAsAcceptance=$false}
exit 0

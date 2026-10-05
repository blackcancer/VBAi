#requires -Version 5.1
param(
    [switch]$Prepare,
    [string]$EvidenceRoot,
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
function Write-Json($Path,$Value) {
    [IO.File]::WriteAllText($Path,($Value|ConvertTo-Json -Depth 40),[Text.UTF8Encoding]::new($true))
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
    Check ([IO.Path]::IsPathRooted($EvidenceRoot) -and [IO.Directory]::Exists($EvidenceRoot)) 'Existing absolute fresh evidence directory required.'
    $planPath=Join-Path $EvidenceRoot 'native-macro-plan.json'
    Check (-not [IO.File]::Exists($planPath)) 'Campaign preparation already exists; never overwrite.'
    $repository=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
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
    foreach($name in @('Q014Native.cs','Q014SolidWorks.cs','Q014VisualStudio.cs','Q020Startup.cs','Q020NativeTeardown.cs','Q020NativeCapture.cs')) { Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $EvidenceRoot $name) }
    # A real owned bitmap exercises binary FRX transport. It is never displayed/executed as a macro.
    Add-Type -AssemblyName System.Drawing
    $bitmap=[Drawing.Bitmap]::new(8,8)
    try { for($x=0;$x -lt 8;$x++){for($y=0;$y -lt 8;$y++){$bitmap.SetPixel($x,$y,[Drawing.Color]::FromArgb(255,($x*30),($y*30),120))}};$bitmap.Save((Join-Path $EvidenceRoot 'owned-picture.bmp'),[Drawing.Imaging.ImageFormat]::Bmp) }
    finally {$bitmap.Dispose()}
    $scenarios=@(
        @{Id='managed';Oracle='Reuse pinned current focused TRX and binary/source manifest; no test replay'},
        @{Id='launch-1';Oracle='One selected private VS Debug.Start; original birth/image/handle and bounded exact warning acknowledgment'},
        @{Id='bootstrap-1';Oracle='Native573 owned bootstrap only to expose VBE; separate from new product creation scope'},
        @{Id='connected-1';Oracle='Explicit independently established loaded candidate prerequisite; no autostart acceptance'},
        @{Id='negative';Oracle='One unsupported source publication prewrite refusal; source/collection unchanged, destination absent'},
        @{Id='create';Oracle='A: one create_solidworks_macro fresh .swp -> unique Type100/ThisLibrary; edit module/class/Q020CreatedForm/picture, compile once, nativeSave once; one visible designer PNG pending manual review'},
        @{Id='publish';Oracle='B: unsaved101 populated and compiled draft with Q020PublishedForm -> explicit new100 mapping; unchanged source full snapshot/version; code/hiddenattributes/designer/picture/references. Publication visual UI is assessed only after fresh reopen with source101 absent; no B-save screenshot'},
        @{Id='q030';Oracle='Known unsafe VBProjects.Open refusal with full state and disk bytes unchanged'},
        @{Id='normal-close-1';Oracle='Normal original SW exit and VS Quit; after source-preservation proof only exact named dirty-source Yes6/No7/Cancel2 prompt permits one No, no save; independent handle exit receipts'},
        @{Id='closed-bytes';Oracle='Archive actual files after normal original exit; record saved/closed byte hashes separately'},
        @{Id='launch-2';Oracle='One fresh selected host; previous owned originals exit proven'},
        @{Id='bootstrap-2';Oracle='Separate owned bootstrap to expose fresh VBE'},
        @{Id='connected-2';Oracle='Fresh exact candidate loaded prerequisite; no Connect setter or autostart claim'},
        @{Id='reload-create';Oracle='NativeEdit84 A same original path once; full code/attributes/FRX semantic designer/picture/refs/General and closed-file hash; Q020CreatedForm designer PNG pending manual review'},
        @{Id='reload-publish';Oracle='NativeEdit84 B same original path once; full persisted readback; source101 absent in fresh host; Q020PublishedForm designer PNG pending manual review; Type100 Saved diagnostic'},
        @{Id='normal-close-2';Oracle='Normal fresh SW and VS exit; original retained handles verified separately'}
    )
    $frozen=@($CandidatePath,$tests,$helper,$base,$ProfileSolution,(Join-Path (Split-Path $ProfileSolution) 'Q014SolidWorks.vcxproj'),(Join-Path (Split-Path $ProfileSolution) 'Q014SolidWorks.vcxproj.user'),$SolidWorksExecutable,$VisualStudioExecutable)
    $frozen+=@(Get-ChildItem $EvidenceRoot -File|Select-Object -ExpandProperty FullName)
    $frozen+=@(Get-ChildItem (Split-Path $tests) -File|Where-Object Extension -in '.dll','.exe','.config'|Select-Object -ExpandProperty FullName)
    $frozen+=@(Get-ChildItem (Join-Path $repository 'src'),(Join-Path $repository 'tests') -Recurse -File|Where-Object {$_.Extension -in '.cs','.csproj','.props','.targets' -and $_.FullName -notmatch '\\(bin|obj)\\'}|Select-Object -ExpandProperty FullName)
    $frozen+=@(Join-Path $repository 'tools/Invoke-VBAi.ps1')
    $frozen+=@($ManagedEvidenceTrxPaths)
    Write-Json $planPath @{Schema='Q020NativeMacroOrderedV1';SourceCommit=(& git -C $repository rev-parse HEAD);SourceStatus=@(& git -C $repository status --porcelain);Repository=$repository;EvidenceRoot=$EvidenceRoot;BaseHarness=$base;BaseHarnessSha256=(Hash $base);InstalledProduct=$CandidatePath;ProductSha256=(Hash $CandidatePath);ProductMvid=[Reflection.Assembly]::ReflectionOnlyLoadFrom($CandidatePath).ManifestModule.ModuleVersionId.ToString('D');TestAssembly=$tests;HelperAssembly=$helper;BuildOutputRoot=$BuildOutputRoot;ManagedEvidence=$managed;SolidWorksYear=$SolidWorksYear;ExpectedNativeRevision=$ExpectedNativeRevision;NativeFramePattern="^SOLIDWORKS.*$SolidWorksYear";SolidWorksExecutable=$SolidWorksExecutable;VisualStudioExecutable=$VisualStudioExecutable;Solution=$ProfileSolution;Picture=(Join-Path $EvidenceRoot 'owned-picture.bmp');Scenarios=$scenarios;FrozenFiles=@($frozen|Select-Object -Unique|ForEach-Object {@{Path=$_;Sha256=(Hash $_)}});NoRetries=$true;NoDesktopSwitch=$true;NoForceTermination=$true;LegacyAdd101SaveAsAcceptance=$false;ConnectionPrerequisite='Independent exact-owner loaded candidate receipt in each generation, bridge status/list_addins corroboration; no autostart claim';ClosedBytesOracle='Capture both saved-before-exit and closed-after-exit hashes; do not erase an observed difference. Reload must preserve closed input bytes and exact complete persisted contents.';PreparedUtc=[DateTime]::UtcNow.ToString('o')}
    Get-Content -LiteralPath $planPath -Raw -Encoding UTF8
    exit 0
}

$plan=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'native-macro-plan.json') -Raw -Encoding UTF8|ConvertFrom-Json
Check ([Threading.Thread]::CurrentThread.ApartmentState -eq 'STA' -and $env:VBAi_TEST_DESKTOP_NAME) 'Reviewed inactive desktop STA worker required.'
foreach($file in $plan.FrozenFiles) {Check ((Hash $file.Path) -ceq $file.Sha256) ('Frozen file changed: '+$file.Path)}
$claim=Join-Path $plan.EvidenceRoot 'execution-claim.json'
$stream=[IO.File]::Open($claim,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
try {$bytes=[Text.Encoding]::UTF8.GetBytes((@{Pid=$PID;Utc=[DateTime]::UtcNow.ToString('o');PlanSha256=(Hash (Join-Path $PSScriptRoot 'native-macro-plan.json'))}|ConvertTo-Json));$stream.Write($bytes,0,$bytes.Length);$stream.Flush($true)}finally{$stream.Dispose()}
$campaignRoot=$plan.EvidenceRoot;$ledger=Join-Path $campaignRoot 'campaign.json';$script:blocked=$false;$script:wire=0
# The registration utility deliberately enables StrictMode. Initialize retained
# ownership before the first launch, so a pre-host failure still observes the VS
# creation handle instead of throwing while constructing its recovery receipt.
$script:hostProcess=$null;$script:vsChild=$null;$script:vsPid=$null
$script:hostBirth=$null;$script:vsBirth=$null;$script:dte=$null
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
$launchAst=@($baseAst.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -ceq 'Stage' -and $n.CommandElements.Count -eq 3 -and $n.CommandElements[1].Value -ceq 'launch'},$true));Check ($launchAst.Count -eq 1) 'Original bounded launch stage missing.'
$launchText=$launchAst[0].CommandElements[2].ScriptBlock.Extent.Text
$launchText=$launchText.Replace('$PSScriptRoot','$script:qualificationScriptRoot')
# Retain stage identities in script scope. The legacy local-variable lifetime is never used as exit evidence.
foreach($name in @('vsPid','vsThread','vsHandle','instanceFlags','dte')) {$launchText=$launchText.Replace('$'+$name,'$script:'+ $name)}
$launchText=$launchText.Replace("Add-Type -Path (Join-Path `$script:qualificationScriptRoot 'Q014VisualStudio.cs') -ReferencedAssemblies `$interop,`$envdte","if(-not ('Q014VisualStudio' -as [type])) { Add-Type -Path (Join-Path `$script:qualificationScriptRoot 'Q014VisualStudio.cs') -ReferencedAssemblies `$interop,`$envdte }")
$launchBody=[scriptblock]::Create($launchText.Substring(1,$launchText.Length-2))
$desktopType=[Reflection.Assembly]::LoadFrom($plan.HelperAssembly).GetType('VBAi.Tests.Integration.IsolatedTestDesktop',$true);$flags=[Reflection.BindingFlags]'Static,NonPublic'
Add-Type -Path (Join-Path $PSScriptRoot 'Q014Native.cs')
Add-Type -Path (Join-Path $PSScriptRoot 'Q020Startup.cs') -ReferencedAssemblies UIAutomationClient,UIAutomationTypes,WindowsBase
Add-Type -Path (Join-Path $PSScriptRoot 'Q020NativeTeardown.cs')
Add-Type -Path (Join-Path $PSScriptRoot 'Q020NativeCapture.cs') -ReferencedAssemblies System.Drawing
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
    Write-Json (Join-Path $plan.EvidenceRoot 'sentinel.json') @{ProcessId=[uint32]$env:VBAi_TEST_DESKTOP_SENTINEL_PID;ThreadId=[uint32]$env:VBAi_TEST_DESKTOP_SENTINEL_TID;Window=[long]$env:VBAi_TEST_DESKTOP_SENTINEL_HWND;Desktop=$env:VBAi_TEST_DESKTOP_NAME;PrivateInputMembershipVerified=$true;Utc=[DateTime]::UtcNow.ToString('o')}
}
function Require-Connected {
    $path=Join-Path $plan.EvidenceRoot 'independent-connected-prerequisite.json'
    $deadline=[DateTime]::UtcNow.AddSeconds(120)
    while(-not [IO.File]::Exists($path) -and [DateTime]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 200}
    Check ([IO.File]::Exists($path)) 'Separate reviewed loaded precondition receipt missing; no automatic Connect.'
    $proof=Get-Content -LiteralPath $path -Raw -Encoding UTF8|ConvertFrom-Json
    Check ($proof.Connected -and $proof.HostProcessId -eq $script:hostProcess.Id -and $proof.AssemblyModuleVersionId -ceq $plan.ProductMvid -and $proof.AssemblyPath -ieq $plan.InstalledProduct) 'Independent connection proof differs.'
    $status=Send-Bridge @{Command='status'}
    Check ($status.Connected -and $status.HostProcessId -eq $script:hostProcess.Id -and $status.ProcessBitness -eq 64 -and $status.AssemblyModuleVersionId -ceq $plan.ProductMvid -and $status.AssemblyPath -ieq $plan.InstalledProduct -and (Hash $status.AssemblyPath) -ceq $plan.ProductSha256) 'Actual loaded candidate differs.'
    $addins=Send-Bridge @{Command='list_addins'}
    $candidate=@($addins.AddIns|Where-Object {$_.Properties.ProgId -ceq 'VBAi.AddIn'})
    Check ($candidate.Count -eq 1 -and $candidate[0].Properties.Connect -eq $true -and
        ([Guid]$candidate[0].Properties.Guid) -eq [Guid]'8E854243-087F-4D6C-9E0E-8622B0E50883' -and
        $null -ne $candidate[0].Errors -and @($candidate[0].Errors.PSObject.Properties).Count -eq 0) 'Exact registered AddIn GUID/connection/errors differ.'
    Write-Json (Join-Path $plan.EvidenceRoot 'loaded-candidate.json') $status
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
            $attributes[$m.Name]=@([IO.File]::ReadAllLines($out)|Where-Object {$_ -match '^Attribute '}) -join "`n"
            if($m.Type -eq 3){$frx=[IO.Path]::ChangeExtension($out,'.frx');Check ([IO.File]::Exists($frx) -and (Get-Item $frx).Length -gt 0) 'Real FRX absent.';$resources[$m.Name]=@{Path=$frx;Sha256=(Hash $frx);Bytes=(Get-Item $frx).Length}}
        }
    }
    $form=$null;$formTree=$null;$pictureDigest=$null
    if(-not $WithoutForm){
        Check ($formName) 'Owned form mapping required.'
        $form=Send-Bridge @{Command='form_state';Project=$selector;Form=$formName}
        $formTree=Send-Bridge @{Command='form_tree';Project=$selector;Form=$formName}
        $picture=@($formTree.Properties|Where-Object Name -ceq 'Picture')
        Check ($picture.Count -eq 1 -and -not $picture[0].Error -and -not [string]::IsNullOrWhiteSpace([string]$picture[0].Digest)) 'Actual UserForm Picture.Digest missing or unreadable.'
        $pictureDigest=[string]$picture[0].Digest
        Check ($form.Version -ceq $formTree.TreeVersion) 'Form state and full designer tree revisions differ.'
    }
    $general=@($state.Properties|Where-Object {$_.Name -in @('Name','Description','HelpFile','HelpContextID','ConditionalCompilationArguments')})
    Check ($general.Count -eq 5) 'Full project General property descriptors missing.'
    return @{General=$general;Components=$state.Components;References=$state.References;Code=$code;Attributes=$attributes;Resources=$resources;Form=$form;FormTree=$formTree;PictureDigest=$pictureDigest}
}
function Compare($before,$after,[switch]$Subset) {
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
    if($before.Form){Check ($before.Form.Version -ceq $after.Form.Version -and ($before.Form.Controls|ConvertTo-Json -Depth 30 -Compress) -ceq ($after.Form.Controls|ConvertTo-Json -Depth 30 -Compress)) 'Designer/control/picture semantic state differs.'}
    if($before.FormTree){Check ($before.FormTree.TreeVersion -ceq $after.FormTree.TreeVersion -and $before.PictureDigest -ceq $after.PictureDigest -and
        ($before.FormTree.Properties|ConvertTo-Json -Depth 40 -Compress) -ceq ($after.FormTree.Properties|ConvertTo-Json -Depth 40 -Compress) -and
        ($before.FormTree.Controls|ConvertTo-Json -Depth 40 -Compress) -ceq ($after.FormTree.Controls|ConvertTo-Json -Depth 40 -Compress)) 'Full designer/control/resource digest readback differs.'}
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
function Capture-Designer([string]$selector,[string]$label) {
    Desktop-Check;$formName=$script:formNames[$selector];Check ($formName) 'Owned designer form mapping required.'
    $opened=Send-Bridge @{Command='open_form';Project=$selector;Form=$formName}
    $windows=Send-Bridge @{Command='vbe_windows'}
    $designers=@($windows.Windows|Where-Object {$_.Properties.Type -eq 1 -and $_.Properties.Visible -and $_.Properties.Caption -match ([Regex]::Escape($formName))})
    Check ($designers.Count -eq 1 -and @($designers[0].Errors.PSObject.Properties).Count -eq 0) 'Unique readable mapped designer caption required; never guess an ambiguous window.'
    $focused=Send-Bridge @{Command='show_vbe_window';WindowCaption=$designers[0].Properties.Caption;WindowType=1}
    Check ($focused.Visible -and $focused.FocusVerified -and $focused.ActiveWindow.Properties.Type -eq 1 -and
        $focused.ActiveWindow.Properties.Caption -ceq $designers[0].Properties.Caption) 'Exact form designer visibility/active readback differs.'
    $roots=@([Q014Native]::Windows($script:hostProcess.Id)|Where-Object {$_.Class -ceq 'wndclass_desked_gsk' -and $_.Visible})
    Check ($roots.Count -eq 1) 'Unique visible VBE capture root required.';Window-Check $roots[0]
    $png=Join-Path $plan.EvidenceRoot ($label+'-designer.png')
    $capture=[Q020NativeCapture]::Begin($script:hostProcess.Id,$script:hostBirth,$plan.SolidWorksExecutable,$env:VBAi_TEST_DESKTOP_NAME,$roots[0].Handle,$png)
    $deadline=[Diagnostics.Stopwatch]::StartNew()
    while(-not $capture.Completed -and $deadline.Elapsed.TotalSeconds -lt 20){Start-Sleep -Milliseconds 100}
    Write-Json (Join-Path $plan.EvidenceRoot ($label+'-designer-capture.json')) @{Capture=$capture;Opened=$opened;SelectedWindow=$designers[0];Focused=$focused;VisualReview='PENDING_MANUAL_REVIEW';NativeCaptureAttempts=1;NoMacroExecution=$true;DeadlineSeconds=20}
    Check ($capture.Completed -and $capture.Returned -and -not $capture.Error -and $capture.Bytes -gt 0 -and
        [IO.File]::Exists($png) -and (Get-Item -LiteralPath $png).Length -eq $capture.Bytes) 'One bounded designer capture failed or remains pending; retain owners without replay.'
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
    $frame=$null
    if($script:disposableSourceName) {
        $p=Send-Bridge @{Command='project_properties';Project=$script:disposableSourceName}
        Check ($p.Version -ceq $script:disposableSourceVersion) 'Original dirty source changed before teardown; do not discard.'
        $frames=@([Q014Native]::Windows($script:hostProcess.Id)|Where-Object {$_.Class -match '^Afx:' -and $_.Caption -match $plan.NativeFramePattern -and $_.Visible})
        Check ($frames.Count -eq 1) 'Original teardown frame unavailable.';$frame=$frames[0];Window-Check $frame
    }
    $exit=[Q014SolidWorks]::BeginNormalExit($script:hostProcess.Id,$plan.ExpectedNativeRevision);$bound=[DateTime]::UtcNow.AddSeconds(90)
    $discarded=$false;$discardDialog=0L
    while(((-not $exit.Completed) -or (-not $script:hostProcess.WaitForExit(0))) -and [DateTime]::UtcNow -lt $bound){
        if($frame -and -not $script:hostProcess.HasExited) {
            Desktop-Check
            $desktopType.GetMethod('RequireOfficeWindowInventory',$flags).Invoke($null,@([string]$env:VBAi_TEST_DESKTOP_NAME,[uint32]$script:hostProcess.Id,[bool]$false,[IntPtr]::Zero))|Out-Null
            $prompt=[Q020NativeTeardown]::Observe($script:hostProcess.Id,$frame.Handle,$script:disposableSourceName)
            if($prompt) {
                if($discarded){Check ($prompt.Dialog -eq $discardDialog) 'Named dirty-source prompt reappeared after one discard; no retry.'}
                else {
                    $discarded=$true;$discardDialog=$prompt.Dialog
                    $intent=Join-Path $plan.EvidenceRoot 'dirty-source-discard-intent.json';Check (-not [IO.File]::Exists($intent)) 'Source discard already claimed.'
                    Write-Json $intent @{Candidate=$prompt;SourceVersion=$script:disposableSourceVersion;SourcePreservationPreviouslyVerified=$true;Purpose='Owned test teardown only; publication source was preserved through accepted readback';NoSave=$true;Attempts=1}
                    $guard=[Action]{Desktop-Check;Check (-not $script:hostProcess.HasExited -and $script:hostProcess.StartTime.ToUniversalTime().ToString('o') -ceq $script:hostBirth -and $script:hostProcess.Path -ieq $plan.SolidWorksExecutable) 'Original source-discard host changed.';$desktopType.GetMethod('RequireOfficeWindowInventory',$flags).Invoke($null,@([string]$env:VBAi_TEST_DESKTOP_NAME,[uint32]$script:hostProcess.Id,[bool]$true,[IntPtr]::new([long]$discardDialog)))|Out-Null;Check ([DateTime]::UtcNow -lt $bound) 'Discard delivery deadline expired.'}
                    [Q020NativeTeardown]::Discard($prompt,$guard)
                }
            }
        }
        Start-Sleep -Milliseconds 100
    }
    Check ($exit.Completed -and $exit.Returned -and -not $exit.Error) 'Original ExitApp did not settle; retain without retry.'
    Check ($script:hostProcess.WaitForExit(0)) 'Original host exit unobserved within the single 90-second deadline.'
    Write-Json (Join-Path $plan.EvidenceRoot 'host-original-exit.json') @{Pid=$script:hostProcess.Id;Birth=$script:hostBirth;Handle=$script:hostProcess.Handle.ToInt64();HandleSource='WorkerFirstQueryHandle';ExitObserved=$true;ExitCode=$script:hostProcess.ExitCode;Forced=$false;NormalApiReturned=$true}
    if($discarded){Write-Json (Join-Path $plan.EvidenceRoot 'dirty-source-discard-terminal.json') @{Source=$script:disposableSourceName;OriginalDialog=$discardDialog;Attempts=1;NormalOriginalExitObserved=$true;NoSave=$true}}
    Check ($script:hostProcess.ExitCode -eq 0) 'Original host exit abnormal.'
    [Q014VisualStudio]::Quit($script:dte,$plan.Solution)
    Check ($script:vsChild.GetType().GetMethod('Wait',$script:instanceFlags).Invoke($script:vsChild,@([int]30000))) 'Original VS exit unobserved.'
    $code=$script:vsChild.GetType().GetMethod('ExitCode',$script:instanceFlags).Invoke($script:vsChild,@())
    Write-Json (Join-Path $plan.EvidenceRoot 'vs-original-exit.json') @{Pid=$script:vsPid;Birth=$script:vsBirth;Handle=$script:vsHandle.ToInt64();HandleSource='OriginalCreateProcessHandle';ExitObserved=$true;ExitCode=$code;Forced=$false}
    Check ($code -eq 0) 'Original VS exit abnormal.'
    $script:hostProcess.Dispose();$script:hostProcess=$null;$script:vsChild.Dispose();$script:vsChild=$null;$script:disposableSourceName=$null
}
Record
Stage 'managed' {foreach($e in $plan.ManagedEvidence){Check ((Hash $e.Path) -ceq $e.Sha256 -and (Read-Managed $e.Path $plan.TestAssembly).Total -eq $e.Total) 'Current managed evidence changed.'};Write-Json (Join-Path $campaignRoot 'managed-reused.json') @{Evidence=$plan.ManagedEvidence;FreshTestExecution=$false}}
$plan.EvidenceRoot=Join-Path $campaignRoot 'generation-1';[IO.Directory]::CreateDirectory($plan.EvidenceRoot)|Out-Null
Stage 'launch-1' $launchBody
Stage 'bootstrap-1' {File-Dialog 573 (Join-Path $plan.EvidenceRoot 'bootstrap.swp') 'Save';Write-LoaderContext}
Stage 'connected-1' {Require-Connected}
Stage 'negative' {
    $collection=Send-Bridge @{Command='project_collection_state'};$hosted=@($collection.Projects|Where-Object Type -eq 100);Check ($hosted.Count -gt 0) 'Owned unsupported native100 source absent.'
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
    $after=Capture $source 'source-after';Compare $before $after;Check ((Send-Bridge @{Command='project_properties';Project=$source}).Version -ceq $p.Version) 'Publication changed canonical source version.'
    $script:disposableSourceName=$source;$script:disposableSourceVersion=$p.Version
    $expected['publish']=Capture $bPath 'publish-saved'
    $union=@($before.References)+@($published.AddedNativeHostReferences)
    $keys=@($union|ForEach-Object {$_.Guid.ToUpperInvariant()+'|'+$_.Major+'|'+$_.Minor+'|'+$_.BuiltIn}|Sort-Object)
    $actualKeys=@($expected['publish'].References|ForEach-Object {$_.Guid.ToUpperInvariant()+'|'+$_.Major+'|'+$_.Minor+'|'+$_.BuiltIn}|Sort-Object)
    Check ($keys.Count -eq @($keys|Select-Object -Unique).Count -and ($keys -join ';') -ceq ($actualKeys -join ';')) 'Destination references must equal source union declared native host references, without loss/duplicate/extra.'
    $mapped=@{};foreach($key in $before.Keys){$mapped[$key]=$before[$key]};$mapped.References=$expected['publish'].References
    Compare $mapped $expected['publish'] -Subset;$savedHashes['publish']=Read-OwnedFileHash $bPath
    Write-Json (Join-Path $plan.EvidenceRoot 'publish-saved.json') @{Result=$published;SourceBefore=$before;SourceAfter=$after;Snapshot=$expected['publish'];Sha256=$savedHashes['publish']}
}
Stage 'q030' {
    $c=Send-Bridge @{Command='project_collection_state'};$before=Capture $aPath 'q030-before';$bytes=Read-OwnedFileHash $aPath
    $r=Send-Bridge @{Command='open_standalone_project';Path=$aPath;ExpectedProjectVersion=$c.Version} -AllowRefusal
    Check (-not $r.Ok -and $r.Error -ceq 'Opening a standalone SWP through VBProjects.Open is disabled in SOLIDWORKS after an observed host termination. Use SOLIDWORKS Tools > Macro > Edit instead.') 'Unsafe Open guard differs.'
    Compare $before (Capture $aPath 'q030-after');Check ((Read-OwnedFileHash $aPath) -ceq $bytes -and (Send-Bridge @{Command='project_collection_state'}).Version -ceq $c.Version) 'Unsafe Open refusal changed state.'
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
Stage 'bootstrap-2' {File-Dialog 573 (Join-Path $plan.EvidenceRoot 'bootstrap.swp') 'Save';Write-LoaderContext}
Stage 'connected-2' {Require-Connected}
foreach($entry in @(@{Id='create';Path=$aPath},@{Id='publish';Path=$bPath})) {
    Stage ('reload-'+$entry.Id) {
        Check ((Hash $entry.Path) -ceq $closedHashes[$entry.Id]) 'Closed original changed before reload.'
        File-Dialog 84 $entry.Path 'Open';$c=Send-Bridge @{Command='project_collection_state'};$row=@($c.Projects|Where-Object {$_.Path -ieq $entry.Path -and $_.Type -eq 100 -and $_.Mode -eq 2 -and $_.Protection -eq 0})
        Check ($row.Count -eq 1) 'Fresh exact-path native Type100/design identity differs.'
        if($entry.Id -ceq 'publish'){Check (@($c.Projects|Where-Object Type -eq 101).Count -eq 0) 'Publication visual proof requires source101 absent from the fresh host.'}
        $actual=Capture $entry.Path ('reload-'+$entry.Id);Compare $expected[$entry.Id] $actual;Capture-Designer $entry.Path ('reload-'+$entry.Id)
        Check ((Hash $entry.Path) -ceq $closedHashes[$entry.Id]) 'Fresh native Edit changed closed original bytes.'
        Write-Json (Join-Path $plan.EvidenceRoot ($entry.Id+'-reload.json')) @{OriginalPath=$entry.Path;ObservedType=$row[0].Type;SavedDiagnostic=$row[0].Saved;Snapshot=$actual;ClosedInputSha256=$closedHashes[$entry.Id];ReloadSha256=(Hash $entry.Path);NoMacroExecution=$true}
    }
}
Stage 'normal-close-2' {Exit-Originals}
if($script:blocked) {
    Write-Json (Join-Path $campaignRoot 'retained-owner.json') @{State='FAILED_OR_UNCERTAIN';HostPid=$(if($script:hostProcess){$script:hostProcess.Id}else{$null});VsPid=$script:vsPid;NativeQualificationPromoted=$false;NoRetry=$true}
    # Observe both original handles independently; neither exit receipt depends on the other's lifetime.
    $hostRecorded=$false;$vsRecorded=$false
    while(($script:hostProcess -and -not $hostRecorded) -or ($script:vsChild -and -not $vsRecorded)) {
        if($script:hostProcess -and -not $hostRecorded -and $script:hostProcess.WaitForExit(0)){$hostRecorded=$true;Write-Json (Join-Path $plan.EvidenceRoot 'failed-original-host-exit.json') @{Pid=$script:hostProcess.Id;Birth=$script:hostBirth;Handle=$script:hostProcess.Handle.ToInt64();HandleSource='WorkerFirstQueryHandle';ExitObserved=$true;ExitCode=$script:hostProcess.ExitCode;NormalExitProven=$false;NativeQualificationPromoted=$false}}
        if($script:vsChild -and -not $vsRecorded -and $script:vsChild.GetType().GetMethod('Wait',$script:instanceFlags).Invoke($script:vsChild,@([int]0))){$vsRecorded=$true;$code=$script:vsChild.GetType().GetMethod('ExitCode',$script:instanceFlags).Invoke($script:vsChild,@());Write-Json (Join-Path $plan.EvidenceRoot 'failed-original-vs-exit.json') @{Pid=$script:vsPid;Birth=$script:vsBirth;Handle=$script:vsHandle.ToInt64();HandleSource='OriginalCreateProcessHandle';ExitObserved=$true;ExitCode=$code;NormalExitProven=$false;NativeQualificationPromoted=$false}}
        Start-Sleep -Milliseconds 500
    }
    exit 1
}
Write-Json (Join-Path $campaignRoot 'terminal.json') @{State='COMPLETE';AllPlannedStagesPassed=$true;VisualReview='PENDING_MANUAL_REVIEW';RequiredVisualCaptures=@('generation-1/create-saved-designer.png','generation-2/reload-create-designer.png','generation-2/reload-publish-designer.png');PublicationVisualScope='Fresh reopened native100 with source101 absent; no pre-close publication UI acceptance';VisualAcceptanceClaimed=$false;NoMacroExecution=$true;ProductMvid=$plan.ProductMvid;ProductSha256=$plan.ProductSha256;LegacyAdd101SaveAsAcceptance=$false}
exit 0

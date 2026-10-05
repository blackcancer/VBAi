#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$OutputRoot)
$ErrorActionPreference='Stop'
$env:PSModulePath=(Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules')+';'+$env:PSModulePath
$tools=$PSScriptRoot
$BaseHarness=Join-Path $PSScriptRoot '../Invoke-Q020Qualification.ps1'
$runner=Join-Path $tools 'Invoke-Q020MainNativeMacroQualification.ps1'

$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($runner,[ref]$tokens,[ref]$errors)
if($errors.Count){throw ($errors|Out-String)}
$definitions=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst]},$true))
$functions=@{}
foreach($name in @('Check','Assert-Q020SnapshotEquivalent')){
 $matches=@($definitions|Where-Object Name -ceq $name)
 if($matches.Count -ne 1){throw ('Unique actual function required: '+$name)}
 $functions[$name]=$matches[0].Extent.Text
 . ([scriptblock]::Create($functions[$name]))
}
. (Join-Path $tools 'FormOracle.ps1')
$calls=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -ceq 'Assert-Q020SnapshotEquivalent'},$true))
$bare=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -ieq 'Compare'},$true))
if($calls.Count -ne 5 -or $bare.Count -ne 0 -or @($definitions|Where-Object Name -ieq Compare).Count){throw 'Expected five unique-name calls and no old declaration/call.'}
$baseTokens=$null;$baseErrors=$null
$baseAst=[Management.Automation.Language.Parser]::ParseFile($BaseHarness,[ref]$baseTokens,[ref]$baseErrors)
if($baseErrors.Count){throw 'Base AST invalid.'}
$importedNames=@('Read-OwnedFileHash','Desktop-Check','Window-Check','Send-Bridge','File-Dialog')
$names=@($definitions|ForEach-Object Name)+$importedNames
$oracleTokens=$null;$oracleErrors=$null
$oracleAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $tools 'FormOracle.ps1'),[ref]$oracleTokens,[ref]$oracleErrors)
$names+=@($oracleAst.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst]},$true)|ForEach-Object Name)
foreach($name in $importedNames){if(@($baseAst.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq $name},$true)).Count -ne 1){throw 'Imported definition missing/ambiguous.'}}
$aliasNames=@(Get-Alias|ForEach-Object Name)
$collisions=@($names|Where-Object {$_ -iin $aliasNames})
if($collisions.Count){throw ('Function/alias collision: '+($collisions -join ','))}
if((Get-Command Assert-Q020SnapshotEquivalent).CommandType -ne 'Function'){throw 'Actual helper must resolve as Function.'}
if((Get-Alias compare).Definition -cne 'Compare-Object'){throw 'Builtin alias was changed; this canary must preserve it.'}
if(-not [IO.Path]::IsPathRooted($OutputRoot)){throw 'Absolute fresh OutputRoot required.'}
$output=[IO.Path]::GetFullPath($OutputRoot)
if(Test-Path -LiteralPath $output){throw 'Fresh one-shot canary output required.'}
[IO.Directory]::CreateDirectory($output)|Out-Null
$fixturePath=Join-Path $PSScriptRoot 'fixtures/snapshot.json'
$fixture=Get-Content -LiteralPath $fixturePath -Raw -Encoding UTF8|ConvertFrom-Json
$frm=Join-Path $output 'source/Q020PublishedForm.frm'
$targetFrm=Join-Path $output 'target/Q020PublishedForm.frm'
$header="VERSION 5.00`r`nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Q020PublishedForm`r`n   ClientHeight    =   3015`r`n   ClientWidth     =   4000`r`n   OleObjectBlob   =   `"Q020PublishedForm.frx`":0000`r`nEnd`r`nAttribute VB_Name = `"Q020PublishedForm`"`r`nOption Explicit`r`n"
foreach($path in @($frm,$targetFrm)){
 [IO.Directory]::CreateDirectory((Split-Path $path))|Out-Null
 [IO.File]::WriteAllBytes($path,[Text.Encoding]::GetEncoding(28591).GetBytes($header))
 [IO.File]::WriteAllBytes([IO.Path]::ChangeExtension($path,'.frx'),[byte[]](1,2,3,4))
}
function Clone($value){ConvertFrom-Json (ConvertTo-Json -InputObject $value -Depth 100)}
function Snapshot($value){
 $copy=Clone $value;$result=@{}
 foreach($p in $copy.PSObject.Properties){$result[$p.Name]=$p.Value}
 foreach($name in @('Code','Attributes','Resources')){$map=@{};foreach($p in $copy.$name.PSObject.Properties){$map[$p.Name]=$p.Value};$result[$name]=$map}
 return $result
}
function Pair {
 $left=Snapshot $fixture;$right=Snapshot $fixture
 $right.FormTree.Properties+=@(
  [pscustomobject]@{Name='HelpContextID';Kind='scalar';Value=0;Error=$null},
  [pscustomobject]@{Name='ShowModal';Kind='scalar';Value=$true;Error=$null},
  [pscustomobject]@{Name='WhatsThisButton';Kind='scalar';Value=$false;Error=$null},
  [pscustomobject]@{Name='WhatsThisHelp';Kind='scalar';Value=$false;Error=$null})
 $left.FormExport=Get-CanonicalExportedForm $frm 'Q020PublishedForm'
 $right.FormExport=Get-CanonicalExportedForm $targetFrm 'Q020PublishedForm'
 return @($left,$right)
}
$script:cases=@()
function Case([string]$name,[bool]$reject,[scriptblock]$body,[string]$message){
 $failed=$false;$error=$null
 try{& $body}catch{$failed=$true;$error=$_.Exception.Message}
 if($failed -ne $reject -or ($failed -and $message -and $error -notlike ('*'+$message+'*'))){throw ('Unexpected canary outcome: '+$name+' '+$error)}
 $script:cases+=@{Name=$name;Pass=$true;ExpectedRefusal=$reject;Error=$error}
}
function Compare($before,$after,[switch]$Subset){Assert-Q020SnapshotEquivalent $before $after -Subset:$Subset}
Case 'RED Compare-named wrapper shadowed: changed code not rejected' $false {$p=Pair;$p[1].Code.Q020Module.Code+='changed';Compare $p[0] $p[1]|Out-Null} ''
Case 'RED Compare-named wrapper -Subset binds builtin Compare-Object and fails' $true {$p=Pair;Compare $p[0] $p[1] -Subset|Out-Null} 'Subset'
Case 'GREEN actual comparator exact snapshot' $false {$p=Pair;Assert-Q020SnapshotEquivalent $p[0] $p[1]} ''
Case 'GREEN exact source code changed' $true {$p=Pair;$p[1].Code.Q020Module.Code+='changed';Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'Exact full source differs'
Case 'GREEN exact source SHA changed' $true {$p=Pair;$p[1].Code.Q020Module.Sha256='different';Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'Exact full source differs'
Case 'GREEN source component absent' $true {$p=Pair;$p[1].Code.Remove('Q020Module');Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'Exact full source differs'
Case 'GREEN typed component inventory changed' $true {$p=Pair;$p[1].Components[0].Type=999;Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'typed component manifest'
Case 'GREEN reference version changed' $true {$p=Pair;$p[1].References[0].Major++;Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'typed reference manifest'
Case 'GREEN reference broken state changed' $true {$p=Pair;$p[1].References[0].IsBroken=-not $p[1].References[0].IsBroken;Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'typed reference manifest'
Case 'GREEN General string changed' $true {$p=Pair;@($p[1].General|Where-Object Name -ceq Description)[0].Value='changed';Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'General string value'
Case 'GREEN General context changed' $true {$p=Pair;@($p[1].General|Where-Object Name -ceq HelpContextID)[0].Value=42;Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'HelpContextID differs'
Case 'GREEN General context malformed' $true {$p=Pair;@($p[1].General|Where-Object Name -ceq HelpContextID)[0].Value='invalid';Assert-Q020SnapshotEquivalent $p[0] $p[1]} ''
Case 'GREEN General read error' $true {$p=Pair;$p[1].General[0].Error='GetterUnavailable';Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'readable General'
Case 'GREEN General missing' $true {$p=Pair;$p[1].General=@($p[1].General|Where-Object Name -cne HelpFile);Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'readable General'
Case 'GREEN hidden class attribute changed' $true {$p=Pair;$p[1].Attributes.Q020Class+='different';Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'Hidden class/form attributes'
Case 'GREEN FRX snapshot absent' $true {$p=Pair;$p[1].Resources.Clear();Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'FRX resource companion absent'
Case 'GREEN FRX empty' $true {$p=Pair;$p[1].Resources.Q020PublishedForm.Bytes=0;Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'FRX resource companion absent'
Case 'GREEN form font changed' $true {$p=Pair;@($p[1].FormTree.Properties|Where-Object Name -ceq Font)[0].ReadOnly=$true;Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'stable root property content'
Case 'GREEN form picture content changed' $true {$p=Pair;$p[1].PictureDigest='different';Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'picture content digest'
Case 'GREEN control property changed' $true {$p=Pair;$q=@($p[1].FormTree.Controls[0].Properties|Where-Object Name -ceq Visible)[0];$q.Value=-not $q.Value;Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'ordered control'
Case 'GREEN exported header changed' $true {$p=Pair;$p[1].FormExport.FormHeaderCanonical+='different';Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'header byte content'
Case 'GREEN root nondefault missing' $true {$p=Pair;@($p[0].FormTree.Properties|Where-Object Name -ceq Visible)[0].Value=$false;Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'persisted root setting'
Case 'GREEN recursive getter error refuses before projection' $true {$p=Pair;$p[1].FormTree.Controls[0].Properties[0].Error='';Assert-Q020SnapshotEquivalent $p[0] $p[1]} 'Recursive designer getter error'
Case 'GREEN subset permits generated native ThisLibrary and different project Name only' $false {
 $p=Pair;$p[1].Components+=@([pscustomobject]@{Name='ThisLibrary';Type=100});$p[1].Code.ThisLibrary=@{Code='';Sha256='empty'}
 @($p[1].General|Where-Object Name -ceq Name)[0].Value='new_native_project'
 Assert-Q020SnapshotEquivalent $p[0] $p[1] -Subset
} ''
Case 'GREEN subset still refuses differing Description' $true {$p=Pair;@($p[1].General|Where-Object Name -ceq Description)[0].Value='changed';Assert-Q020SnapshotEquivalent $p[0] $p[1] -Subset} 'General string value'
Case 'GREEN subset refuses undeclared extra native reference' $true {
 $p=Pair;$p[1].References+=@([pscustomobject]@{Guid='11111111-2222-3333-4444-555555555555';Major=1;Minor=0;BuiltIn=$true;IsBroken=$false})
 Assert-Q020SnapshotEquivalent $p[0] $p[1] -Subset
} 'typed reference manifest'
Case 'GREEN subset accepts declared reference union mapped identically' $false {
 $p=Pair;$extra=[pscustomobject]@{Guid='11111111-2222-3333-4444-555555555555';Major=1;Minor=0;BuiltIn=$true;IsBroken=$false}
 $p[0].References+=@($extra);$p[1].References+=@(Clone $extra)
 Assert-Q020SnapshotEquivalent $p[0] $p[1] -Subset
} ''
$receipt=@{State='PURE_WRAPPER_CANARY_PASS';Utc=[DateTime]::UtcNow.ToString('o');Cases=$script:cases;ExpectedCallSites=5;ActualCallSites=$calls.Count;BareCompareCalls=$bare.Count;AliasCollisions=$collisions;CheckedFunctionNames=$names;NativeAcceptance=$false;
 Scope='Actual runner functions extracted by AST; old alias red reproducer and unique-wrapper negative/positive calls. Sanitized synthetic snapshot/header/resource fixture; not native exports or persistence acceptance. No host, COM, UI, registry, process launch or product mutation.';
 Pins=@($runner,(Join-Path $tools 'FormOracle.ps1'),$BaseHarness,$fixturePath,$PSCommandPath)|ForEach-Object {@{Path=$_;Sha256=(Get-FileHash -LiteralPath $_).Hash}}}
[IO.File]::WriteAllText((Join-Path $output 'snapshot-comparator-canary.json'),($receipt|ConvertTo-Json -Depth 100),[Text.UTF8Encoding]::new($false))
('PASS '+$script:cases.Count+' pure comparator cases; actual five call sites and alias inventory checked; no native acceptance.')

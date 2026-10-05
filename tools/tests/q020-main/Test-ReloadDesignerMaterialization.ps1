#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$OutputRoot,[string]$RunnerPath)
$ErrorActionPreference='Stop'
if([string]::IsNullOrWhiteSpace($RunnerPath)){$RunnerPath=Join-Path $PSScriptRoot 'Invoke-Q020MainNativeMacroQualification.ps1'}
if(-not [IO.Path]::IsPathRooted($OutputRoot) -or (Test-Path -LiteralPath $OutputRoot)){throw 'Fresh absolute OutputRoot required.'}
[IO.Directory]::CreateDirectory($OutputRoot)|Out-Null
$source=$RunnerPath
$tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($source,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Runner AST errors.'}
function Check($ok,$message){if(-not $ok){throw $message}}
$aPath='owned-A.swp';$bPath='owned-B.swp';$closedHashes=@{create='sha';publish='sha'};$expected=@{create='expected';publish='expected'}
$script:formNames=@{$aPath='Q020CreatedForm';$bPath='Q020PublishedForm'}
$plan=@{EvidenceRoot='offline'};$script:opened=@{};$script:counts=@{Edit=0;Open=0;Capture=0;Designer=0};$script:trace=@()
function Stage([string]$scenarioId,$body){& $body}
function Hash($path){'sha'}
function Read-OwnedFileHash($path){'sha'}
function File-Dialog($id,$path,$caption){Check ($id -eq 84) 'Unexpected native command';$script:counts.Edit++;$script:trace+='edit:'+ $path}
function Send-Bridge($request){
 switch($request.Command){
  'project_collection_state'{return @{Projects=@(@{Path=$aPath;Type=100;Mode=2;Protection=0},@{Path=$bPath;Type=100;Mode=2;Protection=0})}}
  'open_form'{$script:counts.Open++;$script:opened[$request.Project]=$true;$script:trace+='open:'+ $request.Project;return @{Project=$request.Project;Form=$request.Form;Version='settled-version'}}
  default{throw 'Unexpected bridge command in offline loop'}
 }
}
function Capture($path,$label){Check $script:opened[$path] 'Closed designer would fail before inspection';$script:counts.Capture++;$script:trace+='capture:'+ $path;return 'actual'}
function Assert-Q020SnapshotEquivalent($before,$after){Check ($before -ceq 'expected' -and $after -ceq 'actual') 'Comparer binding'}
function Capture-Designer($selector,$label,$settledOpened){Check ($settledOpened.Project -ceq $selector -and $settledOpened.Form -ceq $script:formNames[$selector]) 'Opened result reuse differs';$script:counts.Designer++}
function Write-Json($path,$data){}
$loops=@($ast.FindAll({param($node)$node -is [Management.Automation.Language.ForEachStatementAst] -and $node.Extent.Text -match "Stage \('reload-'"},$true))
Check ($loops.Count -eq 1) 'Unique actual reload loop required'
& ([scriptblock]::Create($loops[0].Extent.Text))
Check ($script:counts.Edit -eq 2 -and $script:counts.Open -eq 2 -and $script:counts.Capture -eq 2 -and $script:counts.Designer -eq 2) 'Exact one open per reload violated'
# Exercise the real designer function prefix through its actual parameter binding, stopping before native windows work.
$definition=@($ast.FindAll({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Capture-Designer'},$true))[0].Extent.Text
$boundary=$definition.IndexOf(" `$windows=Send-Bridge")
Check ($boundary -gt 0) 'Exact native-window boundary missing'
$prefix=$definition.Substring(0,$boundary)+" return `$opened`n}"
function Desktop-Check{}
function Send-Bridge($request){if($request.Command -ceq 'project_properties'){return @{Version='stable'}};if($request.Command -ceq 'open_form'){$script:counts.Open++;return @{Project=$request.Project;Form=$request.Form;Version='settled-version'}};throw 'Unexpected designer prefix request'}
. ([scriptblock]::Create($prefix))
$beforeOpen=$script:counts.Open
$result=Capture-Designer $aPath 'reload-create' -settledOpened @{Project=$aPath;Form='Q020CreatedForm';Version='settled-version'}
Check ($script:counts.Open -eq $beforeOpen -and $result.Form -ceq 'Q020CreatedForm') 'Reused designer caused second open'
$rejected=$false;try{Capture-Designer $aPath 'bad' -settledOpened @{Project=$bPath;Form='Q020CreatedForm';Version='settled-version'}}catch{$rejected=$true}
Check $rejected 'Foreign settled response accepted'
$script:opened[$aPath]=$true
$null=Capture-Designer $aPath 'first-generation'
Check ($script:counts.Open -eq $beforeOpen+1) 'First-generation one-open behavior changed'
$output=@{Passed=$true;NativeCalls=0;ActualStageAndLoopAst=$true;ColdDesignerOpenedBeforeInspection=$true;Native84PerScope=1;OpenFormPerScope=1;NoSecondDesignerOpen=$true;ForeignSettledResponseRejected=$true;FirstGenerationBehaviorPreserved=$true;Trace=$script:trace;SourceSha256=(Get-FileHash -LiteralPath $source).Hash}
$target=Join-Path $OutputRoot 'reload-materialization-canary.json';$stream=[IO.File]::Open($target,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
try{$bytes=[Text.UTF8Encoding]::new($false).GetBytes((ConvertTo-Json -InputObject $output -Depth 10));$stream.Write($bytes,0,$bytes.Length);$stream.Flush($true)}finally{$stream.Dispose()}
$output|ConvertTo-Json -Depth 10

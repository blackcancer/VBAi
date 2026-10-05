#requires -Version 5.1
param([Parameter(Mandatory)][string]$OutputRoot,[string]$RunnerPath)
$ErrorActionPreference='Stop'
if([string]::IsNullOrWhiteSpace($RunnerPath)){$RunnerPath=Join-Path $PSScriptRoot 'Invoke-Q020MainNativeMacroQualification.ps1'}
if(-not [IO.Path]::IsPathRooted($OutputRoot) -or [IO.Directory]::Exists($OutputRoot)){throw 'Fresh absolute canary output required.'}
[IO.Directory]::CreateDirectory($OutputRoot)|Out-Null
$source=$RunnerPath
$tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($source,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Runner AST invalid.'}
foreach($name in @('Check','Wait-Q020BootstrapReadiness')){
 $nodes=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq $name},$true))
 if($nodes.Count -ne 1){throw 'Unique actual helper required.'}
 . ([scriptblock]::Create($nodes[0].Extent.Text))
}
$stages=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -ceq 'Stage' -and $n.CommandElements[1].Extent.Text -in "'bootstrap-1'","'bootstrap-2'"},$true))
if($stages.Count -ne 2){throw 'Both bootstrap stage bodies required.'}
foreach($stage in $stages){if(-not $stage.CommandElements[2].ScriptBlock.Extent.Text.Contains('{Wait-Q020BootstrapReadiness;File-Dialog 573')){throw 'Readiness must precede the first command in both actual stages.'}}
function Start-Sleep {param($Milliseconds)}
function Write-Json($path,$value){$script:lastReceipt=$value}
function Write-LoaderContext {$script:contextCount++}
function Read-Q020BootstrapReadiness {
 $script:reads++;$dialog=@()
 if($script:persistent -or $script:reads -eq 1){$dialog=@([pscustomobject]@{Handle=19;Class='#32770';Visible=$true;Caption='';Pid=123;Thread=7})}
 return @{Frames=@([pscustomobject]@{Handle=18;Class='Afx:owned';Visible=$true;Caption='SOLIDWORKS 2025';Pid=123;Thread=7});Dialogs=$dialog}
}
function File-Dialog($command,$path,$action){if($command -ne 573){throw 'Unexpected command.'};$script:commandCount++;[IO.File]::WriteAllBytes($path,[byte[]](1,2,3,4))}
$rows=@()
foreach($mode in @('transient','persistent')){
 $root=Join-Path $OutputRoot $mode;[IO.Directory]::CreateDirectory($root)|Out-Null
 $plan=@{EvidenceRoot=$root};$script:persistent=$mode -ceq 'persistent';$script:reads=0;$script:commandCount=0;$script:contextCount=0;$script:lastReceipt=$null;$failure=$null
 try{. ([scriptblock]::Create($stages[0].CommandElements[2].ScriptBlock.Extent.Text.Trim().Substring(1).TrimEnd('}')))}catch{$failure=$_.Exception.Message}
 if($mode -ceq 'transient'){
  Check ($null -eq $failure -and $script:reads -eq 2 -and $script:commandCount -eq 1 -and $script:contextCount -eq 1 -and $script:lastReceipt.State -ceq 'READY') 'Transient actual stage must observe closed dialog before exactly one command.'
 }else{
  Check ($failure -like 'Owned bootstrap*' -and $script:commandCount -eq 0 -and $script:contextCount -eq 0 -and $script:lastReceipt.State -ceq 'REFUSED') 'Persistent actual stage must never enter command or dismiss dialog.'
 }
 $rows+=@{Case=$mode;Passed=$true;ReadObservations=$script:reads;Native573Entries=$script:commandCount;DismissEntries=0;ActualStageAst=$true;NativeCalls=0;FinalState=$script:lastReceipt.State}
}
$receipt=@{Passed=$true;SourceSha256=(Get-FileHash -LiteralPath $source).Hash;Cases=$rows;NativeCalls=0;InjectedReadOnlyObservations=$true;OriginalFileDialogPreconditionUnchanged=$true}
[IO.File]::WriteAllText((Join-Path $OutputRoot 'proof.json'),(ConvertTo-Json -InputObject $receipt -Depth 20),[Text.UTF8Encoding]::new($false))
$receipt|ConvertTo-Json -Depth 20

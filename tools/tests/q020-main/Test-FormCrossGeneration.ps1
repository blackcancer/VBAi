#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$OutputRoot,[string]$OraclePath)
$ErrorActionPreference='Stop'
if([string]::IsNullOrWhiteSpace($OraclePath)){$OraclePath=Join-Path $PSScriptRoot 'FormOracle.ps1'}
if(-not [IO.Path]::IsPathRooted($OutputRoot) -or (Test-Path -LiteralPath $OutputRoot)){throw 'Fresh absolute OutputRoot required.'}
[IO.Directory]::CreateDirectory($OutputRoot)|Out-Null
$tokens=$null;$errors=$null;$null=[Management.Automation.Language.Parser]::ParseFile($OraclePath,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Oracle AST errors.'};. $OraclePath
$fixturePath=Join-Path $PSScriptRoot 'fixtures\form-cross-generation-synthetic.json'
$raw=[IO.File]::ReadAllText($fixturePath)
if($raw -match '(?i)([A-Z]:\\|C:/Users/|init-)'){throw 'Machine/user path leaked into synthetic fixture.'}
function Read-Fixture {
 $fixture=ConvertFrom-Json $raw
 foreach($snapshot in @($fixture.CreateSaved,$fixture.CreateReopened,$fixture.PublicationSource,$fixture.PublicationDestination)){
  foreach($name in @('DocumentedDefaults','MetadataDefaults5Values')){
   $map=[ordered]@{};foreach($p in $snapshot.FormExport.$name.PSObject.Properties){$map[$p.Name]=$p.Value};$snapshot.FormExport.$name=$map
  }
 }
 return $fixture
}
$f=Read-Fixture;$before=ConvertTo-Json -InputObject $f -Depth 90 -Compress
Assert-IndependentFormContent $f.CreateSaved $f.CreateReopened
Assert-IndependentFormContent $f.PublicationSource $f.PublicationDestination
if((Get-Q020CrossGenerationFormPair $f.CreateSaved.FormTree $f.CreateReopened.FormTree).ProjectedDescriptorCount -ne 20){throw 'Expected 20 observed generated enum descriptors.'}
if($before -cne (ConvertTo-Json -InputObject $f -Depth 90 -Compress)){throw 'Raw snapshots/revision fields were changed.'}
$cases=@()
foreach($case in @('unknownsuffix','choices','value','error','readonly','status','font','picture','order','numericstring','bool','othernull','layoutreadonly','layoutstatus','layoutvalue','ignoredruntimeerror')){
 $f=Read-Fixture;$r=$f.CreateReopened.FormTree.Properties|Where-Object Name -ceq 'BorderStyle'
 switch($case){
 'unknownsuffix'{$r.Type='2051015560_fmUnknown'}
 'choices'{$r.AllowedValues[0]='changed'}
 'value'{$r.Value=1}
 'error'{$r.Error='failed'}
 'readonly'{$r.ReadOnly=-not $r.ReadOnly}
 'status'{$r.SetterStatus='changed'}
 'font'{($f.CreateReopened.FormTree.Properties|Where-Object Name -ceq 'Font').Members[0].Value='changed'}
 'picture'{($f.CreateReopened.FormTree.Properties|Where-Object Name -ceq 'Picture').Digest='changed'}
 'order'{$rows=$f.CreateReopened.FormTree.Controls[0].Properties;$tmp=$rows[0];$rows[0]=$rows[1];$rows[1]=$tmp}
 'numericstring'{$r.Value='0'}
 'bool'{$r.Value=$false}
 'othernull'{$r.Value=$null}
 'layoutreadonly'{($f.CreateReopened.FormTree.Controls[0].Properties|Where-Object Name -ceq 'LayoutEffect').ReadOnly=$false}
 'layoutstatus'{($f.CreateReopened.FormTree.Controls[0].Properties|Where-Object Name -ceq 'LayoutEffect').SetterStatus='changed'}
 'layoutvalue'{($f.CreateReopened.FormTree.Controls[0].Properties|Where-Object Name -ceq 'LayoutEffect').Value='None'}
 'ignoredruntimeerror'{($f.CreateReopened.FormTree.Properties|Where-Object Name -ceq 'CanUndo').Error='failed'}
 }
 $rejected=$false;try{Assert-IndependentFormContent $f.CreateSaved $f.CreateReopened}catch{$rejected=$true}
 if(-not $rejected){throw ('Strict negative accepted: '+$case)};$cases+=@{Case=$case;Rejected=$true}
}
foreach($case in @('root5-getter-header-mismatch','root5-missing-nondefault')){
 $f=Read-Fixture
 if($case -ceq 'root5-getter-header-mismatch'){($f.PublicationDestination.FormTree.Properties|Where-Object Name -ceq 'HelpContextID').Value=1}
 else{$f.PublicationSource.FormExport.MetadataDefaults5Values['ShowModal']=$false}
 $rejected=$false;try{Assert-IndependentFormContent $f.PublicationSource $f.PublicationDestination}catch{$rejected=$true}
 if(-not $rejected){throw ('ROOT5 negative accepted: '+$case)};$cases+=@{Case=$case;Rejected=$true}
}
$result=@{Passed=$true;NativeCalls=0;SyntheticFixtureOnly=$true;CreateCrossGenerationPositive=$true;PublicationRoot48To51Positive=$true;RawTreesAndRevisionFieldsUnchanged=$true;Cases=$cases;OracleSha256=(Get-FileHash -LiteralPath $OraclePath).Hash;FixtureSha256=(Get-FileHash -LiteralPath $fixturePath).Hash;TestSha256=(Get-FileHash -LiteralPath $PSCommandPath).Hash}
$bytes=[Text.UTF8Encoding]::new($false).GetBytes((ConvertTo-Json -InputObject $result -Depth 12));$stream=[IO.File]::Open((Join-Path $OutputRoot 'form-cross-generation.json'),[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
try{$stream.Write($bytes,0,$bytes.Length);$stream.Flush($true)}finally{$stream.Dispose()}
$result|ConvertTo-Json -Depth 12

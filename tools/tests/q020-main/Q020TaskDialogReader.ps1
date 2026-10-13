#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$Request,[Parameter(Mandatory=$true)][string]$ExpectedSha256)
$ErrorActionPreference='Stop'
if(-not [Environment]::Is64BitProcess -or -not [IO.Path]::IsPathRooted($Request) -or -not (Test-Path -LiteralPath $Request) -or (Get-Item -LiteralPath $Request).Length -gt 65536){throw 'Bounded original x64 observer request required.'}
if((Get-FileHash -LiteralPath $Request).Hash -cne $ExpectedSha256){throw 'Original observer request changed.'}
$raw=Get-Content -LiteralPath $Request -Raw -Encoding UTF8|ConvertFrom-Json
if(-not [IO.Path]::IsPathRooted($raw.Tests) -or $raw.Tests.StartsWith('\\') -or [IO.Path]::GetFileName($raw.Tests) -cne 'VBAi.Tests.dll'){throw 'Canonical local test assembly required before file access.'}
if((Get-FileHash -LiteralPath $raw.Tests).Hash -cne $raw.TestsSha256){throw 'Frozen tests changed before observer load.'}
$tests=[Reflection.Assembly]::LoadFrom($raw.Tests)
Add-Type -AssemblyName System.Web.Extensions
$serializer=[System.Web.Script.Serialization.JavaScriptSerializer]::new()
$serializer.MaxJsonLength=1048576
$requestValue=$serializer.Deserialize((Get-Content -LiteralPath $Request -Raw -Encoding UTF8),[Q020TaskDialogTeardown+ReaderRequest])
$self=[Diagnostics.Process]::GetCurrentProcess()
[Q020TaskDialogTeardown]::ValidateReaderRequest($requestValue,$Request,$self.MainModule.FileName,(Get-FileHash -LiteralPath $raw.Tests).Hash,$tests.ManifestModule.ModuleVersionId.ToString('D'))
if(Test-Path -LiteralPath $requestValue.Receipt){throw 'One-shot observer response already exists.'}
$desktop=$tests.GetType('VBAi.Tests.Integration.IsolatedTestDesktop',$true)
$static=[Reflection.BindingFlags]'Static,NonPublic'
function Context {
 $desktop.GetMethod('RequireMainCurrent',$static).Invoke($null,@())|Out-Null
 if([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -cne $requestValue.UserSid){throw 'Original observer actor changed.'}
 foreach($identity in @(@{Pid=$requestValue.ParentPid;Birth=$requestValue.ParentBirthUtc;Image=$requestValue.ParentImage},@{Pid=$requestValue.HostPid;Birth=$requestValue.HostBirthUtc;Image=$requestValue.HostImage})){
  $process=Get-Process -Id $identity.Pid
  try{if($process.HasExited -or $process.Path -ine $identity.Image -or $process.StartTime.ToUniversalTime().Ticks -ne [DateTime]::Parse($identity.Birth).ToUniversalTime().Ticks){throw 'Original observer parent or host changed.'}}finally{$process.Dispose()}
 }
}
$result=[Q020TaskDialogTeardown+ReaderResponse]::new()
$result.Version=1;$result.Nonce=$requestValue.Nonce;$result.RequestSha256=$ExpectedSha256
$result.TestsSha256=$requestValue.TestsSha256;$result.TestsMvid=$requestValue.TestsMvid
$result.ReaderPid=$PID;$result.ReaderBirthUtc=$self.StartTime.ToUniversalTime().ToString('o');$result.ReaderImage=$self.MainModule.FileName
$result.HostPid=$requestValue.HostPid;$result.HostBirthUtc=$requestValue.HostBirthUtc
$result.ReadOnly=$true;$result.NativeInputSent=$false;$code=1
try{
 Context
 $scope=[Q020TaskDialogTeardown+Scope]::new($requestValue.HostPid,$requestValue.Thread,$requestValue.VbeRoot,$requestValue.NativeRoot,$requestValue.Macro)
 $snapshots=[Q020TaskDialogTeardown]::Observe($scope)
 if($snapshots.Count -gt 0){[Q020TaskDialogTeardown]::Admit($scope,$snapshots)|Out-Null}
 Context
 $result.Snapshots=$snapshots;$result.Success=$true;$code=0
}catch{
 $cause=$_.Exception;while($cause.InnerException){$cause=$cause.InnerException}
 $result.Error=$cause.Message;$result.Snapshots=$null
 if($cause.GetType().FullName -ceq 'Q020TaskDialogTeardown+NotReadyException'){$result.ErrorKind='NotReady';$code=2}
 elseif($cause.GetType().FullName -ceq 'Q020TaskDialogTeardown+ObservationUnavailableException'){$result.ErrorKind='ObservationUnavailable';$code=2}
 elseif($cause -is [InvalidOperationException]){$result.ErrorKind='Rejected'}
 else{$result.ErrorKind='RuntimeError'}
}finally{
 $result.Completed=$true
 $stream=[IO.File]::Open($requestValue.Receipt,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
 try{$bytes=[Text.UTF8Encoding]::new($false).GetBytes(($result|ConvertTo-Json -Depth 12 -Compress));if($bytes.Length -gt 1048576){throw 'Observer response bound exceeded.'};$stream.Write($bytes,0,$bytes.Length);$stream.Flush($true)}finally{$stream.Dispose();$self.Dispose()}
}
exit $code

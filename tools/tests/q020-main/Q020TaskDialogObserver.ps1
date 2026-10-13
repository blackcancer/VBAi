#requires -Version 5.1
# Dot-source only in the owned qualification worker. The child exposes one read operation.
function Read-OwnedTaskDialogs([Q020TaskDialogTeardown+Scope]$scope) {
 Desktop-Check
 $script:hostProcess.Refresh()
 if($script:hostProcess.HasExited -or $scope.Pid -ne $script:hostProcess.Id -or $scope.NativeRoot -ne $frames[0].Handle -or $scope.Thread -ne $frames[0].Thread -or $scope.Macro -cne $plan.Macro -or $script:hostProcess.Path -ine $plan.Executable){throw 'Original task-dialog scope changed.'}
 $nonce=[Guid]::NewGuid().ToString('N')
 if(-not [IO.Path]::IsPathRooted($script:taskDialogCaseRoot)){throw 'Owned case evidence root required.'}
 $folder=Join-Path $script:taskDialogCaseRoot ('task-reader/'+$nonce)
 if(Test-Path -LiteralPath $folder){throw 'One-shot observer directory exists.'}
 [void][IO.Directory]::CreateDirectory($folder)
 $manifest=Join-Path $folder ($nonce+'.request.json')
 $request=[Q020TaskDialogTeardown+ReaderRequest]::new()
 $request.Version=1;$request.Nonce=$nonce;$request.UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value
 $request.ParentPid=$PID;$parent=[Diagnostics.Process]::GetCurrentProcess()
 try{$request.ParentBirthUtc=$parent.StartTime.ToUniversalTime().ToString('o');$request.ParentImage=$parent.MainModule.FileName}finally{$parent.Dispose()}
 $request.HostPid=$scope.Pid;$request.HostBirthUtc=$script:hostProcess.StartTime.ToUniversalTime().ToString('o');$request.HostImage=$script:hostProcess.Path
 $request.Thread=$scope.Thread;$request.VbeRoot=$scope.VbeRoot;$request.NativeRoot=$scope.NativeRoot;$request.Macro=$scope.Macro
 $request.ReaderImage=Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
 $request.Tests=$plan.TestAssembly;$request.TestsSha256=$plan.TestAssemblySha256;$request.TestsMvid=$tests.ManifestModule.ModuleVersionId.ToString('D')
 $request.Receipt=Join-Path $folder ($nonce+'.response.json')
 [Q020TaskDialogTeardown]::ValidateReaderRequest($request,$manifest,$request.ReaderImage,(Get-FileHash -LiteralPath $request.Tests).Hash,$request.TestsMvid)
 $request|ConvertTo-Json -Depth 4|Set-Content -LiteralPath $manifest -Encoding UTF8
 $requestHash=(Get-FileHash -LiteralPath $manifest).Hash
 $entry=Join-Path $plan.Repository 'tools/tests/q020-main/Q020TaskDialogReader.ps1'
 $child=$null;$observed=$false
 try{
  $child=$desktop.GetMethod('LaunchMain',$static).Invoke($null,@([string]$request.ReaderImage,[string[]]@('-NoProfile','-NonInteractive','-STA','-File',$entry,'-Request',$manifest,'-ExpectedSha256',$requestHash),[string]$folder))
  $readerPid=$child.GetType().GetProperty('ProcessId',$instance).GetValue($child,$null)
  $handle=$child.GetType().GetProperty('ProcessHandle',$instance).GetValue($child,$null).ToInt64()
  $birth=[Q020TaskDialogTeardown]::OriginalReaderBirthUtc($handle)
  @{ProcessId=$readerPid;OriginalHandle=$handle;BirthUtc=$birth;RequestSha256=$requestHash;ReadOnly=$true;NoInput=$true}|ConvertTo-Json|Set-Content (Join-Path $folder 'original-create.json') -Encoding UTF8
  $deadline=[DateTime]::UtcNow.AddSeconds(10)
  while(-not ($observed=$child.GetType().GetMethod('Wait',$instance).Invoke($child,@([int]20)))){
   if([DateTime]::UtcNow -ge $deadline){throw 'Original observer exit is unproven; retain its handle and host without force or input.'}
  }
  $exitCode=$child.GetType().GetMethod('ExitCode',$instance).Invoke($child,@())
  @{ProcessId=$readerPid;OriginalHandle=$handle;BirthUtc=$birth;ExitCode=$exitCode;ExitObserved=$true;Forced=$false}|ConvertTo-Json|Set-Content (Join-Path $folder 'original-exit.json') -Encoding UTF8
  if(-not (Test-Path -LiteralPath $request.Receipt) -or (Get-Item -LiteralPath $request.Receipt).Length -gt 1048576){throw 'Bounded observer terminal response unavailable; no input.'}
  Add-Type -AssemblyName System.Web.Extensions
  $serializer=[System.Web.Script.Serialization.JavaScriptSerializer]::new();$serializer.MaxJsonLength=1048576
  $result=$serializer.Deserialize((Get-Content -LiteralPath $request.Receipt -Raw -Encoding UTF8),[Q020TaskDialogTeardown+ReaderResponse])
  return ,([Q020TaskDialogTeardown]::RequireReaderTerminal($request,$result,$requestHash,$readerPid,$birth,$observed,$exitCode))
 }finally{
  if($child){if($observed){$child.Dispose()}else{$script:retainedTaskReaders+=@($child)}}
 }
}

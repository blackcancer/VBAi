#requires -Version 5.1
param([Parameter(Mandatory)][string]$CapturePath,[Parameter(Mandatory)][string]$OutputPath)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition -ne 'Desktop'){throw 'Use Windows PowerShell .NET Framework for the original serialization algorithm.'}
if(-not [IO.Path]::IsPathRooted($CapturePath) -or -not [IO.Path]::IsPathRooted($OutputPath) -or
 (Get-Item -LiteralPath $CapturePath).Length -gt 16MB -or (Test-Path -LiteralPath $OutputPath)){throw 'Bounded retained capture and fresh absolute output paths required.'}
Add-Type -AssemblyName System.Web.Extensions
$json=[Web.Script.Serialization.JavaScriptSerializer]::new();$json.MaxJsonLength=16MB
$record=$json.DeserializeObject([IO.File]::ReadAllText($CapturePath))
$revisions=@()
foreach($capture in $record['Captures']){
 if(-not $capture['Tabs'] -or -not $capture['Request'] -or $capture['Source'] -cne 'FirstChanceClrException.OwningGuardFrame.HeapReadOnly'){throw 'The complete owned guard graph is missing.'}
 $serialized=$json.Serialize($capture['Tabs'])
 $sha=[Security.Cryptography.SHA256]::Create()
 try{$version=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($serialized))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
 $revisions+=@{Capture=$capture['Capture'];Request=$capture['Request'];ObservedOptionsVersion=$version;Tabs=$capture['Tabs'];
  ExpectedMatchesRecomputedObserved=($capture['Request']['ExpectedOptionsVersion'] -ieq $version);Algorithm='Historical .NET Framework JavaScriptSerializer(Tabs), UTF-8, SHA-256'}
}
[IO.File]::WriteAllText($OutputPath,$json.Serialize(@{Scope='Offline captured CLR graph; no host access or native mutation';Revisions=$revisions}),[Text.UTF8Encoding]::new($false))
Write-Output ('Captured revisions: '+$revisions.Count)

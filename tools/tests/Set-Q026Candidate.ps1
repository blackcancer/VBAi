#requires -Version 5.1
param([string]$CandidateDirectory,[string]$ExpectedCurrentProduct,
    [Parameter(Mandatory=$true)][string]$BackupPath,[switch]$Restore)
$ErrorActionPreference='Stop'
$hosts=@(Get-Process EXCEL,WINWORD,POWERPNT,MSACCESS,MSPUB,SLDWORKS -ErrorAction SilentlyContinue)
if($hosts.Count){throw 'A VBE host is still loaded; registration is not changed.'}
$server='Software\Classes\CLSID\{8E854243-087F-4D6C-9E0E-8622B0E50883}\InprocServer32'
$registry=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryView]::Registry64)
try {
    if($Restore){
        $backup=Import-Clixml -LiteralPath $BackupPath
        foreach($entry in $backup.Entries){
            if($entry.Path -ne $server -and $entry.Path -ne ($server+'\'+$backup.Version)){throw 'Unreviewed registry path in backup.'}
            $key=$registry.OpenSubKey($entry.Path,$true)
            try{if(-not $key -or [string]$key.GetValue('CodeBase') -cne $backup.AppliedCodeBase){throw 'Registration changed outside this qualification; no overwrite.'}}
            finally{if($key){$key.Dispose()}}
        }
        foreach($entry in $backup.Entries){$key=$registry.OpenSubKey($entry.Path,$true);try{$key.SetValue('CodeBase',$entry.Value,[Microsoft.Win32.RegistryValueKind]$entry.Kind)}finally{$key.Dispose()}}
        foreach($entry in $backup.Entries){$key=$registry.OpenSubKey($entry.Path);try{if([string]$key.GetValue('CodeBase') -cne $entry.Value){throw 'Original CodeBase restoration readback failed.'}}finally{$key.Dispose()}}
        Write-Output 'Original per-user CodeBase restored.'
    } else {
        if(-not [IO.Path]::IsPathRooted($BackupPath) -or (Test-Path -LiteralPath $BackupPath)){throw 'A fresh absolute registration backup is required.'}
        $dll=Join-Path $CandidateDirectory 'VBAi.dll'
        if(-not [IO.Path]::IsPathRooted($CandidateDirectory) -or -not (Test-Path -LiteralPath $dll) -or -not [IO.Path]::IsPathRooted($ExpectedCurrentProduct)){throw 'Reviewed absolute candidate and original product paths required.'}
        $assembly=[Reflection.Assembly]::ReflectionOnlyLoadFrom($dll);$version=$assembly.GetName().Version.ToString()
        $applied=([Uri][IO.Path]::GetFullPath($dll)).AbsoluteUri
        $entries=@()
        foreach($path in @($server,($server+'\'+$version))){
            $key=$registry.OpenSubKey($path)
            try {
                if(-not $key -or [string]$key.GetValue('Assembly') -cne $assembly.FullName){throw 'Registered assembly identity differs; no registration migration is permitted.'}
                $old=[string]$key.GetValue('CodeBase')
                if(-not $old -or ([Uri]$old).LocalPath -ine $ExpectedCurrentProduct){throw 'Original registered candidate differs; no change.'}
                $entries+=@{Path=$path;Value=$old;Kind=[int]$key.GetValueKind('CodeBase')}
            } finally {if($key){$key.Dispose()}}
        }
        @{Entries=$entries;Version=$version;AppliedCodeBase=$applied;ProductMvid=$assembly.ManifestModule.ModuleVersionId.ToString('D');
            CandidateSha256=(Get-FileHash -LiteralPath $dll).Hash;Hive='HKCU';View='Registry64';Utc=[DateTime]::UtcNow.ToString('o')} |
            Export-Clixml -LiteralPath $BackupPath
        if(@(Get-Process EXCEL,WINWORD,POWERPNT,MSACCESS,MSPUB,SLDWORKS -ErrorAction SilentlyContinue).Count){throw 'A host appeared after preparation; no registration mutation.'}
        foreach($entry in $entries){$key=$registry.OpenSubKey($entry.Path,$true);try{$key.SetValue('CodeBase',$applied,[Microsoft.Win32.RegistryValueKind]$entry.Kind)}finally{$key.Dispose()}}
        foreach($entry in $entries){$key=$registry.OpenSubKey($entry.Path);try{if([string]$key.GetValue('CodeBase') -cne $applied){throw 'Candidate CodeBase readback failed.'}}finally{$key.Dispose()}}
        Write-Output ('Applied reviewed per-user candidate '+$assembly.ManifestModule.ModuleVersionId.ToString('D'))
    }
} finally {$registry.Dispose()}

#requires -Version 5.1
param([string]$CandidateDirectory,[string]$ExpectedCurrentProduct,
    [Parameter(Mandatory=$true)][string]$BackupPath,[switch]$Restore,[switch]$EnableAutoLoad)
$ErrorActionPreference='Stop'
$hosts=@(Get-Process EXCEL,WINWORD,POWERPNT,MSACCESS,MSPUB,SLDWORKS -ErrorAction SilentlyContinue)
if($hosts.Count){throw 'A VBE host is still loaded; registration is not changed.'}
$server='Software\Classes\CLSID\{8E854243-087F-4D6C-9E0E-8622B0E50883}\InprocServer32'
$loadPath='Software\Microsoft\VBA\VBE\6.0\Addins64\VBAi.AddIn'
$registry=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryView]::Registry64)
try {
    if($Restore){
        $backup=Import-Clixml -LiteralPath $BackupPath
        if($backup.AutoLoad){
            if($backup.AutoLoad.Path -cne $loadPath){throw 'Unreviewed add-in load path.'}
            $key=$registry.OpenSubKey($loadPath)
            try{if(-not $key -or $key.GetValue('LoadBehavior') -ne 3){throw 'Add-in loading changed outside this qualification; no overwrite.'}}
            finally{if($key){$key.Dispose()}}
        }
        foreach($entry in $backup.Entries){
            if($entry.Path -ne $server -and $entry.Path -ne ($server+'\'+$backup.Version)){throw 'Unreviewed registry path in backup.'}
            $key=$registry.OpenSubKey($entry.Path,$true)
            try{if(-not $key -or [string]$key.GetValue('CodeBase') -cne $backup.AppliedCodeBase){throw 'Registration changed outside this qualification; no overwrite.'}}
            finally{if($key){$key.Dispose()}}
        }
        foreach($entry in $backup.Entries){$key=$registry.OpenSubKey($entry.Path,$true);try{$key.SetValue('CodeBase',$entry.Value,[Microsoft.Win32.RegistryValueKind]$entry.Kind)}finally{$key.Dispose()}}
        if($backup.AutoLoad){$key=$registry.OpenSubKey($loadPath,$true);try{$key.SetValue('LoadBehavior',$backup.AutoLoad.Value,[Microsoft.Win32.RegistryValueKind]$backup.AutoLoad.Kind)}finally{$key.Dispose()}}
        foreach($entry in $backup.Entries){$key=$registry.OpenSubKey($entry.Path);try{if([string]$key.GetValue('CodeBase') -cne $entry.Value){throw 'Original CodeBase restoration readback failed.'}}finally{$key.Dispose()}}
        if($backup.AutoLoad){$key=$registry.OpenSubKey($loadPath);try{if($key.GetValue('LoadBehavior') -ne $backup.AutoLoad.Value){throw 'Original add-in loading restoration readback failed.'}}finally{$key.Dispose()}}
        Write-Output 'Original per-user CodeBase restored.'
    } else {
        if(-not [IO.Path]::IsPathRooted($BackupPath) -or (Test-Path -LiteralPath $BackupPath)){throw 'A fresh absolute registration backup is required.'}
        $dll=Join-Path $CandidateDirectory 'VBAi.dll'
        if(-not [IO.Path]::IsPathRooted($CandidateDirectory) -or -not (Test-Path -LiteralPath $dll) -or -not [IO.Path]::IsPathRooted($ExpectedCurrentProduct)){throw 'Reviewed absolute candidate and original product paths required.'}
        $assembly=[Reflection.Assembly]::ReflectionOnlyLoadFrom($dll);$version=$assembly.GetName().Version.ToString()
        # Match the repository's managed-COM installer representation, including Unicode paths.
        $applied='file:///'+([IO.Path]::GetFullPath($dll)).Replace([char]92,[char]47)
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
        $autoLoad=$null
        if($EnableAutoLoad){
            $key=$registry.OpenSubKey($loadPath)
            try{
                if(-not $key -or $key.GetValueKind('LoadBehavior') -ne [Microsoft.Win32.RegistryValueKind]::DWord -or $key.GetValue('LoadBehavior') -notin @(0,3)){throw 'Existing reviewed add-in load setting required; no creation or policy change.'}
                $autoLoad=@{Path=$loadPath;Value=$key.GetValue('LoadBehavior');Kind=[int]$key.GetValueKind('LoadBehavior')}
            }finally{if($key){$key.Dispose()}}
        }
        @{Entries=$entries;Version=$version;AppliedCodeBase=$applied;AutoLoad=$autoLoad;ProductMvid=$assembly.ManifestModule.ModuleVersionId.ToString('D');
            CandidateSha256=(Get-FileHash -LiteralPath $dll).Hash;Hive='HKCU';View='Registry64';Utc=[DateTime]::UtcNow.ToString('o')} |
            Export-Clixml -LiteralPath $BackupPath
        if(@(Get-Process EXCEL,WINWORD,POWERPNT,MSACCESS,MSPUB,SLDWORKS -ErrorAction SilentlyContinue).Count){throw 'A host appeared after preparation; no registration mutation.'}
        foreach($entry in $entries){$key=$registry.OpenSubKey($entry.Path,$true);try{$key.SetValue('CodeBase',$applied,[Microsoft.Win32.RegistryValueKind]$entry.Kind)}finally{$key.Dispose()}}
        if($autoLoad){$key=$registry.OpenSubKey($loadPath,$true);try{$key.SetValue('LoadBehavior',3,[Microsoft.Win32.RegistryValueKind]::DWord)}finally{$key.Dispose()}}
        foreach($entry in $entries){$key=$registry.OpenSubKey($entry.Path);try{if([string]$key.GetValue('CodeBase') -cne $applied){throw 'Candidate CodeBase readback failed.'}}finally{$key.Dispose()}}
        if($autoLoad){$key=$registry.OpenSubKey($loadPath);try{if($key.GetValue('LoadBehavior') -ne 3){throw 'Candidate add-in loading readback failed.'}}finally{$key.Dispose()}}
        Write-Output ('Applied reviewed per-user candidate '+$assembly.ManifestModule.ModuleVersionId.ToString('D'))
    }
} finally {$registry.Dispose()}

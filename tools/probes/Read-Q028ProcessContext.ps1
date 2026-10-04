param([Parameter(Mandatory=$true)][string]$EvidenceRoot,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop'
$plan=Get-Content (Join-Path $EvidenceRoot 'q028-plan.json') -Raw -Encoding UTF8|ConvertFrom-Json
$rows=@()
foreach($view in @([Microsoft.Win32.RegistryView]::Registry64,[Microsoft.Win32.RegistryView]::Registry32)){
    $root=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$view)
    try {
        foreach($id in @('{8E854243-087F-4D6C-9E0E-8622B0E50883}','{0F4D723B-97D8-42E5-9B31-70646B97C8D2}')){
            foreach($suffix in @('','\0.1.0.0')){
                $path='Software\Classes\CLSID\'+$id+'\InprocServer32'+$suffix
                $key=$root.OpenSubKey($path,$false)
                try {$rows+=@{View=$view.ToString();Path=$path;CodeBase=$key.GetValue('CodeBase',$null)}}
                catch {$rows+=@{View=$view.ToString();Path=$path;Error=$_.Exception.Message}}
                finally {if($key){$key.Dispose()}}
            }
        }
    } finally {$root.Dispose()}
}
$who=[Security.Principal.WindowsIdentity]::GetCurrent()
$settingsPath=Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)) 'VBAi/settings.json'
$provider=Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8|ConvertFrom-Json
$result=@{Actor=$who.Name;Sid=$who.User.Value;X64=[Environment]::Is64BitProcess;Shell=$PSVersionTable.PSVersion.ToString();SettingsPath=$settingsPath;SettingsProvider=$provider.ProviderName;SettingsModifiedUtc=(Get-Item -LiteralPath $settingsPath).LastWriteTimeUtc.ToString('o');Registry=$rows}
$result|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $OutputPath -Encoding UTF8
Get-Content -LiteralPath $OutputPath -Raw -Encoding UTF8

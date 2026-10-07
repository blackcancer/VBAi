#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Version,
    [Parameter(Mandatory=$true)][string]$PackageDirectory,
    [Parameter(Mandatory=$true)][string]$CompilerPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [string]$SigningScript,
    [switch]$Unsigned
)
$ErrorActionPreference = 'Stop'
if ($Version -notmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(-[0-9A-Za-z.-]+)?$') { throw 'Use a semantic product version.' }
if ($Unsigned -and $SigningScript) { throw 'Select signed or unsigned compilation, not both.' }
if (-not $Unsigned -and -not $SigningScript) { throw 'A signing adapter is required. Unsigned builds need explicit -Unsigned.' }
$package = (Resolve-Path -LiteralPath $PackageDirectory).Path
$compiler = (Resolve-Path -LiteralPath $CompilerPath).Path
$output = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $output) { throw 'Choose a fresh output directory; previous installers are preserved.' }
$required = @('VBAi.dll','VBAi.tlb','VBAi.Updater.exe','EditorAssets\index.html',
    'Licenses\MPL-2.0.txt','Licenses\MIT.txt','Licenses\Scope.md','ThirdPartyNotices',
    'Help\VBAi.en-US.chm','Help\VBAi.fr-FR.chm','vbai-installation.json')
foreach ($item in $required) { if (-not (Test-Path -LiteralPath (Join-Path $package $item))) { throw ('Incomplete installation payload: ' + $item) } }
$marker = Get-Content -LiteralPath (Join-Path $package 'vbai-installation.json') -Raw | ConvertFrom-Json
if ($marker.Product -cne 'VBAi' -or $marker.Version -cne $Version -or $marker.UpdateProtocol -ne 1) { throw 'Payload marker/version mismatch.' }
if ([Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $package 'VBAi.dll')).ProductVersion -cne $Version) { throw 'Assembly product version mismatch.' }
if (@(Get-ChildItem -LiteralPath $package -Recurse -Force | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }).Count) { throw 'Payload reparse points are refused.' }
if (@(Get-ChildItem -LiteralPath $package -Recurse -File | Where-Object { $_.Extension -in '.pdb','.pfx','.key','.trx','.clixml' }).Count) { throw 'Development/private artifacts cannot be packaged.' }
$arguments = @(('/DProductVersion=' + $Version), ('/DPackageDirectory=' + $package), ('/DSetupOutputDirectory=' + $output))
$oldSigner = $env:VBAI_SETUP_SIGNING_SCRIPT
try {
    if (-not $Unsigned) {
        $env:VBAI_SETUP_SIGNING_SCRIPT = (Resolve-Path -LiteralPath $SigningScript).Path
        $callback = Join-Path $PSScriptRoot 'Sign-SetupFile.ps1'
        $command = '"' + $env:WINDIR + '\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $callback + '" -Path $f'
        $arguments += @('/DSignedBuild=1', ('/Svbai=' + $command))
    }
    & $compiler @arguments (Join-Path $PSScriptRoot 'VBAi.iss')
    if ($LASTEXITCODE -ne 0) { throw 'Installer compilation failed; no accepted package was produced.' }
    $installer = Join-Path $output 'VBAi-Setup-win-x64.exe'
    $signature = Get-AuthenticodeSignature -LiteralPath $installer
    if (-not $Unsigned -and ($signature.Status -ne 'Valid' -or $signature.SignatureType -ne 'Authenticode' -or -not $signature.TimeStamperCertificate)) { throw 'Signed setup verification failed.' }
    if ($Unsigned -and $signature.Status -ne 'NotSigned') { throw 'Unexpected signature in unsigned build.' }
    $hash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant()
    [IO.File]::WriteAllText((Join-Path $output 'SHA256SUMS.txt'), ($hash + '  VBAi-Setup-win-x64.exe' + "`r`n"), (New-Object Text.UTF8Encoding($false)))
    $receipt = @{Product='VBAi'; Version=$Version; Architecture='win-x64'; Sha256=$hash; SignatureStatus=[string]$signature.Status; SignedUninstaller=[bool](-not $Unsigned); NativeInstallation='NOT_RUN'; Tests='STOPPED_BY_MAINTAINER'}
    [IO.File]::WriteAllText((Join-Path $output 'build-receipt.json'), ($receipt | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))
    Write-Output ('Installer: ' + $installer)
} finally { $env:VBAI_SETUP_SIGNING_SCRIPT = $oldSigner }

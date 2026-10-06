param(
    [Parameter(Mandatory = $true)][string]$Version,
    [string]$OutputDirectory,
    [string]$InstallerPath,
    [string]$HelpOutputRoot
)
$ErrorActionPreference = 'Stop'
if ($Version -notmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(-[0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*)?$') { throw 'Use a semantic product version, for example 1.2.3 or 1.2.3-beta.1.' }
[void][Version]::Parse($Version.Split('-')[0])
if ($Version.Contains('-')) { foreach ($identifier in $Version.Substring($Version.IndexOf('-') + 1).Split('.')) { if ($identifier -match '^0[0-9]+$') { throw 'Numeric preview identifiers must not have leading zeros.' } } }
$repository = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repository "dist\releases\$Version" }
if (-not $HelpOutputRoot) { $HelpOutputRoot = Join-Path $repository 'dist\help' }
$helpRoot = [IO.Path]::GetFullPath($HelpOutputRoot).TrimEnd('\') + '\'
$helpCultures = Get-ChildItem -LiteralPath (Join-Path $repository 'docs\help') -Directory | Where-Object { $_.Name -match '^[a-z]{2}-[A-Z]{2}$' }
foreach ($culture in $helpCultures) {
    $archive = Join-Path $helpRoot ('VBAi.' + $culture.Name + '.chm')
    if (-not (Test-Path -LiteralPath $archive -PathType Leaf)) { throw ('Build all localized CHM guides before preparing the payload: ' + $archive) }
}
$output = [IO.Path]::GetFullPath($OutputDirectory)
if (-not $output.StartsWith($repository + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'The release output must stay inside this checkout.' }
if (Test-Path -LiteralPath $output) { throw 'Release output already exists. Choose another directory; preserve the previous package.' }
$buildRoot = Join-Path $repository "artifacts\release-build\$Version"
if (Test-Path -LiteralPath $buildRoot) { throw 'Release build staging already exists. Choose a new version or clean that staging explicitly.' }
& dotnet build (Join-Path $repository 'src\VBAi\VBAi.csproj') -c Release --nologo "-p:ProductVersion=$Version" "-p:BuildOutputRoot=$buildRoot" "-p:HelpOutputRoot=$helpRoot" -v:q
if ($LASTEXITCODE -ne 0) { throw 'Release build failed.' }
$binary = Join-Path $buildRoot 'VBAi\Release\net48'
$assembly = Join-Path $binary 'VBAi.dll'
if ([Diagnostics.FileVersionInfo]::GetVersionInfo($assembly).ProductVersion -ne $Version) { throw 'Product version does not match the release.' }
$typeExporter = Get-Command TlbExp.exe -ErrorAction SilentlyContinue
$exportPath = if ($typeExporter) { $typeExporter.Source } else { Join-Path ${env:ProgramFiles(x86)} 'Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\TlbExp.exe' }
if (-not (Test-Path -LiteralPath $exportPath)) { throw 'The .NET Framework SDK TlbExp.exe is required to prepare the COM package.' }
& $exportPath $assembly "/out:$(Join-Path $binary 'VBAi.tlb')" /nologo
if ($LASTEXITCODE -ne 0) { throw 'Type library export failed.' }
$package = Join-Path $output 'package'
New-Item -ItemType Directory -Path $package | Out-Null
Get-ChildItem -LiteralPath $binary -File | Where-Object { $_.Extension -in '.dll', '.exe', '.config', '.tlb' } | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $package }
foreach ($payloadDirectory in @('EditorAssets', 'runtimes', 'ThirdPartyNotices', 'Licenses')) {
    $sourceDirectory = Join-Path $binary $payloadDirectory
    if (Test-Path -LiteralPath $sourceDirectory) { Copy-Item -LiteralPath $sourceDirectory -Destination $package -Recurse }
}
# The compiler staging, screenshots and logs are never installer payload.
$helpSource = Join-Path $binary 'Help'
if (Test-Path -LiteralPath $helpSource) {
    $helpDestination = Join-Path $package 'Help'
    New-Item -ItemType Directory -Path $helpDestination | Out-Null
    Get-ChildItem -LiteralPath $helpSource -File -Filter 'VBAi.*.chm' | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $helpDestination
    }
}
$marker = @{ Product = 'VBAi'; Architecture = 'win-x64'; UpdateProtocol = 1; InstallationId = [Guid]::NewGuid().ToString(); Version = $Version }
[IO.File]::WriteAllText((Join-Path $package 'vbai-installation.json'), ($marker | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))
# The setup author must preserve InstallationId on upgrades, register COM, and honor the documented update protocol.
if ($InstallerPath) {
    $installer = (Resolve-Path -LiteralPath $InstallerPath).Path
    $extension = [IO.Path]::GetExtension($installer).ToLowerInvariant()
    if ($extension -notin '.exe', '.msi') { throw 'Only an EXE or MSI installer is supported.' }
    $signature = Get-AuthenticodeSignature -LiteralPath $installer
    if ($signature.Status -ne 'Valid' -or $signature.SignatureType -ne 'Authenticode') { throw 'Sign the installer with a trusted Authenticode certificate before publishing.' }
    $releaseDirectory = Join-Path $output 'release-assets'
    New-Item -ItemType Directory -Path $releaseDirectory | Out-Null
    $asset = Join-Path $releaseDirectory "VBAi-Setup-win-x64$extension"
    Copy-Item -LiteralPath $installer -Destination $asset
    $hash = (Get-FileHash -LiteralPath $asset -Algorithm SHA256).Hash.ToLowerInvariant()
    [IO.File]::WriteAllText((Join-Path $releaseDirectory 'SHA256SUMS.txt'), "$hash  $([IO.Path]::GetFileName($asset))`r`n", (New-Object Text.UTF8Encoding($false)))
    Write-Output "Release asset: $asset"
}
Write-Output "Installer payload: $package"
Write-Output 'No release was published. Upload the signed installer to the matching GitHub release when it is ready.'

#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$Path)
$ErrorActionPreference = 'Stop'
# Optional local/HSM adapter. No certificate is generated or imported here.
$tool = $env:VBAI_SIGNTOOL_PATH
$thumbprint = $env:VBAI_SIGNING_CERTIFICATE_THUMBPRINT
$timestamp = $env:VBAI_SIGNING_TIMESTAMP_URL
if (-not $tool -or -not (Test-Path -LiteralPath $tool -PathType Leaf) -or $thumbprint -notmatch '^[0-9A-Fa-f]{40}$') { throw 'Configure SignTool and the exact code-signing certificate thumbprint.' }
if (-not $timestamp -or ([Uri]$timestamp).Scheme -notin 'https','http') { throw 'Configure the signing provider timestamp endpoint.' }
& $tool sign /fd SHA256 /sha1 $thumbprint /tr $timestamp /td SHA256 $Path
if ($LASTEXITCODE -ne 0) { throw 'Authenticode signing failed.' }
& $tool verify /pa /all $Path
if ($LASTEXITCODE -ne 0) { throw 'Authenticode verification failed.' }

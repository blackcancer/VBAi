#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$Path)
$ErrorActionPreference = 'Stop'
# Inno calls this for the uninstaller and then the outer setup. Secrets stay in
# the chosen provider's environment/key store, never in compiler arguments.
$signer = $env:VBAI_SETUP_SIGNING_SCRIPT
if (-not $signer -or -not (Test-Path -LiteralPath $signer -PathType Leaf)) { throw 'Configure a trusted provider signing script.' }
& $signer -Path (Resolve-Path -LiteralPath $Path).Path
if (-not $?) { throw 'The signing provider failed.' }
$signature = Get-AuthenticodeSignature -LiteralPath $Path
if ($signature.Status -ne 'Valid' -or $signature.SignatureType -ne 'Authenticode' -or -not $signature.TimeStamperCertificate) {
    throw 'A trusted, timestamped embedded Authenticode signature is required.'
}

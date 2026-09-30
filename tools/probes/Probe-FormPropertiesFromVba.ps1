# Legacy filename retained for callers; this probe only reads properties through the bridge.
# It never injects or executes VBA and never writes workbook contents.
param(
    [Parameter(Mandatory = $true)] [ValidateRange(1, 2147483647)] [int] $HostProcessId,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $Project,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $Form
)

$ErrorActionPreference = 'Stop'
$client = Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-VBAi.ps1'
$request = @{ Command = 'form_properties'; Project = $Project; Form = $Form }
$response = (& $client -HostProcessId $HostProcessId -RequestJson (ConvertTo-Json -InputObject $request -Compress)) | ConvertFrom-Json
if (-not $response.Ok) { throw $response.Error }
$response.Data | ConvertTo-Json -Depth 20

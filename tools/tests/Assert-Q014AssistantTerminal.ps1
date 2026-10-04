#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$EvidenceRoot)
$ErrorActionPreference = 'Stop'
if (-not [IO.Path]::IsPathRooted($EvidenceRoot) -or -not (Test-Path -LiteralPath $EvidenceRoot -PathType Container)) {
    throw 'An existing absolute Q014 evidence root is required.'
}
# A sent request can finish on the wire before its independent native readback.
# Host teardown must wait for that readback's terminal receipt, including failures.
foreach ($scenario in @('local-read','permission-refusal','revision-refusal','cancel','recovery')) {
    if (-not (Test-Path -LiteralPath (Join-Path $EvidenceRoot ($scenario + '-button-intent.json')))) { continue }
    $result = Join-Path $EvidenceRoot ($scenario + '-result.json')
    $error = Join-Path $EvidenceRoot ($scenario + '-error.json')
    if (-not (Test-Path -LiteralPath $result) -and -not (Test-Path -LiteralPath $error)) {
        throw ('Native assistant readback still pending: ' + $scenario + '; no host teardown allowed.')
    }
    if (Test-Path -LiteralPath $result) {
        $receipt = Get-Content -LiteralPath $result -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($receipt.State -cne 'PASS' -or $receipt.Scenario -cne $scenario) {
            throw ('Assistant success receipt differs: ' + $scenario)
        }
    } else {
        $receipt = Get-Content -LiteralPath $error -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([string]::IsNullOrWhiteSpace($receipt.Error)) { throw ('Assistant failure receipt incomplete: ' + $scenario) }
    }
}

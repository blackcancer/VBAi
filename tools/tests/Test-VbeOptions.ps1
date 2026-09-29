param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
if ((Get-Process -Id $HostProcessId -ErrorAction Stop).ProcessName -ne 'EXCEL') {
    throw 'An isolated Excel process is required.'
}

function Read-Options {
    $response = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId `
        -RequestJson '{"Command":"read_vbe_options"}' | ConvertFrom-Json
    if (-not $response.Ok) { throw $response.Error }
    return $response.Data
}

$first = Read-Options
$second = Read-Options
if ($first.Count -ne 4 -or -not $first.DialogClosed -or -not $second.DialogClosed) {
    throw 'The four native Options tabs were not read and closed.'
}
$firstNames = @($first.Tabs | ForEach-Object { $_.Tab })
$secondNames = @($second.Tabs | ForEach-Object { $_.Tab })
if (($firstNames -join '|') -ne ($secondNames -join '|')) {
    throw 'The tab names changed between two read-only observations.'
}
$firstValues = @($first.Tabs | ForEach-Object { $_.Controls } |
    Where-Object { $_.Type -eq 'ControlType.CheckBox' -or $_.Type -eq 'ControlType.RadioButton' } |
    ForEach-Object { "$($_.Name)=$($_.Value)" })
$secondValues = @($second.Tabs | ForEach-Object { $_.Controls } |
    Where-Object { $_.Type -eq 'ControlType.CheckBox' -or $_.Type -eq 'ControlType.RadioButton' } |
    ForEach-Object { "$($_.Name)=$($_.Value)" })
if (($firstValues -join '|') -ne ($secondValues -join '|')) {
    throw 'Native check or radio settings changed between two reads.'
}
[pscustomobject]@{ HostProcessId = $HostProcessId; Tabs = $firstNames -join ', ';
    SettingsCompared = $firstValues.Count; DialogClosed = $true;
    ValuesStable = $true } | Format-List

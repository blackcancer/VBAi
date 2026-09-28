param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$windows = @{}
$imports = @()
$calls = New-Object 'Collections.Generic.List[object]'
$session = $null
$summary = $null
Get-Content -LiteralPath $Path -Encoding UTF8 | ForEach-Object {
    $row = $_ | ConvertFrom-Json
    switch ($row.type) {
        'session' { $session = $row }
        'window' { $windows[$row.hwnd] = $row.class }
        'import' { $imports += $row }
        'summary' { $summary = $row }
        'call' {
            $row | Add-Member NoteProperty ContextClass ([string]$windows[$row.contextWindow])
            $row | Add-Member NoteProperty DirectClass ([string]$windows[$row.directWindow])
            $row | Add-Member NoteProperty LineageClass ([string]$windows[$row.lineageWindow])
            $calls.Add($row)
        }
    }
}
if (-not $session -or -not $summary) { throw 'The trace is missing its session or completion record.' }
$textCalls = @($calls | Where-Object { $_.api -match '^((Ext)?TextOut|DrawText)[AW]$' })
$report = [ordered]@{
    Session = $session
    Summary = $summary
    StartSucceeded = ($null -eq $session.startError -or $session.startError -eq 0)
    RestoredImportCount = @($imports | Where-Object restored).Count
    ImportCount = $imports.Count
    CallsByApi = @($calls | Group-Object api | Sort-Object Count -Descending | Select-Object Name,Count)
    TextByContextAndTarget = @($textCalls | Group-Object ContextClass,DirectClass,LineageClass,dcType,message | Sort-Object Count -Descending | Select-Object Name,Count)
    TextByCallSite = @($textCalls | Group-Object callerRva,ContextClass,DirectClass,dcType | Sort-Object Count -Descending | Select-Object Name,Count)
    TextWithoutWindowAttribution = @($textCalls | Where-Object { -not $_.ContextClass -and -not $_.DirectClass -and -not $_.LineageClass }).Count
    TextOutsidePaintOrPrint = @($textCalls | Where-Object { $_.message -notin @(15,43,49,133,791,792) }).Count
    Limits = @(
        'Only normal and already-resolved delay IAT imports of the selected module on the owner thread are observed.',
        'Memory DC lineage is inferred; context alone does not prove ownership.',
        'New windows created after Start are not added to this bounded snapshot.',
        'The trace does not observe every GDI, user32, Office or DirectWrite drawing path.',
        'Foreground/background values are COLORREF values, not RGB integers.'
    )
}
$output = [IO.Path]::ChangeExtension([IO.Path]::GetFullPath($Path), '.summary.json')
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $output -Encoding UTF8
$report | ConvertTo-Json -Depth 8

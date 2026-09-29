param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
if ((Get-Process -Id $HostProcessId -ErrorAction Stop).ProcessName -ne 'EXCEL') {
    throw 'An isolated Excel process is required.'
}

function Invoke-VbeRaw([hashtable] $Request) {
    $payload = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    return (& (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $payload |
        ConvertFrom-Json)
}

function Invoke-Vbe([hashtable] $Request) {
    $response = Invoke-VbeRaw $Request
    if (-not $response.Ok) { throw "$($Request.Command): $($response.Error)" }
    return $response.Data
}

$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2 -or $projects[0].FileName) {
    throw 'Use one unsaved disposable design-mode VBA project.'
}
$project = [string]$projects[0].Name
$environment = Invoke-Vbe @{ Command = 'vbe_environment' }
if (-not $environment.Properties.Version -or $environment.Properties.ProjectCount -ne 1 -or
    $environment.Properties.ActiveProject -ne $project) {
    throw 'VBE environment readback is incomplete.'
}
$addIns = Invoke-Vbe @{ Command = 'list_addins' }
$self = @($addIns.AddIns | Where-Object { $_.Properties.ProgId -eq 'VBAi.AddIn' })
if ($self.Count -ne 1 -or -not $self[0].Properties.Connect) {
    throw 'The connected VBE add-in was not read back.'
}
$windows = Invoke-Vbe @{ Command = 'vbe_windows' }
$projectWindow = @($windows.Windows | Where-Object {
    $_.Properties.Type -eq 6 -and $_.Properties.Visible
})
if ($projectWindow.Count -ne 1) { throw 'Expected one visible Project Explorer window.' }
$focus = Invoke-Vbe @{ Command = 'focus_vbe_window';
    WindowCaption = [string]$projectWindow[0].Properties.Caption; WindowType = 6 }
if ($focus.Verification -ne 'ActiveWindowReadback') { throw 'The window focus was not verified.' }
$hidden = @($windows.Windows | Where-Object { $_.Properties.Type -eq 2 -and -not $_.Properties.Visible })
if ($hidden.Count -eq 1) {
    $hiddenFocus = Invoke-VbeRaw @{ Command = 'focus_vbe_window';
        WindowCaption = [string]$hidden[0].Properties.Caption; WindowType = 2 }
    if ($hiddenFocus.Ok -or $hiddenFocus.Error -notmatch 'hidden') {
        throw 'A hidden window was not rejected.'
    }
}

$module = 'CodexSearchProbe'
$created = Invoke-Vbe @{ Command = 'create_module'; Project = $project; Module = $module; ExpectedMode = 2 }
$code = "Option Explicit`r`nPublic Sub PatternProbe()`r`n    Dim needle As String`r`n    needle = ""Needle""`r`n    Debug.Print needle, ""needless""`r`nEnd Sub"
Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $module;
    ExpectedSha256 = $created.Sha256; StartLine = 1; Count = $created.Lines; Text = $code } | Out-Null
$literal = Invoke-Vbe @{ Command = 'find_code'; Project = $project; Module = $module;
    Query = 'needle'; WholeWord = $true; MatchCase = $false }
$wildcard = Invoke-Vbe @{ Command = 'find_code'; Project = $project; Module = $module;
    Query = 'nee?le'; PatternSearch = $true; WholeWord = $true; MatchCase = $false }
$caseSensitive = Invoke-Vbe @{ Command = 'find_code'; Project = $project; Module = $module;
    Query = 'Nee*le'; PatternSearch = $true; WholeWord = $true; MatchCase = $true }
if (@($literal.Matches).Count -ne 4 -or @($wildcard.Matches).Count -ne 4 -or
    @($caseSensitive.Matches).Count -ne 1 -or $wildcard.Matches[0].Sha256 -ne $literal.Matches[0].Sha256) {
    throw 'Literal or wildcard search produced unexpected locations.'
}
$invalid = Invoke-VbeRaw @{ Command = 'find_code'; Project = $project; Module = $module;
    Query = '**'; PatternSearch = $true }
if ($invalid.Ok) { throw 'A pattern of only asterisks was accepted.' }
[pscustomobject]@{ HostProcessId = $HostProcessId; VbeVersion = $environment.Properties.Version;
    AddInCount = $addIns.Count; FocusVerified = $focus.Verification;
    LiteralMatches = @($literal.Matches).Count; WildcardMatches = @($wildcard.Matches).Count;
    CaseSensitiveMatches = @($caseSensitive.Matches).Count; InvalidPatternRejected = -not $invalid.Ok } | Format-List

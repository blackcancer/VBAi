param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
if ((Get-Process -Id $HostProcessId -ErrorAction Stop).ProcessName -ne 'EXCEL') {
    throw 'An isolated Excel process is required.'
}

function Invoke-VbeRaw([hashtable] $Request) {
    $payload = ConvertTo-Json -InputObject $Request -Compress -Depth 6
    return (& (Join-Path $PSScriptRoot '..\Invoke-CodexVBE.ps1') -HostProcessId $HostProcessId -RequestJson $payload |
        ConvertFrom-Json)
}

$windows = Invoke-VbeRaw @{ Command = 'vbe_windows' }
if (-not $windows.Ok) { throw $windows.Error }
$browser = @($windows.Data.Windows | Where-Object { $_.Properties.Type -eq 2 -and $_.Properties.Visible })
if ($browser.Count -ne 1) { throw 'Open exactly one visible Object Browser before this test.' }
$self = Invoke-VbeRaw @{ Command = 'close_vbe_window'; WindowCaption = 'CodexVBE'; WindowType = 15 }
if ($self.Ok -or $self.Error -notmatch 'cannot close itself') {
    throw 'The add-in did not refuse closing its own tool window.'
}
$caption = [string]$browser[0].Properties.Caption
$closed = Invoke-VbeRaw @{ Command = 'close_vbe_window'; WindowCaption = $caption; WindowType = 2 }
if (-not $closed.Ok -or $closed.Data.Verification -notin @('HiddenInWindows', 'RemovedFromWindows')) {
    throw "Object Browser close was not verified: $($closed.Error) $($closed.Data.Verification)"
}
$readback = Invoke-VbeRaw @{ Command = 'vbe_windows' }
$remaining = @($readback.Data.Windows | Where-Object {
    $_.Properties.Type -eq 2 -and $_.Properties.Caption -eq $caption -and $_.Properties.Visible
})
if (-not $readback.Ok -or $remaining.Count -ne 0) {
    throw 'The Object Browser is still visible after a separate readback.'
}
$reopened = $null
if ($closed.Data.Verification -eq 'HiddenInWindows') {
    $reopened = Invoke-VbeRaw @{ Command = 'show_vbe_window'; WindowCaption = $caption; WindowType = 2 }
    if (-not $reopened.Ok -or -not $reopened.Data.Visible -or
        -not $reopened.Data.FocusVerified) {
        throw "Object Browser reopen was not verified: $($reopened.Error)"
    }
    $shown = Invoke-VbeRaw @{ Command = 'vbe_windows' }
    $visibleAgain = @($shown.Data.Windows | Where-Object {
        $_.Properties.Type -eq 2 -and $_.Properties.Caption -eq $caption -and $_.Properties.Visible
    })
    if (-not $shown.Ok -or $visibleAgain.Count -ne 1) {
        throw 'A separate read did not confirm the Object Browser reopened.'
    }
}
$linkage = Invoke-VbeRaw @{ Command = 'window_linkage'; WindowCaption = $caption; WindowType = 2 }
if (-not $linkage.Ok -or @($linkage.Data.Properties.PSObject.Properties.Name) -notcontains 'IsLinked') {
    throw "The Object Browser linkage read failed: $($linkage.Error)"
}
[pscustomobject]@{ HostProcessId = $HostProcessId; CloseVerification = $closed.Data.Verification;
    SubsequentVisibleCount = $remaining.Count; Reopened = $reopened -ne $null;
    LinkageRead = $linkage.Data.Properties.IsLinked -ne $null;
    SelfCloseRejected = -not $self.Ok } | Format-List

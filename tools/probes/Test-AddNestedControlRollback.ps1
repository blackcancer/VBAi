param([Parameter(Mandatory = $true)] [int] $HostProcessId)

$ErrorActionPreference = 'Stop'
$hostProcess = Get-Process -Id $HostProcessId -ErrorAction Stop
if ($hostProcess.ProcessName -ne 'EXCEL') { throw 'The probe requires Excel.' }
function Invoke-Reply([hashtable] $Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json
}
function Invoke-Vbe([hashtable] $Request) {
    $reply = Invoke-Reply $Request
    if (-not $reply.Ok) { throw "$($Request.Command): $($reply.Error)" }
    return $reply.Data
}

$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) { throw 'Expected one disposable design-mode project.' }
$project = $projects[0].Name
$form = 'CodexNestedAtomicProbe'
if (@(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } | Where-Object { $_.Name -eq $form }).Count) {
    throw "Form $form already exists."
}
Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form } | Out-Null
$initial = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
    Control = 'fraProbe'; ControlType = 'Forms.Frame.1'; Caption = 'Parent';
    Left = 20; Top = 20; Width = 170; Height = 130;
    ExpectedFormVersion = $initial.FormVersion } | Out-Null
$before = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
$invalid = Invoke-Reply @{ Command = 'add_nested_form_control'; Project = $project; Form = $form;
    ParentPath = 'Controls/fraProbe'; Control = 'imgRejected'; ControlType = 'Forms.Image.1';
    Left = 10; Top = 20; Width = 60; Height = 50; Caption = 'Unsupported';
    ExpectedTreeVersion = $before.TreeVersion }
if ($invalid.Ok -or $invalid.Error -notmatch 'does not expose Caption') {
    throw "Nested Image.Caption did not fail as expected: $($invalid.Error)"
}
$afterRefusal = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
if ($afterRefusal.NodeCount -ne 1 -or $afterRefusal.TreeVersion -ne $before.TreeVersion -or
    @($afterRefusal.Controls[0].Children | Where-Object { $_.Name -eq 'imgRejected' }).Count) {
    throw 'The rejected nested Image changed form_tree or remains in its Frame.'
}
Invoke-Vbe @{ Command = 'add_nested_form_control'; Project = $project; Form = $form;
    ParentPath = 'Controls/fraProbe'; Control = 'imgAccepted'; ControlType = 'Forms.Image.1';
    Left = 10; Top = 20; Width = 60; Height = 50;
    ExpectedTreeVersion = $afterRefusal.TreeVersion } | Out-Null
$final = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
if ($final.NodeCount -ne 2 -or
    -not @($final.Controls[0].Children | Where-Object { $_.Name -eq 'imgAccepted' }).Count) {
    throw 'Nested Image without Caption was not created and read back.'
}
[pscustomobject]@{ HostProcessId = $HostProcessId; Form = $form;
    UnsupportedCaptionRefused = $true; RejectedControlAbsent = $true;
    TreeVersionRestored = $true; NestedImageCreated = $true;
    FinalNodeCount = $final.NodeCount;
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue) } | ConvertTo-Json

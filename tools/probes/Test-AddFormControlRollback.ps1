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
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) { throw 'Expected one disposable project in design mode.' }
$project = $projects[0].Name
$form = 'CodexAddAtomicProbe'
if (-not @(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } | Where-Object { $_.Name -eq $form }).Count) {
    Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form } | Out-Null
}
$initial = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
if ($initial.NodeCount -ne 0) { throw "Form $form must be empty before this probe." }
$invalid = Invoke-Reply @{ Command = 'add_form_control'; Project = $project; Form = $form;
    Control = 'imgRejected'; ControlType = 'Forms.Image.1'; Left = 24; Top = 24;
    Width = 90; Height = 60; Caption = 'Unsupported'; ExpectedFormVersion = $initial.FormVersion }
if ($invalid.Ok -or $invalid.Error -notmatch 'does not expose Caption') {
    throw "Image.Caption did not produce a verified rollback: $($invalid.Error)"
}
$afterRefusal = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
if ($afterRefusal.NodeCount -ne 0 -or
    $afterRefusal.TreeVersion -ne $initial.TreeVersion -or
    @($afterRefusal.Controls | Where-Object { $_.Name -eq 'imgRejected' }).Count) {
    throw 'The rejected Image changed form_tree or remains in the form.'
}
$image = Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
    Control = 'imgAccepted'; ControlType = 'Forms.Image.1'; Left = 24; Top = 24;
    Width = 90; Height = 60; ExpectedFormVersion = $afterRefusal.FormVersion }
$tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
$imageNode = @($tree.Controls | Where-Object { $_.Name -eq 'imgAccepted' })[0]
if (-not $imageNode -or $imageNode.Type -ne 'Image' -or $tree.NodeCount -ne 1) {
    throw 'Image without Caption was not created and read back.'
}
$label = Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
    Control = 'lblAccepted'; ControlType = 'Forms.Label.1'; Left = 24; Top = 100;
    Width = 90; Height = 30; Caption = 'Visible'; ExpectedFormVersion = $tree.FormVersion }
$final = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
$labelNode = @($final.Controls | Where-Object { $_.Name -eq 'lblAccepted' })[0]
$caption = @($labelNode.Properties | Where-Object { $_.Name -eq 'Caption' })[0].Value
if ($final.NodeCount -ne 2 -or $caption -ne 'Visible') {
    throw 'Label Caption did not survive the property preflight.'
}
[pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project; Form = $form;
    UnsupportedCaptionRefused = $true; RejectedControlAbsent = $true;
    ImageWithoutCaptionCreated = $true; LabelCaptionReadback = $caption;
    FinalNodeCount = $final.NodeCount;
    ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue) } | ConvertTo-Json -Depth 5

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
if ($projects.Count -ne 1 -or $projects[0].Mode -ne 2) {
    throw 'Expected exactly one disposable project in design mode.'
}
$project = $projects[0].Name
$form = 'CodexAutoInitializeProbe'
if (@(Invoke-Vbe @{ Command = 'list_forms'; Project = $project } | Where-Object { $_.Name -eq $form }).Count) {
    throw "Form $form already exists."
}
Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $form } | Out-Null
$tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $form;
    Control = 'cboProbe'; ControlType = 'Forms.ComboBox.1'; Left = 20; Top = 20;
    Width = 120; Height = 25; ExpectedFormVersion = $tree.FormVersion } | Out-Null
$tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $form }
$before = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $form }
if ($before.Code -match 'UserForm_Initialize') { throw 'Expected no existing Initialize procedure.' }
$request = @{ Command = 'set_form_list_initializer'; Project = $project; Form = $form;
    ControlPath = 'Controls/cboProbe'; Items = @('Alpha','Beta');
    ExpectedTreeVersion = $tree.TreeVersion; ExpectedSha256 = $before.Sha256 }
$first = Invoke-Vbe $request
$after = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $form }
if (-not $first.Applied -or -not $first.Verified -or $first.VerificationPending -or
    $after.Code -notmatch 'Private Sub UserForm_Initialize' -or
    $after.Code -notmatch 'Me.cboProbe.AddItem "Alpha"' -or
    $after.Code -notmatch 'Me.cboProbe.AddItem "Beta"') {
    throw 'Automatic Initialize creation or generated block failed verification.'
}
$request.ExpectedSha256 = $after.Sha256
$second = Invoke-Vbe $request
$again = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $form }
if ($second.Applied -or -not $second.Verified -or $again.Sha256 -ne $after.Sha256) {
    throw 'Repeating the same initializer was not idempotent.'
}
[pscustomobject]@{ HostProcessId = $HostProcessId; Project = $project; Form = $form;
    CreatedInitialize = $true; ItemsGenerated = 2; Idempotent = $true;
    CodeSha256 = $again.Sha256; ExcelAlive = [bool](Get-Process -Id $HostProcessId -ErrorAction SilentlyContinue) } |
    ConvertTo-Json -Depth 5

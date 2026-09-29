param(
    [Parameter(Mandatory = $true)] [int] $HostProcessId,
    [Parameter(Mandatory = $true)] [ValidateSet('Class', 'Form')] [string] $Kind
)

$ErrorActionPreference = 'Stop'
if ((Get-Process -Id $HostProcessId -ErrorAction Stop).ProcessName -ne 'SLDWORKS') {
    throw 'An already-open SOLIDWORKS process is required.'
}

function Invoke-Vbe([hashtable] $Request) {
    $payload = ConvertTo-Json -InputObject $Request -Compress -Depth 8
    $response = & (Join-Path $PSScriptRoot '..\Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $payload |
        ConvertFrom-Json
    if (-not $response.Ok) { throw "$($Request.Command): $($response.Error)" }
    return $response.Data
}

$project = 'test'
$projectFile = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\test.swp'))
$projects = @(Invoke-Vbe @{ Command = 'list_projects' })
if (@($projects | Where-Object { $_.Name -eq $project -and $_.FileName -eq $projectFile -and $_.Mode -eq 2 }).Count -ne 1) {
    throw 'The expected disposable SOLIDWORKS VBA project is unavailable.'
}
$name = if ($Kind -eq 'Class') { 'CodexSwClassRoundTrip' } else { 'CodexSwFormRoundTrip' }
$extension = if ($Kind -eq 'Class') { '.cls' } else { '.frm' }
$path = Join-Path $env:TEMP ("VBAi-sw-{0}-{1}{2}" -f $Kind.ToLowerInvariant(), $HostProcessId, $extension)
$frx = [IO.Path]::ChangeExtension($path, '.frx')
if ((Test-Path -LiteralPath $path) -or (Test-Path -LiteralPath $frx)) {
    throw 'A disposable export path already exists.'
}
$baseline = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
if (@($baseline | Where-Object { $_.Name -eq $name }).Count -ne 0) {
    throw 'The disposable component name already exists.'
}
$original = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
$verified = $false
try {
    if ($Kind -eq 'Class') {
        $created = Invoke-Vbe @{ Command = 'create_class'; Project = $project;
            Module = $name; ExpectedMode = 2 }
        $code = "Option Explicit`r`nPublic Function Value() As Long`r`n    Value = 42`r`nEnd Function"
        Invoke-Vbe @{ Command = 'replace_lines'; Project = $project; Module = $name;
            ExpectedSha256 = $created.Sha256; StartLine = 1; Count = $created.Lines; Text = $code } | Out-Null
    }
    else {
        Invoke-Vbe @{ Command = 'create_form'; Project = $project; Form = $name } | Out-Null
        $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $name }
        Invoke-Vbe @{ Command = 'add_form_control'; Project = $project; Form = $name;
            ExpectedFormVersion = $tree.FormVersion; ControlType = 'Forms.Label.1';
            Control = 'lblProbe'; Left = 12; Top = 12; Width = 80; Height = 20;
            Caption = 'SolidWorks probe' } | Out-Null
        $state = Invoke-Vbe @{ Command = 'form_state'; Project = $project; Form = $name }
        Invoke-Vbe @{ Command = 'set_form_property'; Project = $project; Form = $name;
            ExpectedFormVersion = $state.Version; Property = 'Caption'; Value = 'VBAi form probe' } | Out-Null
        $state = Invoke-Vbe @{ Command = 'form_state'; Project = $project; Form = $name }
        Invoke-Vbe @{ Command = 'set_form_property'; Project = $project; Form = $name;
            ExpectedFormVersion = $state.Version; Property = 'Width'; Value = 240 } | Out-Null
        foreach ($entry in @(
            @{ Property = 'Caption'; Value = 'Label from VBAi' },
            @{ Property = 'Font.Name'; Value = 'Arial' },
            @{ Property = 'Font.Size'; Value = 14 }
        )) {
            $tree = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $name }
            Invoke-Vbe @{ Command = 'set_form_node_property'; Project = $project; Form = $name;
                ControlPath = 'Controls/lblProbe'; ExpectedTreeVersion = $tree.TreeVersion;
                Property = $entry.Property; Value = $entry.Value } | Out-Null
        }
        $state = Invoke-Vbe @{ Command = 'form_state'; Project = $project; Form = $name }
        $label = @($state.Controls | Where-Object { $_.Name -eq 'lblProbe' }) | Select-Object -First 1
        if ($state.Caption -ne 'VBAi form probe' -or [Math]::Abs($state.Width - 240) -gt 0.1 -or
            $label.Caption -ne 'Label from VBAi' -or $label.FontName -ne 'Arial' -or
            [Math]::Abs($label.FontSize - 14) -gt 0.1) {
            throw 'SOLIDWORKS did not retain the form/label designer properties.'
        }
    }
    $codeBefore = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $name }
    $treeBefore = if ($Kind -eq 'Form') { Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $name } } else { $null }
    $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $name }
    $exported = Invoke-Vbe @{ Command = 'export_component'; Project = $project;
        Module = $name; Path = $path; ExpectedComponentVersion = $component.Version }
    if (-not (Test-Path -LiteralPath $path) -or $exported.Bytes -lt 1) {
        throw 'The exported component file is missing or empty.'
    }
    $projectState = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
    $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $name }
    $removed = Invoke-Vbe @{ Command = 'remove_component'; Project = $project; Module = $name;
        ExpectedProjectVersion = $projectState.Version; ExpectedComponentVersion = $component.Version }
    $imported = Invoke-Vbe @{ Command = 'import_component'; Project = $project;
        Path = $path; ExpectedProjectVersion = $removed.Version }
    if (-not $imported.Applied -or $imported.ImportedName -ne $name) {
        throw 'The imported component was not identified.'
    }
    $codeAfter = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = $name }
    if ($codeAfter.Sha256 -ne $codeBefore.Sha256) {
        throw 'The code SHA changed during export/import.'
    }
    if ($Kind -eq 'Form') {
        $treeAfter = Invoke-Vbe @{ Command = 'form_tree'; Project = $project; Form = $name }
        $stateAfter = Invoke-Vbe @{ Command = 'form_state'; Project = $project; Form = $name }
        $labelAfter = @($stateAfter.Controls | Where-Object { $_.Name -eq 'lblProbe' }) | Select-Object -First 1
        if ($treeAfter.TreeVersion -ne $treeBefore.TreeVersion -or
            @($treeAfter.Controls | Where-Object { $_.Name -eq 'lblProbe' }).Count -ne 1 -or
            $stateAfter.Caption -ne 'VBAi form probe' -or [Math]::Abs($stateAfter.Width - 240) -gt 0.1 -or
            $labelAfter.Caption -ne 'Label from VBAi' -or $labelAfter.FontName -ne 'Arial' -or
            [Math]::Abs($labelAfter.FontSize - 14) -gt 0.1) {
            throw 'The UserForm designer tree changed during export/import.'
        }
    }
    $verified = $true
}
finally {
    $current = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
    if (@($current | Where-Object { $_.Name -eq $name }).Count -eq 1) {
        $projectState = Invoke-Vbe @{ Command = 'project_properties'; Project = $project }
        $component = Invoke-Vbe @{ Command = 'component_properties'; Project = $project; Module = $name }
        Invoke-Vbe @{ Command = 'remove_component'; Project = $project; Module = $name;
            ExpectedProjectVersion = $projectState.Version; ExpectedComponentVersion = $component.Version } | Out-Null
    }
    $last = @(Invoke-Vbe @{ Command = 'list_modules'; Project = $project })
    $originalAfter = Invoke-Vbe @{ Command = 'read_module'; Project = $project; Module = 'test1' }
    if ($last.Count -ne $baseline.Count -or $originalAfter.Sha256 -ne $original.Sha256) {
        throw 'The original SOLIDWORKS VBA project inventory was not restored.'
    }
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    if (Test-Path -LiteralPath $frx) { Remove-Item -LiteralPath $frx -Force }
}
[pscustomobject]@{ HostProcessId = $HostProcessId; Kind = $Kind;
    RoundTripVerified = $verified; ModuleInventoryRestored = $true;
    OriginalModuleShaPreserved = $true } | Format-List

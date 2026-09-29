param(
    [Parameter(Mandatory = $true)][string]$AssemblyPath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [switch]$AllowTemporaryVbaAccess
)
$ErrorActionPreference = 'Stop'
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close existing Excel instances before this isolated test.' }
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$directory = [IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($directory) | Out-Null
$document = Join-Path $directory ('Vbai-EditorProbe-' + [Guid]::NewGuid().ToString('N') + '.xlsm')
$securityPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$prior = Get-ItemProperty -LiteralPath $securityPath -ErrorAction Stop
$hadAccess = $null -ne $prior.PSObject.Properties['AccessVBOM']
$priorAccess = $prior.AccessVBOM
if ($priorAccess -ne 1 -and -not $AllowTemporaryVbaAccess) { throw 'Explicit -AllowTemporaryVbaAccess is required to temporarily enable trusted VBA access.' }
$probeProcess = $null
$otherBook = $null
$excel = $null; $book = $null; $form = $null; $session = $null
function Invoke-Session([hashtable]$Fields) {
    $request = New-Object VBAi.Request
    foreach ($key in $Fields.Keys) { $request.$key = $Fields[$key] }
    $response = $script:session.Execute($request)
    if (-not $response.Ok) { throw $response.Error }
    return $response.Data
}
try {
    if ($priorAccess -ne 1) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null }
    $excel = New-Object -ComObject Excel.Application
    $probeProcess = Get-Process EXCEL -ErrorAction Stop
    if (@($probeProcess).Count -ne 1) { throw "Excel isolation was lost." }
    $excel.Visible = $true
    $book = $excel.Workbooks.Add()
    $vbe = $excel.GetType().InvokeMember('VBE', [Reflection.BindingFlags]::GetProperty, $null, $excel, $null)
    $vbe.MainWindow.Visible = $true
    $session = [Activator]::CreateInstance($assembly.GetType('VBAi.VbeSession', $true), [object[]]@($vbe))
    $project = $book.VBProject
    $actions = @('align_left','align_right','align_top','align_bottom','align_centers','align_middles','same_width','same_height','same_size','center_horizontal','center_vertical','snap_grid','distribute_horizontal','distribute_vertical','space_horizontal','space_vertical','increase_horizontal_spacing','increase_vertical_spacing','decrease_horizontal_spacing','decrease_vertical_spacing')
    $proof = @(); $index = 0
    function Read-Boxes($Container) {
        return @(0..2 | ForEach-Object { $c = $Container.Controls.Item("Button$_"); [pscustomobject]@{ Left = [double]$c.Left; Top = [double]$c.Top; Width = [double]$c.Width; Height = [double]$c.Height } })
    }
    function Assert-Near([double]$Actual, [double]$Expected, [string]$Name) {
        if ([Math]::Abs($Actual - $Expected) -gt 0.1) { throw "$Name : expected $Expected, got $Actual" }
    }
    foreach ($kind in @('form', 'frame', 'page')) {
        foreach ($action in $actions) {
            $form = $project.VBComponents.Add(3); $form.Name = ('LayoutProbe' + $index); $index++
            $form.Properties.Item('Width').Value = 500; $form.Properties.Item('Height').Value = 400
            $container = $form.Designer; $prefix = 'Controls/'
            if ($kind -eq 'frame') {
                $container = $form.Designer.Controls.Add('Forms.Frame.1','Frame1',$true)
                $container.Width = 400; $container.Height = 300
                $prefix = 'Controls/Frame1/Controls/'
            } elseif ($kind -eq 'page') {
                $multi = $form.Designer.Controls.Add('Forms.MultiPage.1','MultiPage1',$true)
                $multi.Width = 400; $multi.Height = 300
                $container = $multi.Pages.Item(0); $container.Name = 'Page1'
                $prefix = 'Controls/MultiPage1/Pages/Page1/Controls/'
            }
            $positions = @(@(53,53,60,30), @(143,113,40,20), @(233,203,50,25))
            foreach ($i in 0..2) {
                $control = $container.Controls.Add('Forms.CommandButton.1',"Button$i",$true)
                $control.Left = $positions[$i][0]; $control.Top = $positions[$i][1]
                $control.Width = $positions[$i][2]; $control.Height = $positions[$i][3]
            }
            $tree = Invoke-Session @{ Command = 'form_tree'; Project = $project.Name; Form = $form.Name }
            $paths = [string[]]@(0..2 | ForEach-Object { $prefix + "Button$_" })
            $result = Invoke-Session @{ Command = 'apply_form_layout'; Project = $project.Name; Form = $form.Name; ExpectedTreeVersion = $tree.TreeVersion; Items = $paths; Action = $action; Width = 10 }
            if (-not $result.Verified) { throw "Unverified layout $kind/$action" }
            $boxes = Read-Boxes $container
            foreach ($i in 0..2) {
                $box = $boxes[$i]
                switch ($action) {
                    'align_left' { Assert-Near $box.Left 53 $action }
                    'align_right' { Assert-Near ($box.Left + $box.Width) 113 $action }
                    'align_top' { Assert-Near $box.Top 53 $action }
                    'align_bottom' { Assert-Near ($box.Top + $box.Height) 83 $action }
                    'align_centers' { Assert-Near ($box.Left + $box.Width / 2) 83 $action }
                    'align_middles' { Assert-Near ($box.Top + $box.Height / 2) 68 $action }
                    'same_width' { Assert-Near $box.Width 60 $action }
                    'same_height' { Assert-Near $box.Height 30 $action }
                    'same_size' { Assert-Near $box.Width 60 $action; Assert-Near $box.Height 30 $action }
                    'snap_grid' { Assert-Near $box.Left (@(50,140,230)[$i]) $action; Assert-Near $box.Top (@(50,110,200)[$i]) $action }
                }
            }
            if ($action -eq 'center_horizontal') { Assert-Near (($boxes[0].Left + $boxes[2].Left + $boxes[2].Width)/2) ($container.InsideWidth/2) $action }
            if ($action -eq 'center_vertical') { Assert-Near (($boxes[0].Top + $boxes[2].Top + $boxes[2].Height)/2) ($container.InsideHeight/2) $action }
            if ($action -eq 'distribute_horizontal') { Assert-Near ($boxes[1].Left-$boxes[0].Left-$boxes[0].Width) 40 $action; Assert-Near ($boxes[2].Left-$boxes[1].Left-$boxes[1].Width) 40 $action }
            if ($action -eq 'distribute_vertical') { Assert-Near ($boxes[1].Top-$boxes[0].Top-$boxes[0].Height) 50 $action; Assert-Near ($boxes[2].Top-$boxes[1].Top-$boxes[1].Height) 50 $action }
            if ($action -match '^(space_|increase_|decrease_)') {
                $horizontal = $action.Contains('horizontal')
                $originalGaps = if ($horizontal) { @(30,50) } else { @(30,70) }
                foreach ($i in 1..2) {
                    $expectedGap = if ($action.StartsWith('space_')) { 10 } elseif ($action.StartsWith('increase_')) { $originalGaps[$i-1] + 10 } else { $originalGaps[$i-1] - 10 }
                    $actualGap = if ($horizontal) { $boxes[$i].Left - $boxes[$i-1].Left - $boxes[$i-1].Width } else { $boxes[$i].Top - $boxes[$i-1].Top - $boxes[$i-1].Height }
                    Assert-Near $actualGap $expectedGap $action
                }
            }
            $proof += [pscustomobject]@{ Form = $form.Name; Container = $kind; Action = $action; BeforeSave = $boxes; AfterReopen = $null; Verified = $true; PersistenceVerified = $false }
        }
    }
    $book.SaveAs($document, 52); $book.Close($false); $book = $null
    $book = $excel.Workbooks.Open($document); $project = $book.VBProject
    foreach ($row in $proof) {
        $form = $project.VBComponents.Item($row.Form); $container = $form.Designer
        if ($row.Container -eq 'frame') { $container = $container.Controls.Item('Frame1') }
        if ($row.Container -eq 'page') { $container = $container.Controls.Item('MultiPage1').Pages.Item('Page1') }
        $after = Read-Boxes $container
        foreach ($i in 0..2) { foreach ($property in @('Left','Top','Width','Height')) { Assert-Near $after[$i].$property $row.BeforeSave[$i].$property ($row.Form + ':' + $property) } }
        $row.AfterReopen = $after; $row.PersistenceVerified = $true
    }
    [pscustomobject]@{ Document = $document; Cases = $proof; Count = $proof.Count; Invocation = 'Current VbeSession against native Excel COM' } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath ($document + '.proof.json') -Encoding UTF8
    Write-Output "Verified $($proof.Count) native layouts after save/reopen. Evidence: $document.proof.json"
} finally {
    # Restore trust before Quit, which can block in some COM teardown paths.
    if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
    else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
    if ($null -ne $otherBook) { $otherBook.Close($false) }
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) {
        # This process was created after the no-Excel precondition and contains only
        # the disposable workbook. Quit can leave it alive because of COM references.
        Stop-Process -Id $probeProcess.Id -Force
    }
}

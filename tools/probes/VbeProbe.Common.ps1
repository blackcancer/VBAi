# Shared transport and identical owned-Excel setup. Dot source fixture blocks so
# their variables stay in the selected scenario's script scope. Cleanup stays in
# each scenario to preserve its restoration order, receipts and failure contract.
function Assert-VbeProbeParameters {
    param([hashtable]$Bound, [string[]]$Required, [string[]]$Allowed)
    foreach ($name in $Required) {
        if (-not $Bound.ContainsKey($name) -or $null -eq $Bound[$name] -or [string]::IsNullOrWhiteSpace([string]$Bound[$name])) {
            throw "Scenario '$Scenario' requires -$name."
        }
        if ($Bound[$name] -is [Management.Automation.SwitchParameter] -and -not $Bound[$name]) {
            throw "Scenario '$Scenario' requires explicit -$name."
        }
    }
    foreach ($name in $Bound.Keys) {
        if ($name -ne 'Scenario' -and $name -notin $Allowed -and $name -notin @('Verbose','Debug','ErrorAction','WarningAction','InformationAction','ErrorVariable','WarningVariable','InformationVariable','OutVariable','OutBuffer','PipelineVariable')) {
            throw "Parameter -$name does not apply to scenario '$Scenario'."
        }
    }
}
function Invoke-VbeRaw([hashtable]$Request) {
    $json = ConvertTo-Json -InputObject $Request -Compress -Depth $VbeProbeJsonDepth
    return (& (Join-Path $PSScriptRoot '../Invoke-VBAi.ps1') -HostProcessId $HostProcessId -RequestJson $json | ConvertFrom-Json)
}
function Invoke-Reply([hashtable]$Request) { Invoke-VbeRaw $Request }
function Invoke-Vbe([hashtable]$Request) {
    $response = Invoke-VbeRaw $Request
    if (-not $response.Ok) {
        if ($VbeProbeRawError) { throw $response.Error }
        throw "$($Request.Command): $($response.Error)"
    }
    return $response.Data
}
function Invoke-Session([hashtable]$Fields) {
    $request = New-Object VBAi.Request
    foreach ($key in $Fields.Keys) { $request.$key = $Fields[$key] }
    if ($VbeProbeTraceRequests) { Write-Output ("Request: " + $request.Command + " " + $request.Action) | Out-Host }
    $response = $script:session.Execute($request)
    if (-not $response.Ok) { throw $response.Error }
    return $response.Data
}

$VbeProbeInitializeAssembly = {
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
}

$VbeProbeOpenAssemblyExcel = {
    if ($priorAccess -ne 1) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null }
    $excel = New-Object -ComObject Excel.Application
    $probeProcess = Get-Process EXCEL -ErrorAction Stop
    if (@($probeProcess).Count -ne 1) { throw "Excel isolation was lost." }
    $excel.Visible = $true
    $book = $excel.Workbooks.Add()
    $vbe = $excel.GetType().InvokeMember('VBE', [Reflection.BindingFlags]::GetProperty, $null, $excel, $null)
}

$VbeProbeInitializeRegistered = {
    if (-not $UseBridge -or @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Registered bridge and an isolated Excel session are required.' }
    $assembly=[Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
    $outputRoot=[IO.Path]::GetFullPath($OutputDirectory);[IO.Directory]::CreateDirectory($outputRoot) | Out-Null
    $securityPath='HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    $initial=Get-ItemProperty -LiteralPath $securityPath
    $hadAccess=$null -ne $initial.PSObject.Properties['AccessVBOM'];$initialAccess=$initial.AccessVBOM
    @{HadAccessVBOM=$hadAccess;AccessVBOM=$initialAccess} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $outputRoot 'security-before.json') -Encoding UTF8
    if($initialAccess -ne 1 -and -not $AllowTemporaryVbaAccess){throw 'Explicit temporary AccessVBOM opt-in is required.'}
}

function Invoke-Bridge([hashtable]$fields){
    $response=& (Join-Path $PSScriptRoot '../Invoke-VBAi.ps1') -HostProcessId $script:probeProcess.Id -RequestJson ($fields|ConvertTo-Json -Compress -Depth 12) -ResponseTimeoutSeconds 30 | ConvertFrom-Json
    if(-not $response.Ok){throw "$($fields.Command): $($response.Error)"};return $response.Data
}

function Open-ProbeExcel{
    if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Excel isolation was lost before launch.'}
    $scratch=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'VBAi-scratch.xlsx'))
    $script:probeProcess=Start-Process -FilePath 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE' -ArgumentList @('/x',('"'+$scratch+'"')) -WindowStyle Hidden -PassThru
    $script:excel=$null
    for($i=0;$i -lt 150 -and $null -eq $script:excel;$i++){Start-Sleep -Milliseconds 200;try{$script:excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')}catch{}}
    if($null -eq $script:excel -or @(Get-Process EXCEL).Count -ne 1 -or (Get-Process EXCEL).Id -ne $script:probeProcess.Id){throw 'The owned Excel process is not uniquely attached.'}
    if($script:excel.Workbooks.Count -ne 1 -or $script:excel.Workbooks.Item(1).FullName -ne $scratch){throw 'Unexpected workbook in the isolated process.'}
    $script:excel.Visible=$true;$script:vbe=$script:excel.GetType().InvokeMember('VBE',[Reflection.BindingFlags]::GetProperty,$null,$script:excel,$null);$script:vbe.MainWindow.Visible=$true
    $status=Invoke-Bridge @{Command='status'}
    if($status.HostProcessId -ne $script:probeProcess.Id -or $status.AssemblyModuleVersionId -ne $assembly.ManifestModule.ModuleVersionId.ToString('D')){throw 'Bridge process/assembly mismatch.'}
    $script:excel.Workbooks.Item(1).Close($false);return $status
}

function Close-ProbeExcel{
    if($null -ne $script:book){$script:book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:book);$script:book=$null}
    if($null -ne $script:excel){$script:excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:excel);$script:excel=$null;$script:vbe=$null}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null -ne $script:probeProcess -and -not $script:probeProcess.WaitForExit(5000)){throw 'The owned Excel process has not exited; no next launch is allowed.'}
}

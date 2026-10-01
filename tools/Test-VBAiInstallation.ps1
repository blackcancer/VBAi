param([switch] $Direct)

$ErrorActionPreference = 'Stop'

if ($env:CODEX_SHELL -eq '1' -and -not $Direct) {
    $expectedAssembly = Join-Path (Split-Path -Parent $PSScriptRoot) 'bin\Debug\net48\VBAi.dll'
    & (Join-Path $PSScriptRoot 'Invoke-VBAi-OutsideSandbox.ps1') -Action Verify -ExpectedAssemblyPath $expectedAssembly
    return
}

$classId = '{8E854243-087F-4D6C-9E0E-8622B0E50883}'
$typeLibId = '{AF3C2AF7-155F-4DDB-AC8F-D02CD58DEDC9}'
$progId = 'VBAi.AddIn'
$projectRoot = Split-Path -Parent $PSScriptRoot
$assemblyPath = Join-Path $projectRoot 'bin\Debug\net48\VBAi.dll'
$expectedCodeBase = 'file:///' + ([System.IO.Path]::GetFullPath($assemblyPath)).Replace([char]92, [char]47)
$typeLibPath = Join-Path $projectRoot 'bin\Debug\net48\VBAi.tlb'
$logPath = Join-Path $env:TEMP 'VBAi-load.log'
$expectedAssemblyName = if (Test-Path -LiteralPath $assemblyPath) {
    [Reflection.AssemblyName]::GetAssemblyName($assemblyPath).FullName
} else { $null }

Write-Output "64-bit PowerShell: $([Environment]::Is64BitProcess)"
Write-Output "Assembly exists: $(Test-Path -LiteralPath $assemblyPath)"
Write-Output "Assembly path: $assemblyPath"
Write-Output "Expected CodeBase: $expectedCodeBase"
Write-Output "Type library exists: $(Test-Path -LiteralPath $typeLibPath)"

$registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
    [Microsoft.Win32.RegistryHive]::CurrentUser,
    [Microsoft.Win32.RegistryView]::Registry64)
try {
    $registrationFailures = @()
    foreach ($entry in @(
        @{ Label = 'VBE add-in'; Path = "Software\Microsoft\VBA\VBE\6.0\Addins64\$progId"; Values = @{ LoadBehavior = 3 } },
        @{ Label = 'COM ProgID'; Path = "Software\Classes\$progId\CLSID"; Values = @{ '' = $classId } },
        @{ Label = 'COM class'; Path = "Software\Classes\CLSID\$classId"; Values = @{ '' = $progId } },
        @{ Label = 'COM server'; Path = "Software\Classes\CLSID\$classId\InprocServer32"; Values = @{ '' = 'mscoree.dll'; ThreadingModel = 'Both'; Class = 'VBAi.AddIn'; RuntimeVersion = 'v4.0.30319'; CodeBase = $expectedCodeBase } },
        @{ Label = 'COM class type library'; Path = "Software\Classes\CLSID\$classId\TypeLib"; Values = @{ '' = $typeLibId } },
        @{ Label = 'Registered type library'; Path = "Software\Classes\TypeLib\$typeLibId\0.1\0\win64"; Values = @{ '' = $typeLibPath } },
        @{ Label = 'Test runtime ProgID'; Path = 'Software\Classes\VBAi.TestRuntime\CLSID'; Values = @{ '' = '{5AF2F40B-939B-4CC6-A06C-F0C79841C031}' } },
        @{ Label = 'Test runtime server'; Path = 'Software\Classes\CLSID\{5AF2F40B-939B-4CC6-A06C-F0C79841C031}\InprocServer32'; Values = @{ '' = 'mscoree.dll'; Class = 'VBAi.VbaTestRuntime'; Assembly = $expectedAssemblyName; RuntimeVersion = 'v4.0.30319'; ThreadingModel = 'Both'; CodeBase = $expectedCodeBase } }
    )) {
        $key = $registry.OpenSubKey($entry.Path)
        if ($null -eq $key) {
            Write-Output "$($entry.Label): MISSING"
            $registrationFailures += "$($entry.Label) is missing"
            continue
        }
        try {
            Write-Output "$($entry.Label): PRESENT"
            foreach ($name in $key.GetValueNames()) {
                $label = if ($name -eq '') { '(default)' } else { $name }
                Write-Output "  $label = $($key.GetValue($name))"
            }
            foreach ($name in $entry.Values.Keys) {
                $actual = $key.GetValue($name)
                $expected = $entry.Values[$name]
                $matches = if ($name -eq 'LoadBehavior') { $actual -eq $expected } else { $actual -ceq $expected }
                $label = if ($name -eq '') { '(default)' } else { $name }
                Write-Output "  $label matches expected: $matches"
                if (-not $matches) { $registrationFailures += "$($entry.Label) $label mismatch" }
            }
            if ($entry.Label -eq 'COM server' -and $expectedAssemblyName) {
                $matches = $key.GetValue('Assembly') -ceq $expectedAssemblyName
                Write-Output "  Assembly matches expected: $matches"
                if (-not $matches) { $registrationFailures += 'COM server Assembly mismatch' }
            }
        }
        finally { $key.Dispose() }
    }
}
finally { $registry.Dispose() }

if (-not [Environment]::Is64BitProcess) { $registrationFailures += 'PowerShell is not 64-bit' }
if (-not (Test-Path -LiteralPath $assemblyPath)) { $registrationFailures += 'Assembly is missing' }
if (-not (Test-Path -LiteralPath $typeLibPath)) { $registrationFailures += 'Type library is missing' }

$activationFailure = $null
$instance = $null
try {
    $type = [Type]::GetTypeFromProgID($progId, $true)
    $instance = [Activator]::CreateInstance($type)
    Write-Output 'COM activation: OK'
}
catch {
    $activationFailure = $_.Exception.Message
    Write-Output "COM activation: FAILED ($activationFailure)"
}
finally {
    if ($null -ne $instance -and [Runtime.InteropServices.Marshal]::IsComObject($instance)) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($instance) | Out-Null
    }
}

if (Test-Path -LiteralPath $logPath) {
    Write-Output "Load log: $logPath"
    Get-Content -LiteralPath $logPath -Tail 20
}
else {
    Write-Output "Load log: MISSING ($logPath)"
}

$result = [pscustomobject]@{
    Registration = if ($registrationFailures.Count) { 'FAILED' } else { 'OK' }
    ComActivation = if ($activationFailure) { 'FAILED' } else { 'OK' }
    OnConnection = 'NOT_TESTED'
    ChatMonaco = 'NOT_TESTED'
    RegistrationFailures = @($registrationFailures)
    ActivationFailure = $activationFailure
}
Write-Output ($result | ConvertTo-Json -Compress -Depth 3)
if ($registrationFailures.Count -or $activationFailure) {
    throw ("VBAi installation verification failed: " + ($result | ConvertTo-Json -Compress -Depth 3))
}

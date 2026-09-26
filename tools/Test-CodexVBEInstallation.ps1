$ErrorActionPreference = 'Stop'

if ($env:CODEX_SHELL -eq '1') {
    $expectedAssembly = Join-Path (Split-Path -Parent $PSScriptRoot) 'bin\Debug\net48\CodexVBE.dll'
    & (Join-Path $PSScriptRoot 'Invoke-CodexVBE-OutsideSandbox.ps1') -Action Verify -ExpectedAssemblyPath $expectedAssembly
    return
}

$classId = '{8E854243-087F-4D6C-9E0E-8622B0E50883}'
$typeLibId = '{AF3C2AF7-155F-4DDB-AC8F-D02CD58DEDC9}'
$progId = 'CodexVBE.AddIn'
$projectRoot = Split-Path -Parent $PSScriptRoot
$assemblyPath = Join-Path $projectRoot 'bin\Debug\net48\CodexVBE.dll'
$expectedCodeBase = 'file:///' + ([System.IO.Path]::GetFullPath($assemblyPath)).Replace([char]92, [char]47)
$typeLibPath = Join-Path $projectRoot 'bin\Debug\net48\CodexVBE.tlb'
$logPath = Join-Path $env:TEMP 'CodexVBE-load.log'

Write-Output "64-bit PowerShell: $([Environment]::Is64BitProcess)"
Write-Output "Assembly exists: $(Test-Path -LiteralPath $assemblyPath)"
Write-Output "Assembly path: $assemblyPath"
Write-Output "Expected CodeBase: $expectedCodeBase"
Write-Output "Type library exists: $(Test-Path -LiteralPath $typeLibPath)"

$registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
    [Microsoft.Win32.RegistryHive]::CurrentUser,
    [Microsoft.Win32.RegistryView]::Registry64)
try {
    foreach ($entry in @(
        @{ Label = 'VBE add-in'; Path = "Software\Microsoft\VBA\VBE\6.0\Addins64\$progId" },
        @{ Label = 'COM ProgID'; Path = "Software\Classes\$progId\CLSID" },
        @{ Label = 'COM server'; Path = "Software\Classes\CLSID\$classId\InprocServer32" },
        @{ Label = 'COM class type library'; Path = "Software\Classes\CLSID\$classId\TypeLib" },
        @{ Label = 'Registered type library'; Path = "Software\Classes\TypeLib\$typeLibId\0.1\0\win64" }
    )) {
        $key = $registry.OpenSubKey($entry.Path)
        if ($null -eq $key) {
            Write-Output "$($entry.Label): MISSING"
            continue
        }
        try {
            Write-Output "$($entry.Label): PRESENT"
            foreach ($name in $key.GetValueNames()) {
                $label = if ($name -eq '') { '(default)' } else { $name }
                Write-Output "  $label = $($key.GetValue($name))"
            }
            if ($entry.Label -eq 'COM server') {
                Write-Output "  CodeBase matches RegAsm format: $($key.GetValue('CodeBase') -ceq $expectedCodeBase)"
            }
            if ($entry.Label -eq 'COM class type library') {
                Write-Output "  Expected type library: $($key.GetValue('') -eq $typeLibId)"
            }
            if ($entry.Label -eq 'Registered type library') {
                Write-Output "  Type library path matches: $($key.GetValue('') -ceq $typeLibPath)"
            }
        }
        finally { $key.Dispose() }
    }
}
finally { $registry.Dispose() }

try {
    $type = [Type]::GetTypeFromProgID($progId, $true)
    $instance = [Activator]::CreateInstance($type)
    Write-Output 'COM activation: OK'
    if ([Runtime.InteropServices.Marshal]::IsComObject($instance)) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($instance) | Out-Null
    }
}
catch {
    Write-Output "COM activation: FAILED ($($_.Exception.ToString()))"
}

if (Test-Path -LiteralPath $logPath) {
    Write-Output "Load log: $logPath"
    Get-Content -LiteralPath $logPath -Tail 20
}
else {
    Write-Output "Load log: MISSING ($logPath)"
}

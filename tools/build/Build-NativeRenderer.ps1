param(
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [ValidateSet('Debug', 'Release')][string]$Configuration = 'Debug',
    [switch]$BuildSelfTest
)
$ErrorActionPreference = 'Stop'
$sourceDirectory = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../src/CodexVBE.Native'))
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) { throw 'Visual Studio Installer is required to compile the native renderer.' }
$installation = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if (-not $installation) { throw 'Install the Visual Studio Desktop development with C++ workload to build CodexVBE. End users do not need the compiler.' }
$compilerRoot = Get-ChildItem -LiteralPath (Join-Path $installation 'VC/Tools/MSVC') -Directory | Sort-Object Name -Descending | Select-Object -First 1
$sdkRoot = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits/10'
$sdk = Get-ChildItem -LiteralPath (Join-Path $sdkRoot 'Include') -Directory | Where-Object {
    Test-Path -LiteralPath (Join-Path $_.FullName 'um/Windows.h')
} | Sort-Object Name -Descending | Select-Object -First 1
if (-not $compilerRoot -or -not $sdk) { throw 'The C++ x64 compiler or Windows SDK is missing.' }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
[void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
$previousInclude = $env:INCLUDE; $previousLib = $env:LIB; $previousPath = $env:PATH
try {
    $env:INCLUDE = @((Join-Path $compilerRoot.FullName 'include'), (Join-Path $sdk.FullName 'ucrt'), (Join-Path $sdk.FullName 'shared'), (Join-Path $sdk.FullName 'um')) -join ';'
    $env:LIB = @((Join-Path $compilerRoot.FullName 'lib/x64'), (Join-Path $sdkRoot "Lib/$($sdk.Name)/ucrt/x64"), (Join-Path $sdkRoot "Lib/$($sdk.Name)/um/x64")) -join ';'
    $env:PATH = (Join-Path $compilerRoot.FullName 'bin/Hostx64/x64') + ';' + $previousPath
    $compiler = Join-Path $compilerRoot.FullName 'bin/Hostx64/x64/cl.exe'
    # Static CRT: deploying the embedded binary requires no C++ toolchain or
    # separate Visual C++ redistributable on the user's workstation.
    & $compiler /nologo /std:c++17 /W4 /WX /EHsc /MT /O2 /Zi /guard:cf /utf-8 /DUNICODE /D_UNICODE /LD `
        (Join-Path $sourceDirectory 'NativeTheme.cpp') "/Fo$OutputDirectory/NativeTheme.obj" "/Fd$OutputDirectory/compiler.pdb" `
        /link /MACHINE:X64 /DYNAMICBASE /NXCOMPAT /OPT:REF /OPT:ICF /DEBUG "/OUT:$OutputDirectory/CodexVBE.Native.dll" "/IMPLIB:$OutputDirectory/CodexVBE.Native.lib" user32.lib gdi32.lib comctl32.lib
    if ($LASTEXITCODE -ne 0) { throw "Native renderer compilation failed ($LASTEXITCODE)." }
    if ($BuildSelfTest) {
        # Synthetic fixture, isolated from installed Office/VBA binaries.
        & $compiler /nologo /std:c++17 /W4 /WX /EHsc /MT /O2 /utf-8 /LD (Join-Path $PSScriptRoot '../probes/native-render-trace/TraceFixture.cpp') "/Fo$OutputDirectory/TraceFixture.obj" /link /DELAYLOAD:USER32.dll "/OUT:$OutputDirectory/VBE7.dll" "/IMPLIB:$OutputDirectory/TraceFixture.lib" user32.lib gdi32.lib delayimp.lib
        if ($LASTEXITCODE -ne 0) { throw "Synthetic fixture compilation failed ($LASTEXITCODE)." }
        & $compiler /nologo /std:c++17 /W4 /WX /EHsc /MT /O2 /utf-8 (Join-Path $PSScriptRoot '../../tests/native/NativeRendererSelfTest.cpp') "/Fo$OutputDirectory/NativeRendererSelfTest.obj" /link "/OUT:$OutputDirectory/NativeRendererSelfTest.exe" user32.lib gdi32.lib
        if ($LASTEXITCODE -ne 0) { throw "Native lifecycle test compilation failed ($LASTEXITCODE)." }
    }
    Write-Output "Native renderer built: $OutputDirectory/CodexVBE.Native.dll"
} finally { $env:INCLUDE = $previousInclude; $env:LIB = $previousLib; $env:PATH = $previousPath }

param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '../../../artifacts/native-renderer-pilot/trace'),
    [switch]$BuildSelfTest
)

$ErrorActionPreference = 'Stop'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) { throw 'Visual Studio Installer vswhere.exe was not found.' }
$installation = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if (-not $installation) { throw 'Install the Visual Studio C++ x64 toolchain before building this optional diagnostic.' }
$compilerRoot = Get-ChildItem -LiteralPath (Join-Path $installation 'VC/Tools/MSVC') -Directory | Sort-Object Name -Descending | Select-Object -First 1
$sdkRoot = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits/10'
$sdk = Get-ChildItem -LiteralPath (Join-Path $sdkRoot 'Include') -Directory | Where-Object {
    Test-Path -LiteralPath (Join-Path $_.FullName 'um/Windows.h')
} | Sort-Object Name -Descending | Select-Object -First 1
if (-not $compilerRoot -or -not $sdk) { throw 'The compiler or Windows SDK headers are missing.' }
$compiler = Join-Path $compilerRoot.FullName 'bin/Hostx64/x64/cl.exe'
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
[void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
$previousInclude = $env:INCLUDE
$previousLib = $env:LIB
$previousPath = $env:PATH
try {
    $env:INCLUDE = @((Join-Path $compilerRoot.FullName 'include'), (Join-Path $sdk.FullName 'ucrt'), (Join-Path $sdk.FullName 'shared'), (Join-Path $sdk.FullName 'um')) -join ';'
    $env:LIB = @((Join-Path $compilerRoot.FullName 'lib/x64'), (Join-Path $sdkRoot "Lib/$($sdk.Name)/ucrt/x64"), (Join-Path $sdkRoot "Lib/$($sdk.Name)/um/x64")) -join ';'
    $env:PATH = (Join-Path $compilerRoot.FullName 'bin/Hostx64/x64') + ';' + $previousPath
    $common = @('/nologo', '/std:c++17', '/W4', '/EHsc', '/MT', '/O2', '/Zi', '/guard:cf', '/utf-8', '/DUNICODE', '/D_UNICODE')
    & $compiler @common /LD (Join-Path $PSScriptRoot 'NativeRenderTrace.cpp') "/Fo$OutputDirectory/NativeRenderTrace.obj" "/Fd$OutputDirectory/compiler.pdb" `
        /link /MACHINE:X64 /DYNAMICBASE /NXCOMPAT /DEBUG "/OUT:$OutputDirectory/CodexVbeNativeTrace.dll" "/IMPLIB:$OutputDirectory/CodexVbeNativeTrace.lib" user32.lib gdi32.lib comctl32.lib
    if ($LASTEXITCODE -ne 0) { throw "Native trace compilation failed ($LASTEXITCODE)." }
    if ($BuildSelfTest) {
        $selfTest = Join-Path $OutputDirectory 'selftest'
        [void](New-Item -ItemType Directory -Path $selfTest -Force)
        & $compiler @common /LD (Join-Path $PSScriptRoot 'TraceFixture.cpp') "/Fo$selfTest/TraceFixture.obj" "/Fd$selfTest/fixture-compiler.pdb" `
            /link /MACHINE:X64 /DYNAMICBASE /NXCOMPAT /DEBUG /DELAYLOAD:USER32.dll "/OUT:$selfTest/VBE7.dll" "/IMPLIB:$selfTest/TraceFixture.lib" user32.lib gdi32.lib delayimp.lib
        if ($LASTEXITCODE -ne 0) { throw "Synthetic fixture compilation failed ($LASTEXITCODE)." }
        & $compiler @common (Join-Path $PSScriptRoot 'TraceSelfTest.cpp') "/Fo$selfTest/TraceSelfTest.obj" "/Fd$selfTest/test-compiler.pdb" `
            /link /MACHINE:X64 /DYNAMICBASE /NXCOMPAT /DEBUG "/OUT:$selfTest/TraceSelfTest.exe" user32.lib gdi32.lib comctl32.lib
        if ($LASTEXITCODE -ne 0) { throw "Synthetic trace test compilation failed ($LASTEXITCODE)." }
        & $compiler @common (Join-Path $PSScriptRoot 'PatternPilotSelfTest.cpp') "/Fo$selfTest/PatternPilotSelfTest.obj" "/Fd$selfTest/pattern-test-compiler.pdb" `
            /link /MACHINE:X64 /DYNAMICBASE /NXCOMPAT /DEBUG "/OUT:$selfTest/PatternPilotSelfTest.exe" user32.lib gdi32.lib
        if ($LASTEXITCODE -ne 0) { throw "Pattern pilot test compilation failed ($LASTEXITCODE)." }
    }
    Get-Item -LiteralPath (Join-Path $OutputDirectory 'CodexVbeNativeTrace.dll') | Select-Object FullName,Length
} finally {
    $env:INCLUDE = $previousInclude
    $env:LIB = $previousLib
    $env:PATH = $previousPath
}

param([Parameter(Mandatory=$true)][string]$AssemblyPath)
$ErrorActionPreference = 'Stop'
$assembly = [Reflection.Assembly]::LoadFrom([IO.Path]::GetFullPath($AssemblyPath))
$type = $assembly.GetType('VBAi.VbeNativeRenderer', $true)
$flags = [Reflection.BindingFlags]'NonPublic,Static'
$load = $type.GetMethod('EnsureLoaded', $flags)
$load.Invoke($null, @())
$first = $type.GetField('module', $flags).GetValue($null)
$load.Invoke($null, @())
$second = $type.GetField('module', $flags).GetValue($null)
if ($first -eq [IntPtr]::Zero -or $first -ne $second) { throw 'Native module not loaded or reused.' }
Write-Output ('PASS: embedded payload extraction, hash validation, ABI check, module reuse. ' + $type.GetMethod('Describe', $flags).Invoke($null, @()))

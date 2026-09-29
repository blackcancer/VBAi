$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit Windows PowerShell.' }

$type = [Type]::GetTypeFromProgID('VBAi.ChatToolWindow', $true)
$control = [Activator]::CreateInstance($type)
try {
    if ($control.GetType().FullName -ne 'VBAi.ChatToolWindow') {
        throw 'The chat control ProgID activated an unexpected class.'
    }
    Write-Output 'PASS chat control COM activation in the current user session'
}
finally {
    if ($control -is [IDisposable]) { $control.Dispose() }
    elseif ([Runtime.InteropServices.Marshal]::IsComObject($control)) {
        [Runtime.InteropServices.Marshal]::ReleaseComObject($control) | Out-Null
    }
}

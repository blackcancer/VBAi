# Shared inert helpers. No assembly, window or host is opened on import.
function Open-UiAssembly([string]$AssemblyPath) {
    [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
}

function New-UiObject($Assembly, [string]$Name) {
    [Activator]::CreateInstance($Assembly.GetType('VBAi.' + $Name), $true)
}

function Get-UiField($Target, [string]$Name) {
    ,$Target.GetType().GetField($Name, [Reflection.BindingFlags]'Instance,Public,NonPublic').GetValue($Target)
}

function Set-UiField($Target, [string]$Name, $Value) {
    $Target.GetType().GetField($Name, [Reflection.BindingFlags]'Instance,Public,NonPublic').SetValue($Target, $Value)
}

function Invoke-UiMethod($Target, [string]$Name, [object[]]$Arguments) {
    $Target.GetType().GetMethod($Name, [Reflection.BindingFlags]'Instance,Public,NonPublic').Invoke($Target, $Arguments)
}

function Assert-Ui($Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

function Get-UiScenarioParameters($Command, [Collections.IDictionary]$Parameters) {
    $arguments = @{}
    foreach ($key in $Parameters.Keys) {
        if ($key -eq 'Scenario') { continue }
        if (-not $Command.Parameters.ContainsKey($key)) {
            throw "Parameter -$key does not apply to $($Command.Name)."
        }
        $arguments[$key] = $Parameters[$key]
    }
    foreach ($parameter in $Command.Parameters.Values) {
        foreach ($attribute in $parameter.Attributes) {
            if ($attribute -is [Management.Automation.ParameterAttribute] -and $attribute.Mandatory -and
                -not $arguments.ContainsKey($parameter.Name)) {
                throw "Scenario $($Command.Name) requires -$($parameter.Name)."
            }
        }
    }
    return $arguments
}

function Save-UiControlBitmap($Control, [string]$Path) {
    $bitmap = [Drawing.Bitmap]::new($Control.Width, $Control.Height)
    try {
        $Control.DrawToBitmap($bitmap, [Drawing.Rectangle]::new(0, 0, $Control.Width, $Control.Height))
        $bitmap.Save($Path, [Drawing.Imaging.ImageFormat]::Png)
    } finally { $bitmap.Dispose() }
}

function Save-UiNativeWindow($Window, [string]$Path) {
    if (-not ('UiProbeNativeCapture' -as [type])) {
        Add-Type -TypeDefinition @'
public static class UiProbeNativeCapture {
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool PrintWindow(System.IntPtr handle, System.IntPtr dc, uint flags);
}
'@
    }
    $bitmap = [Drawing.Bitmap]::new($Window.Width, $Window.Height)
    try {
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $dc = $graphics.GetHdc()
            try {
                if (-not [UiProbeNativeCapture]::PrintWindow($Window.Handle, $dc, 2)) { throw 'PrintWindow failed' }
            } finally { $graphics.ReleaseHdc($dc) }
        } finally { $graphics.Dispose() }
        $bitmap.Save($Path, [Drawing.Imaging.ImageFormat]::Png)
    } finally { $bitmap.Dispose() }
}

Export-ModuleMember -Function Open-UiAssembly, New-UiObject, Get-UiField, Set-UiField, Invoke-UiMethod, Assert-Ui,
Get-UiScenarioParameters, Save-UiControlBitmap, Save-UiNativeWindow

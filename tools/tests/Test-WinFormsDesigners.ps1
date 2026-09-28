param([string]$AssemblyPath = 'artifacts/designer-build/CodexVBE.dll')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, System.Design
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$types = $assembly.GetTypes() | Where-Object {
    -not $_.IsAbstract -and $_.Namespace -eq 'CodexVBE' -and
    ([Windows.Forms.Form].IsAssignableFrom($_) -or [Windows.Forms.UserControl].IsAssignableFrom($_))
}
foreach ($type in $types) {
    $surface = [ComponentModel.Design.DesignSurface]::new()
    try {
        $surface.BeginLoad($type)
        if (-not $surface.IsLoaded -or $surface.LoadErrors.Count) { throw "$($type.Name): $($surface.LoadErrors -join '; ')" }
        $designerHost = $surface.GetService([ComponentModel.Design.IDesignerHost])
        $root = $designerHost.RootComponent
        if ($null -eq $designerHost.GetDesigner($root) -or $surface.View -isnot [Windows.Forms.Control]) { throw "$($type.Name): no editable designer" }
        $property = [ComponentModel.TypeDescriptor]::GetProperties($root)['Size']
        $original = $property.GetValue($root)
        $property.SetValue($root, [Drawing.Size]::new($original.Width + 10, $original.Height + 10))
        if ($root.Width -ne $original.Width + 10) { throw "$($type.Name): runtime code overrides designer sizing" }
        $property.SetValue($root, $original)
        Write-Output "PASS $($type.Name): DesignSurface load and size edit"
    } finally { $surface.Dispose() }
}
Write-Output "PASS $($types.Count) WinForms designers"

param([string]$AssemblyPath = 'artifacts/designer-build/CodexVBE.dll', [string]$OutputDirectory = 'artifacts/designer-validation')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, System.Design
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
[IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($OutputDirectory)) | Out-Null
$sourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../src/CodexVBE'))
$designerSources = @{}
foreach ($file in Get-ChildItem -LiteralPath $sourceRoot -Filter '*.Designer.cs' -Recurse) {
    $source = [IO.File]::ReadAllText($file.FullName)
    if ($source -match 'partial class (\w+)') { $designerSources[$Matches[1]] = $source }
}
$proof = @()
$types = $assembly.GetTypes() | Where-Object {
    -not $_.IsAbstract -and $_.Namespace -eq 'CodexVBE' -and
    ([Windows.Forms.Form].IsAssignableFrom($_) -or [Windows.Forms.UserControl].IsAssignableFrom($_))
}
foreach ($type in $types) {
    $surface = [ComponentModel.Design.DesignSurface]::new()
    try {
        # Model source-owned controls on a fresh Form/UserControl root. Loading the
        # compiled subclass directly models inherited UI and locks private fields.
        $context = New-Object System.ComponentModel.Design.DesigntimeLicenseContext
        $subject = [ComponentModel.LicenseManager]::CreateWithContext($type, $context)
        $surface.BeginLoad($type.BaseType)
        if (-not $surface.IsLoaded -or $surface.LoadErrors.Count) { throw "$($type.Name): $($surface.LoadErrors -join '; ')" }
        $designerHost = $surface.GetService([ComponentModel.Design.IDesignerHost])
        $root = $designerHost.RootComponent
        $root.Font = $subject.Font
        $root.Size = $subject.Size
        $root.MinimumSize = $subject.MinimumSize
        $root.Padding = $subject.Padding
        $root.AutoSize = $subject.AutoSize
        $root.Name = $type.Name
        foreach ($control in @($subject.Controls)) { $root.Controls.Add($control) }
        if ($null -eq $designerHost.GetDesigner($root) -or $surface.View -isnot [Windows.Forms.Control]) { throw "$($type.Name): no editable designer" }
        $property = [ComponentModel.TypeDescriptor]::GetProperties($root)['Size']
        $autoSize = [ComponentModel.TypeDescriptor]::GetProperties($root)['AutoSize']
        $wasAutoSize = $null -ne $autoSize -and $autoSize.GetValue($root)
        if ($wasAutoSize) { $autoSize.SetValue($root, $false) }
        $original = $property.GetValue($root)
        $property.SetValue($root, [Drawing.Size]::new($original.Width + 10, $original.Height + 10))
        if ($root.Width -ne $original.Width + 10) { throw "$($type.Name): runtime code overrides designer sizing" }
        $property.SetValue($root, $original)
        if ($wasAutoSize) { $autoSize.SetValue($root, $true) }
        $children = 0
        if (-not $designerSources.ContainsKey($type.Name)) { throw "$($type.Name): no Designer source" }
        foreach ($match in [regex]::Matches($designerSources[$type.Name], 'this\.(\w+) = new [\w.]+\(')) {
            $field = $type.GetField($match.Groups[1].Value, [Reflection.BindingFlags]'Instance,NonPublic,Public')
            if ($null -eq $field) { continue }
            $child = $field.GetValue($subject)
            if ($child -isnot [Windows.Forms.Control]) { continue }
            if ([string]::IsNullOrEmpty($child.Name)) { throw "$($type.Name)/$($field.Name): unnamed Designer control" }
            # Register the fixed components declared by InitializeComponent in the design host.
            if ($null -eq $child.Site -or $child.Site.Container -ne $designerHost.Container) {
                if ($null -ne $child.Site) { $child.Site.Container.Remove($child) }
                $designerHost.Container.Add($child, $field.Name)
            }
            if ($null -eq $designerHost.GetDesigner($child)) { throw "$($type.Name)/$($field.Name): no child designer" }
            $description = [ComponentModel.TypeDescriptor]::GetProperties($child)['AccessibleDescription']
            if ($description.IsReadOnly) { throw "$($type.Name)/$($field.Name): editable property is read-only" }
            $prior = $description.GetValue($child)
            $description.SetValue($child, 'Designer edit verification')
            if ($description.GetValue($child) -ne 'Designer edit verification') { throw "$($type.Name)/$($field.Name): property edit failed" }
            $serialization = [ComponentModel.Design.Serialization.CodeDomComponentSerializationService]::new($designerHost)
            $store = $serialization.CreateStore()
            try {
                $serialization.SerializeAbsolute($store, $child)
                $store.Close()
                if ($store.Errors.Count) { throw "$($type.Name)/$($field.Name): serialization errors: $($store.Errors -join '; ')" }
                $copyContainer = New-Object System.ComponentModel.Container
                try {
                    $copies = @($serialization.Deserialize($store, $copyContainer))
                    $copy = @($copies | Where-Object { $_ -is [Windows.Forms.Control] -and $_.Name -eq $child.Name })
                    if ($copy.Count -ne 1 -or $copy[0].AccessibleDescription -ne 'Designer edit verification') { throw "$($type.Name)/$($field.Name): edited property did not roundtrip" }
                } finally { $copyContainer.Dispose() }
            } finally { $store.Dispose() }
            $description.SetValue($child, $prior)
            $children++
        }
        $proof += [pscustomobject]@{ View = $type.Name; Load = 'PASS'; Resize = 'PASS'; DesignerChildren = $children; ChildPropertyEdits = 'PASS'; SerializationRoundtrip = 'PASS'; AutoSizeRestored = $wasAutoSize }
        Write-Output "PASS $($type.Name): DesignSurface load, resize and $children editable child components"
    } finally { $surface.Dispose(); if ($null -ne $subject) { $subject.Dispose(); $subject = $null } }
}
$proof | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'designers.json') -Encoding UTF8
Write-Output "PASS $($types.Count) WinForms designers"

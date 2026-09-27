param([string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$sizes = @(16, 20, 24, 32, 48, 64, 128, 256)
foreach ($name in @('assistant', 'settings', 'github')) {
    $source = [Drawing.Image]::FromFile((Join-Path $ProjectRoot "assets/icons/$name.png"))
    try {
        $frames = @()
        foreach ($size in $sizes) {
            $bitmap = [Drawing.Bitmap]::new($size, $size)
            $graphics = [Drawing.Graphics]::FromImage($bitmap)
            $stream = [IO.MemoryStream]::new()
            try {
                $graphics.Clear([Drawing.Color]::Transparent)
                $graphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $graphics.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
                $scale = $size / [double][Math]::Max($source.Width, $source.Height)
                $width = [int][Math]::Round($source.Width * $scale)
                $height = [int][Math]::Round($source.Height * $scale)
                $graphics.DrawImage($source, [int](($size-$width)/2), [int](($size-$height)/2), $width, $height)
                $bitmap.Save($stream, [Drawing.Imaging.ImageFormat]::Png)
                $frames += ,$stream.ToArray()
            } finally { $stream.Dispose(); $graphics.Dispose(); $bitmap.Dispose() }
        }
        $output = [IO.File]::Create((Join-Path $ProjectRoot "assets/icons/$name.ico"))
        $writer = [IO.BinaryWriter]::new($output)
        try {
            $writer.Write([uint16]0); $writer.Write([uint16]1); $writer.Write([uint16]$sizes.Count)
            $offset = 6 + 16 * $sizes.Count
            for ($i = 0; $i -lt $sizes.Count; $i++) {
                $dimension = if ($sizes[$i] -eq 256) { 0 } else { $sizes[$i] }
                $writer.Write([byte]$dimension); $writer.Write([byte]$dimension)
                $writer.Write([byte]0); $writer.Write([byte]0)
                $writer.Write([uint16]1); $writer.Write([uint16]32)
                $writer.Write([uint32]$frames[$i].Length); $writer.Write([uint32]$offset)
                $offset += $frames[$i].Length
            }
            foreach ($frame in $frames) { $writer.Write([byte[]]$frame) }
        } finally { $writer.Dispose() }
    } finally { $source.Dispose() }
}

$forms = @{
    'src/Llm/Chat/ChatWindow.resx' = 'assistant'
    'src/Llm/Settings/LlmSettingsWindow.resx' = 'settings'
    'src/Git/GitWindow.resx' = 'github'
}
foreach ($relative in $forms.Keys) {
    $path = Join-Path $ProjectRoot $relative
    $document = [xml](Get-Content -LiteralPath $path -Raw)
    $old = $document.SelectSingleNode('/root/data[@name="$this.Icon"]')
    if ($old) { [void]$document.root.RemoveChild($old) }
    $data = $document.CreateElement('data')
    $data.SetAttribute('name', '$this.Icon')
    $data.SetAttribute('type', 'System.Drawing.Icon, System.Drawing')
    $data.SetAttribute('mimetype', 'application/x-microsoft.net.object.bytearray.base64')
    $value = $document.CreateElement('value')
    $value.InnerText = [Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $ProjectRoot ('assets/icons/' + $forms[$relative] + '.ico'))))
    [void]$data.AppendChild($value)
    [void]$document.root.AppendChild($data)
    $document.Save($path)
}

$chatPath = Join-Path $ProjectRoot 'src/Llm/Chat/ChatWindow.resx'
$document = [xml](Get-Content -LiteralPath $chatPath -Raw)
foreach ($entry in @(@('configure.Image', 'settings'), @('github.Image', 'github'))) {
    $old = $document.SelectSingleNode('/root/data[@name="' + $entry[0] + '"]')
    if ($old) { [void]$document.root.RemoveChild($old) }
    $icon = [Drawing.Icon]::new((Join-Path $ProjectRoot ('assets/icons/' + $entry[1] + '.ico')), 16, 16)
    $bitmap = $icon.ToBitmap()
    $stream = [IO.MemoryStream]::new()
    try {
        $bitmap.Save($stream, [Drawing.Imaging.ImageFormat]::Png)
        $data = $document.CreateElement('data')
        $data.SetAttribute('name', $entry[0])
        $data.SetAttribute('type', 'System.Drawing.Bitmap, System.Drawing')
        $data.SetAttribute('mimetype', 'application/x-microsoft.net.object.bytearray.base64')
        $value = $document.CreateElement('value')
        $value.InnerText = [Convert]::ToBase64String($stream.ToArray())
        [void]$data.AppendChild($value)
        [void]$document.root.AppendChild($data)
    } finally { $stream.Dispose(); $bitmap.Dispose(); $icon.Dispose() }
}
$document.Save($chatPath)

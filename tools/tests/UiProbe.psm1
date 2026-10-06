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

function Save-UiControlBitmap($Control, [string]$Path, [switch]$RenderRichText, [switch]$RenderHostedSurfaces) {
    $bitmap = [Drawing.Bitmap]::new($Control.Width, $Control.Height)
    try {
        $Control.DrawToBitmap($bitmap, [Drawing.Rectangle]::new(0, 0, $Control.Width, $Control.Height))
        # RichTextBox does not implement DrawToBitmap. Ask its own native RichEdit
        # renderer to paint the actual formatted content into the same surface.
        if ($RenderRichText) {
            if (-not ('UiProbeRichEditCapture' -as [type])) {
                Add-Type -ReferencedAssemblies System.Windows.Forms,System.Drawing -TypeDefinition @'
public static class UiProbeRichEditCapture {
    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
    private struct Rect { public int Left, Top, Right, Bottom; }
    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
    private struct Range { public System.IntPtr Hdc, Target; public Rect Area, Page; public int First, Last; }
    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint="SendMessageW")]
    private static extern System.IntPtr Format(System.IntPtr hwnd, int message, System.IntPtr paint, ref Range range);
    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint="SendMessageW")]
    private static extern System.IntPtr Clear(System.IntPtr hwnd, int message, System.IntPtr paint, System.IntPtr range);
    [System.Runtime.InteropServices.DllImport("gdi32.dll")]
    private static extern bool SetViewportOrgEx(System.IntPtr hdc, int x, int y, System.IntPtr previous);
    [System.Runtime.InteropServices.DllImport("gdi32.dll")]
    private static extern int IntersectClipRect(System.IntPtr hdc, int left, int top, int right, int bottom);
    public static void Paint(System.Windows.Forms.Control root, System.Drawing.Bitmap bitmap) {
        Visit(root, root, bitmap);
    }
    private static void Visit(System.Windows.Forms.Control root, System.Windows.Forms.Control control, System.Drawing.Bitmap bitmap) {
        if (!control.Visible) return;
        var rich = control as System.Windows.Forms.RichTextBox;
        if (rich != null && rich.TextLength > 0) {
            var point = root.PointToClient(rich.PointToScreen(System.Drawing.Point.Empty));
            if (root is System.Windows.Forms.Form) {
                var clientOrigin = root.PointToScreen(System.Drawing.Point.Empty);
                point.Offset(clientOrigin.X-root.Left, clientOrigin.Y-root.Top);
            }
            using (var graphics = System.Drawing.Graphics.FromImage(bitmap)) {
                using (var brush = new System.Drawing.SolidBrush(rich.BackColor))
                    graphics.FillRectangle(brush, point.X, point.Y, rich.ClientSize.Width, rich.ClientSize.Height);
                float scaleX = 1440f / graphics.DpiX, scaleY = 1440f / graphics.DpiY;
                var hdc = graphics.GetHdc();
                try {
                    SetViewportOrgEx(hdc, point.X, point.Y, System.IntPtr.Zero);
                    IntersectClipRect(hdc, 0, 0, rich.ClientSize.Width, rich.ClientSize.Height);
                    var range = new Range { Hdc=hdc, Target=hdc, First=0, Last=-1,
                        Area=new Rect { Left=2, Top=2, Right=(int)(rich.ClientSize.Width*scaleX), Bottom=(int)(rich.ClientSize.Height*scaleY) } };
                    range.Page=range.Area;
                    Format(rich.Handle, 0x0439, new System.IntPtr(1), ref range);
                } finally {
                    Clear(rich.Handle, 0x0439, System.IntPtr.Zero, System.IntPtr.Zero);
                    graphics.ReleaseHdc(hdc);
                }
            }
        }
        foreach (System.Windows.Forms.Control child in control.Controls) Visit(root, child, bitmap);
    }
}
'@
            }
            [UiProbeRichEditCapture]::Paint($Control,$bitmap)
        }
        if ($RenderHostedSurfaces) { Add-UiHostedSurfaces $Control $Control $bitmap }
        $bitmap.Save($Path, [Drawing.Imaging.ImageFormat]::Png)
    } finally { $bitmap.Dispose() }
}

function Add-UiHostedSurfaces($Root, $Control, $Bitmap) {
    if (-not $Control.Visible) { return }
    if ($Control -is [Windows.Forms.Integration.ElementHost] -and $null -ne $Control.Child) {
        $visual=$Control.Child
        $visual.UpdateLayout()
        $render=[Windows.Media.Imaging.RenderTargetBitmap]::new($Control.ClientSize.Width,$Control.ClientSize.Height,96,96,[Windows.Media.PixelFormats]::Pbgra32)
        $render.Render($visual)
        $encoder=[Windows.Media.Imaging.PngBitmapEncoder]::new()
        $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($render))
        $stream=[IO.MemoryStream]::new()
        $graphics=[Drawing.Graphics]::FromImage($Bitmap)
        try {
            $encoder.Save($stream); $stream.Position=0
            $image=[Drawing.Image]::FromStream($stream)
            try {
                $point=$Control.PointToScreen([Drawing.Point]::Empty)
                $origin=if ($Root -is [Windows.Forms.Form]) { $Root.Location } else { $Root.PointToScreen([Drawing.Point]::Empty) }
                $graphics.DrawImageUnscaled($image,$point.X-$origin.X,$point.Y-$origin.Y)
            } finally { $image.Dispose() }
            # RenderTargetBitmap excludes WinForms airspace hosted inside WPF.
            # Capture those actual child controls as well, with native RichEdit.
            $queue=[Collections.Generic.Queue[Windows.DependencyObject]]::new(); $queue.Enqueue($visual)
            while ($queue.Count) {
                $node=$queue.Dequeue()
                if ($node -is [Windows.Forms.Integration.WindowsFormsHost] -and $null -ne $node.Child) {
                    $child=$node.Child
                    if ($child.Visible -and $child.Width -gt 0 -and $child.Height -gt 0) {
                        $childBitmap=[Drawing.Bitmap]::new($child.Width,$child.Height)
                        try {
                            $child.DrawToBitmap($childBitmap,[Drawing.Rectangle]::new(0,0,$child.Width,$child.Height))
                            [UiProbeRichEditCapture]::Paint($child,$childBitmap)
                            $point=$child.PointToScreen([Drawing.Point]::Empty)
                            $graphics.DrawImageUnscaled($childBitmap,$point.X-$origin.X,$point.Y-$origin.Y)
                        } finally { $childBitmap.Dispose() }
                    }
                }
                $count=[Windows.Media.VisualTreeHelper]::GetChildrenCount($node)
                for ($i=0;$i -lt $count;$i++) { $queue.Enqueue([Windows.Media.VisualTreeHelper]::GetChild($node,$i)) }
            }
        } finally { $graphics.Dispose(); $stream.Dispose() }
    }
    foreach ($child in $Control.Controls) { Add-UiHostedSurfaces $Root $child $Bitmap }
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

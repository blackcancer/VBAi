using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace CodexVBE
{
    /// <summary>Remaps legacy neutral chrome after its native renderer has drawn it.</summary>
    internal static class VbeNativeChrome
    {
        [StructLayout(LayoutKind.Sequential)] private struct Rect { internal int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] private struct Point { internal int X, Y; }
        [StructLayout(LayoutKind.Sequential)] private struct ComboInfo
        {
            internal int Size;
            internal Rect Item, Button;
            internal uint ButtonState;
            internal IntPtr Combo, Edit, List;
        }
        [DllImport("user32.dll")] private static extern bool GetComboBoxInfo(IntPtr window, ref ComboInfo information);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
        [DllImport("user32.dll")] private static extern bool GetCursorPos(out Point point);
        [DllImport("user32.dll")] private static extern bool ScreenToClient(IntPtr window, ref Point point);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out Rect rect);
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out Rect rect);
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr window, ref Point point);
        [DllImport("user32.dll")] private static extern IntPtr GetWindowDC(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetDC(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
        [DllImport("user32.dll")] private static extern int ReleaseDC(IntPtr window, IntPtr dc);
        [DllImport("gdi32.dll")] private static extern bool BitBlt(IntPtr destination, int x, int y, int width, int height, IntPtr source, int sourceX, int sourceY, uint operation);
        internal const int EditorBackground = 0x282d35;
        private static readonly int[] Neutral = BuildNeutral();
        private static readonly HashSet<int> OutputColors = BuildOutputs();
        private static readonly int[] CodeText = BuildCodeRamp(0xdcdcdc, 192);
        private static readonly int[] CodeComment = BuildCodeRamp(0x57a64a, 128);
        private static readonly int[] CodeKeyword = BuildCodeRamp(0x569cd6, 128);
        private static readonly HashSet<int> CodeOutputColors = BuildCodeOutputs();
        [ThreadStatic] private static bool painting;

        private static int[] BuildNeutral()
        {
            var colors = new int[256];
            for (int i = 0; i < 256; i++)
            {
                double t = i / 255.0;
                int r = (int)Math.Round(226 - 194 * t);
                int g = (int)Math.Round(232 - 196 * t);
                int b = (int)Math.Round(240 - 197 * t);
                colors[i] = (r << 16) | (g << 8) | b;
            }
            return colors;
        }

        private static HashSet<int> BuildOutputs()
        {
            var colors = new HashSet<int>(Neutral) { 0x344452, 0xf48771, 0x6a9955, 0xdcdcaa, 0x569cd6 };
            return colors;
        }

        internal static int MapPixel(int pixel)
        {
            int rgb = pixel & 0xffffff;
            if (OutputColors.Contains(rgb)) return pixel;
            int alpha = pixel & unchecked((int)0xff000000);
            if (rgb == 0xff0000) return alpha | 0xf48771;
            if (rgb == 0x008000 || rgb == 0x00ff00) return alpha | 0x6a9955;
            if (rgb == 0xffff00) return alpha | 0xdcdcaa;
            if (rgb == 0x0000ff || rgb == 0x000080) return alpha | 0x569cd6;
            int r = (rgb >> 16) & 255, g = (rgb >> 8) & 255, b = rgb & 255;
            int maximum = Math.Max(r, Math.Max(g, b)), minimum = Math.Min(r, Math.Min(g, b));
            if (maximum - minimum <= 12)
                return (pixel & unchecked((int)0xff000000)) | Neutral[(r + g + b) / 3];
            // Legacy inactive dock captions use the Windows blue-grey face color.
            if (r > 160 && b >= g && g >= r && b - r < 65)
                return (pixel & unchecked((int)0xff000000)) | 0x344452;
            // ClearType edge pixels are colored even for neutral text. Convert them
            // with the same luminance curve to avoid bright RGB fringes on dark faces.
            return alpha | Neutral[(3 * r + 6 * g + b) / 10];
        }

        internal static int MapCodePixel(int pixel)
        {
            int alpha = pixel & unchecked((int)0xff000000);
            int rgb = pixel & 0xffffff;
            if (CodeOutputColors.Contains(rgb)) return pixel;
            if (rgb == 0xf0f0f0) return alpha | EditorBackground;
            if (rgb == 0xe3e3e3) return alpha | 0x3e4651;
            int r = rgb >> 16, g = (rgb >> 8) & 255, b = rgb & 255;
            // Translate the antialiased ramp as well as the solid foreground.
            // Leaving its dark edge pixels unchanged produces broken-looking glyphs
            // when the black background is replaced with anthracite.
            if (r == g && g == b && r <= 192)
                return alpha | CodeText[r];
            if (r == 0 && g <= 128 && b == 0)
                return alpha | CodeComment[g];
            if (r == 0 && g == b && g <= 128)
                return alpha | CodeKeyword[g];
            // ClearType uses independent channel coverage: real VBE captures
            // contain edges such as (108,43,0) and (0,72,110). Leaving those
            // pixels on the original black ramp makes the glyph look eroded.
            // Use neutral coverage for the text ramp to avoid colored fringes.
            // Preserve native selection and breakpoint palette entries.
            if (rgb == 0x800000 || rgb == 0x000080 || rgb == 0x800080 || rgb == 0x808000)
                return pixel;
            if (r == 0 && g <= 128 && b <= 128)
                return alpha | CodeKeyword[(g + b + 1) / 2];
            if (r <= 192 && g <= 192 && b <= 192)
                return alpha | CodeText[(3 * r + 6 * g + b + 5) / 10];
            return pixel;
        }

        private static HashSet<int> BuildCodeOutputs()
        {
            var colors = new HashSet<int>(OutputColors);
            colors.UnionWith(CodeText);
            colors.UnionWith(CodeComment);
            colors.UnionWith(CodeKeyword);
            colors.Add(0x3e4651);
            return colors;
        }

        private static int BlendCodeColor(int foreground, int coverage, int scale)
        {
            const int background = EditorBackground;
            int result = 0;
            foreach (int shift in new[] { 16, 8, 0 })
            {
                int start = (background >> shift) & 255;
                int end = (foreground >> shift) & 255;
                int channel = start + ((end - start) * coverage + scale / 2) / scale;
                result |= channel << shift;
            }
            return result;
        }

        private static int[] BuildCodeRamp(int foreground, int scale)
        {
            var result = new int[scale + 1];
            for (int coverage = 0; coverage <= scale; coverage++)
                result[coverage] = BlendCodeColor(foreground, coverage, scale);
            return result;
        }

        internal static void PaintBorder(IntPtr window)
        {
            Rect bounds, client;
            var origin = new Point();
            if (!GetWindowRect(window, out bounds) || !GetClientRect(window, out client) || !ClientToScreen(window, ref origin)) return;
            int width = bounds.Right - bounds.Left, height = bounds.Bottom - bounds.Top;
            int left = origin.X - bounds.Left;
            int top = origin.Y - bounds.Top;
            int right = width - left - client.Right;
            int bottom = height - top - client.Bottom;
            // Only cover real non-client edge pixels. Do not cover a client-drawn
            // combo button, a scrollbar, or the text's client rectangle.
            left = Math.Min(2, Math.Max(0, left));
            right = Math.Min(2, Math.Max(0, right));
            bottom = Math.Min(2, Math.Max(0, bottom));
            if (width <= 0 || height <= 0 || (left == 0 && right == 0 && bottom == 0)) return;
            IntPtr dc = GetWindowDC(window);
            if (dc == IntPtr.Zero) return;
            try
            {
                using (var graphics = Graphics.FromHdc(dc))
                using (var brush = new SolidBrush(Color.FromArgb(62, 70, 81)))
                {
                    if (left > 0) graphics.FillRectangle(brush, 0, 0, left, height);
                    if (right > 0) graphics.FillRectangle(brush, width - right, 0, right, height);
                    if (bottom > 0) graphics.FillRectangle(brush, 0, height - bottom, width, bottom);
                    if (top > 0) graphics.FillRectangle(brush, 0, 0, width, 1);
                }
            }
            catch (Exception error) { LoadLog.Write("Native border painting failed: " + error.Message); }
            finally { ReleaseDC(window, dc); }
        }

        internal static void PaintComboButton(IntPtr window)
        {
            var info = new ComboInfo { Size = Marshal.SizeOf(typeof(ComboInfo)) };
            Rect client;
            if (!GetComboBoxInfo(window, ref info) || !GetClientRect(window, out client) || (info.ButtonState & 0x8000) != 0) return;
            var button = Rectangle.FromLTRB(info.Button.Left, info.Button.Top, info.Button.Right, info.Button.Bottom);
            var area = Rectangle.FromLTRB(0, 0, client.Right, client.Bottom);
            if (button.Width < 4 || button.Height < 4 || !area.Contains(button)) return;
            bool enabled = IsWindowEnabled(window);
            Point cursor;
            bool hot = GetCursorPos(out cursor) && ScreenToClient(window, ref cursor) && button.Contains(cursor.X, cursor.Y);
            Color face = !enabled ? Color.FromArgb(32, 36, 43) : (info.ButtonState & 8) != 0
                ? Color.FromArgb(52, 68, 82) : hot ? Color.FromArgb(62, 70, 81) : Color.FromArgb(40, 45, 53);
            IntPtr dc = GetDC(window);
            if (dc == IntPtr.Zero) return;
            try
            {
                using (var graphics = Graphics.FromHdc(dc))
                using (var background = new SolidBrush(face))
                using (var arrow = new SolidBrush(enabled ? Color.FromArgb(226, 232, 240) : Color.FromArgb(120, 128, 139)))
                using (var border = new Pen(Color.FromArgb(62, 70, 81)))
                {
                    graphics.FillRectangle(background, button);
                    int half = Math.Max(2, Math.Min(button.Width, button.Height) / 5);
                    int x = button.Left + button.Width / 2, y = button.Top + button.Height / 2;
                    graphics.FillPolygon(arrow, new[] { new System.Drawing.Point(x - half, y - 1),
                        new System.Drawing.Point(x + half, y - 1), new System.Drawing.Point(x, y + half - 1) });
                    graphics.DrawRectangle(border, 0, 0, Math.Max(0, client.Right - 1), Math.Max(0, client.Bottom - 1));
                    if (client.Right > 6 && client.Bottom > 6)
                        graphics.DrawRectangle(border, 1, 1, client.Right - 3, client.Bottom - 3);
                }
            }
            catch (Exception error) { LoadLog.Write("Native combo button painting failed: " + error.Message); }
            finally { ReleaseDC(window, dc); }
        }

        internal static void Paint(IntPtr window, bool client, IntPtr suppliedDc, bool hostedCaption = false, bool preserveDarkClient = false, bool codeSurface = false)
        {
            if (painting) return;
            Rect bounds, inner;
            if (!GetWindowRect(window, out bounds) || !GetClientRect(window, out inner)) return;
            var origin = new Point();
            if (!ClientToScreen(window, ref origin)) return;
            int width = client ? inner.Right : bounds.Right - bounds.Left;
            int height = client ? inner.Bottom : origin.Y - bounds.Top;
            if (hostedCaption)
            {
                Rect childBounds;
                IntPtr child = GetWindow(window, 5);
                if (child == IntPtr.Zero || !GetWindowRect(child, out childBounds)) return;
                height = childBounds.Top - bounds.Top;
            }
            if (width <= 0 || height <= 0 || width > 16384 || height > 16384) return;
            IntPtr dc = suppliedDc != IntPtr.Zero ? suppliedDc : client ? GetDC(window) : GetWindowDC(window);
            if (dc == IntPtr.Zero) return;
            painting = true;
            try
            {
                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppRgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        IntPtr copy = graphics.GetHdc();
                        try { if (!BitBlt(copy, 0, 0, width, height, dc, 0, 0, 0x00CC0020)) return; }
                        finally { graphics.ReleaseHdc(copy); }
                    }
                    // The Immediate window shares the native code palette. Inverting an
                    // already black editor background would turn it light again.
                    bool darkCode = codeSurface && HasDarkBackground(bitmap);
                    if (preserveDarkClient && !codeSurface && HasDarkBackground(bitmap)) return;
                    if (codeSurface && !darkCode) return;
                    int marginWidth = darkCode ? LightCodeMargin(bitmap) : 0;
                    bool changed = false;
                    var data = bitmap.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.ReadWrite, PixelFormat.Format32bppRgb);
                    try
                    {
                        var row = new int[width];
                        var converted = new int[width];
                        for (int y = 0; y < height; y++)
                        {
                            IntPtr address = IntPtr.Add(data.Scan0, y * data.Stride);
                            Marshal.Copy(address, row, 0, width);
                            for (int x = 0; x < width; x++)
                            {
                                converted[x] = darkCode ? (x < marginWidth ? MapPixel(row[x]) : MapCodePixel(row[x])) : MapPixel(row[x]);
                                if (converted[x] != row[x]) changed = true;
                            }
                            Marshal.Copy(converted, 0, address, width);
                        }
                    }
                    finally { bitmap.UnlockBits(data); }
                    // A native state notification can leave the surface unchanged.
                    // Do not present a second full image of an already themed bar.
                    if (!changed) return;
                    using (var target = Graphics.FromHdc(dc)) target.DrawImageUnscaled(bitmap, 0, 0);
                }
            }
            catch (Exception error) { LoadLog.Write("Native chrome painting failed: " + error.Message); }
            finally { painting = false; if (suppliedDc == IntPtr.Zero) ReleaseDC(window, dc); }
        }

        internal static void PaintPropertyRow(IntPtr dc, VbeNativeTheme.NativeRect bounds)
        {
            int width = bounds.Right - bounds.Left, height = bounds.Bottom - bounds.Top;
            if (width <= 0 || height <= 0 || width > 16384 || height > 2048) return;
            try
            {
                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppRgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        IntPtr copy = graphics.GetHdc();
                        try { if (!BitBlt(copy, 0, 0, width, height, dc, bounds.Left, bounds.Top, 0x00CC0020)) return; }
                        finally { graphics.ReleaseHdc(copy); }
                    }
                    var data = bitmap.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.ReadWrite, PixelFormat.Format32bppRgb);
                    try
                    {
                        var row = new int[width];
                        for (int y = 0; y < height; y++)
                        {
                            IntPtr address = IntPtr.Add(data.Scan0, y * data.Stride);
                            Marshal.Copy(address, row, 0, width);
                            for (int x = 0; x < width; x++) row[x] = MapPixel(row[x]);
                            Marshal.Copy(row, 0, address, width);
                        }
                    }
                    finally { bitmap.UnlockBits(data); }
                    using (var target = Graphics.FromHdc(dc)) target.DrawImageUnscaled(bitmap, bounds.Left, bounds.Top);
                }
            }
            catch (Exception error) { LoadLog.Write("Native property row painting failed: " + error.Message); }
        }

        private static bool HasDarkBackground(Bitmap bitmap)
        {
            int dark = 0;
            foreach (int y in new[] { bitmap.Height / 3, bitmap.Height * 2 / 3 })
                foreach (int x in new[] { bitmap.Width / 3, bitmap.Width * 2 / 3 })
                {
                    Color color = bitmap.GetPixel(x, y);
                    if (color.R < 64 && color.G < 64 && color.B < 64) dark++;
                }
            return dark == 4;
        }

        private static int LightCodeMargin(Bitmap bitmap)
        {
            int width = 0;
            for (int x = 0; x < Math.Min(48, bitmap.Width / 4); x++)
                if ((bitmap.GetPixel(x, bitmap.Height / 2).ToArgb() & 0xffffff) == 0xf0f0f0) width = x + 1;
            return width;
        }
    }
}

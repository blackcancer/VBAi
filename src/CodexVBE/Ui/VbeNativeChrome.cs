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
        /// <summary>Native rectangle bounds used by User32 window-coordinate APIs.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct Rect { /// <summary>Left, top, right, and bottom edge coordinates.</summary>
internal int Left, Top, Right, Bottom; }
        /// <summary>Native point used by User32 screen- and client-coordinate APIs.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct Point { /// <summary>X and Y coordinates in the API's current coordinate space.</summary>
internal int X, Y; }
        /// <summary>COMBOBOXINFO-compatible data describing a combo's rectangles, state, and child handles.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct ComboInfo
        {
            /// <summary>Structure size in bytes.</summary>
            internal int Size;
            /// <summary>Item and drop-down button rectangles in screen coordinates.</summary>
            internal Rect Item, Button;
            /// <summary>Native state flags for the drop-down button.</summary>
            internal uint ButtonState;
            /// <summary>Combo box, edit child, and list child handles.</summary>
            internal IntPtr Combo, Edit, List;
        }
        /// <summary>Reads native combo-box geometry and child handles.</summary><param name="window">Combo-box handle.</param><param name="information">Receives the combo-box information.</param><returns>Whether the information was read.</returns>
        [DllImport("user32.dll")] private static extern bool GetComboBoxInfo(IntPtr window, ref ComboInfo information);
        /// <summary>Checks whether a native window accepts input.</summary><param name="window">Window handle.</param><returns>Whether it is enabled.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
        /// <summary>Reads the cursor's screen position.</summary><param name="point">Receives the cursor coordinates.</param><returns>Whether the position was read.</returns>
        [DllImport("user32.dll")] private static extern bool GetCursorPos(out Point point);
        /// <summary>Converts a screen position into client coordinates for a window.</summary><param name="window">Target window.</param><param name="point">Coordinates to convert.</param><returns>Whether conversion succeeded.</returns>
        [DllImport("user32.dll")] private static extern bool ScreenToClient(IntPtr window, ref Point point);
        /// <summary>Gets a native window rectangle in screen coordinates.</summary><param name="window">Window handle.</param><param name="rect">Receives the bounds.</param><returns>Whether the bounds were read.</returns>
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out Rect rect);
        /// <summary>Gets a native window client rectangle.</summary><param name="window">Window handle.</param><param name="rect">Receives the client bounds.</param><returns>Whether the bounds were read.</returns>
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out Rect rect);
        /// <summary>Converts a client position into screen coordinates.</summary><param name="window">Source window.</param><param name="point">Coordinates to convert.</param><returns>Whether conversion succeeded.</returns>
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr window, ref Point point);
        /// <summary>Gets a device context for a native window's non-client area.</summary><param name="window">Window handle.</param><returns>Device context, or zero.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetWindowDC(IntPtr window);
        /// <summary>Gets a device context for a native window's client area.</summary><param name="window">Window handle.</param><returns>Device context, or zero.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetDC(IntPtr window);
        /// <summary>Gets a related native window using a Windows relationship selector.</summary><param name="window">Starting handle.</param><param name="command">Relationship selector.</param><returns>Related handle or zero.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
        /// <summary>Releases a device context acquired from a window.</summary><param name="window">Owning window.</param><param name="dc">Device context to release.</param><returns>Native release result.</returns>
        [DllImport("user32.dll")] private static extern int ReleaseDC(IntPtr window, IntPtr dc);
        /// <summary>Copies a rectangular pixel area between device contexts.</summary><param name="destination">Destination context.</param><param name="x">Destination x coordinate.</param><param name="y">Destination y coordinate.</param><param name="width">Copy width.</param><param name="height">Copy height.</param><param name="source">Source context.</param><param name="sourceX">Source x coordinate.</param><param name="sourceY">Source y coordinate.</param><param name="operation">Raster operation code.</param><returns>Whether the copy succeeded.</returns>
        [DllImport("gdi32.dll")] private static extern bool BitBlt(IntPtr destination, int x, int y, int width, int height, IntPtr source, int sourceX, int sourceY, uint operation);
        /// <summary>Defines the read rectangle callback.</summary>
/// <param name="window">The window used by this operation.</param>
/// <param name="rectangle">The rectangle used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
        internal delegate bool ReadRectangle(IntPtr window, out Rect rectangle);
        /// <summary>Defines the convert point callback.</summary>
/// <param name="window">The window used by this operation.</param>
/// <param name="point">The point used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
        internal delegate bool ConvertPoint(IntPtr window, ref Point point);
        /// <summary>Defines the read pointer callback.</summary>
/// <param name="point">The point used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
        internal delegate bool ReadPointer(out Point point);
        /// <summary>Defines the read combo callback.</summary>
/// <param name="window">The window used by this operation.</param>
/// <param name="information">The information used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
        internal delegate bool ReadCombo(IntPtr window, ref ComboInfo information);
        /// <summary>Stores the window bounds,client bounds used by VbeNativeChrome.</summary>
        internal static ReadRectangle WindowBounds = GetWindowRect, ClientBounds = GetClientRect;
        /// <summary>Stores the to screen,to client used by VbeNativeChrome.</summary>
        internal static ConvertPoint ToScreen = ClientToScreen, ToClient = ScreenToClient;
        /// <summary>Stores the pointer position used by VbeNativeChrome.</summary>
        internal static ReadPointer PointerPosition = GetCursorPos;
        /// <summary>Stores the combo information used by VbeNativeChrome.</summary>
        internal static ReadCombo ComboInformation = GetComboBoxInfo;
        /// <summary>Stores the window enabled used by VbeNativeChrome.</summary>
        internal static Func<IntPtr, bool> WindowEnabled = IsWindowEnabled;
        /// <summary>Stores the acquire window dc,acquire client dc used by VbeNativeChrome.</summary>
        internal static Func<IntPtr, IntPtr> AcquireWindowDc = GetWindowDC, AcquireClientDc = GetDC;
        /// <summary>Stores the related window used by VbeNativeChrome.</summary>
        internal static Func<IntPtr, uint, IntPtr> RelatedWindow = GetWindow;
        /// <summary>Stores the release window dc used by VbeNativeChrome.</summary>
        internal static Func<IntPtr, IntPtr, int> ReleaseWindowDc = ReleaseDC;
        /// <summary>Stores the create graphics from dc used by VbeNativeChrome.</summary>
        internal static Func<IntPtr, Graphics> CreateGraphicsFromDc = Graphics.FromHdc;
        /// <summary>Stores the create chrome bitmap used by VbeNativeChrome.</summary>
        internal static Func<int, int, Bitmap> CreateChromeBitmap = (width, height) => new Bitmap(width, height, PixelFormat.Format32bppRgb);

        /// <summary>RGB color used as the remapped native code-editor background.</summary>
        internal const int EditorBackground = 0x282d35;
        /// <summary>Lookup table mapping legacy neutral colors to the experiment palette.</summary>
        private static readonly int[] Neutral = BuildNeutral();
        /// <summary>Colors already emitted by the chrome remapper and left unchanged.</summary>
        private static readonly HashSet<int> OutputColors = BuildOutputs();
        /// <summary>Antialiasing ramp for neutral code text.</summary>
        private static readonly int[] CodeText = BuildCodeRamp(0xdcdcdc, 192);
        /// <summary>Antialiasing ramp for code comments.</summary>
        private static readonly int[] CodeComment = BuildCodeRamp(0x57a64a, 128);
        /// <summary>Antialiasing ramp for code keywords.</summary>
        private static readonly int[] CodeKeyword = BuildCodeRamp(0x569cd6, 128);
        /// <summary>Colors already mapped in code surfaces.</summary>
        private static readonly HashSet<int> CodeOutputColors = BuildCodeOutputs();
        /// <summary>Prevents recursive entry while this thread paints a captured surface.</summary>
        [ThreadStatic] private static bool painting;

        /// <summary>Builds the grayscale lookup table used to darken neutral legacy chrome.</summary>
        /// <returns>One remapped RGB value for each 8-bit grayscale input.</returns>
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

        /// <summary>Builds the set of chrome colors that are already in the destination palette.</summary>
        /// <returns>Set of output RGB colors preserved by <see cref="MapPixel"/>.</returns>
        private static HashSet<int> BuildOutputs()
        {
            var colors = new HashSet<int>(Neutral) { 0x344452, 0xf48771, 0x6a9955, 0xdcdcaa, 0x569cd6 };
            return colors;
        }

        /// <summary>Maps a captured legacy chrome pixel to the experiment's dark palette.</summary>
        /// <param name="pixel">32-bit pixel value from the native window surface.</param>
        /// <returns>The remapped pixel, preserving its alpha byte.</returns>
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

        /// <summary>Maps a captured editor pixel while retaining syntax colors and editor palette entries.</summary>
        /// <param name="pixel">32-bit pixel value from the native code surface.</param>
        /// <returns>The remapped pixel, preserving its alpha byte.</returns>
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

        /// <summary>Builds the set of colors already emitted by the code-surface mapping.</summary>
        /// <returns>Set of mapped code RGB values.</returns>
        private static HashSet<int> BuildCodeOutputs()
        {
            var colors = new HashSet<int>(OutputColors);
            colors.UnionWith(CodeText);
            colors.UnionWith(CodeComment);
            colors.UnionWith(CodeKeyword);
            colors.Add(0x3e4651);
            return colors;
        }

        /// <summary>Interpolates between the editor background and a foreground color at a coverage level.</summary>
        /// <param name="foreground">RGB foreground color.</param><param name="coverage">Coverage numerator.</param><param name="scale">Maximum coverage value.</param>
        /// <returns>Blended RGB color.</returns>
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

        /// <summary>Builds an antialiasing ramp from the editor background to one syntax foreground color.</summary>
        /// <param name="foreground">RGB foreground color at full coverage.</param><param name="scale">Number of coverage intervals.</param>
        /// <returns>Ramp containing the background through full-foreground colors.</returns>
        private static int[] BuildCodeRamp(int foreground, int scale)
        {
            var result = new int[scale + 1];
            for (int coverage = 0; coverage <= scale; coverage++)
                result[coverage] = BlendCodeColor(foreground, coverage, scale);
            return result;
        }

        /// <summary>Recolors the thin non-client edges of a native window.</summary>
        /// <param name="window">Window whose border is repainted.</param>
        internal static void PaintBorder(IntPtr window)
        {
            Rect bounds, client;
            var origin = new Point();
            if (!WindowBounds(window, out bounds) || !ClientBounds(window, out client) || !ToScreen(window, ref origin)) return;
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
            IntPtr dc = AcquireWindowDc(window);
            if (dc == IntPtr.Zero) return;
            try
            {
                using (var graphics = CreateGraphicsFromDc(dc))
                using (var brush = new SolidBrush(Color.FromArgb(62, 70, 81)))
                {
                    if (left > 0) graphics.FillRectangle(brush, 0, 0, left, height);
                    if (right > 0) graphics.FillRectangle(brush, width - right, 0, right, height);
                    if (bottom > 0) graphics.FillRectangle(brush, 0, height - bottom, width, bottom);
                    if (top > 0) graphics.FillRectangle(brush, 0, 0, width, 1);
                }
            }
            catch (Exception error) { LoadLog.Write("Native border painting failed: " + error.Message); }
            finally { ReleaseWindowDc(window, dc); }
        }

        /// <summary>Draws the combo-box button using the dark palette and current enabled/hot state.</summary>
        /// <param name="window">Native combo-box handle.</param>
        internal static void PaintComboButton(IntPtr window)
        {
            var info = new ComboInfo { Size = Marshal.SizeOf(typeof(ComboInfo)) };
            Rect client;
            if (!ComboInformation(window, ref info) || !ClientBounds(window, out client) || (info.ButtonState & 0x8000) != 0) return;
            var button = Rectangle.FromLTRB(info.Button.Left, info.Button.Top, info.Button.Right, info.Button.Bottom);
            var area = Rectangle.FromLTRB(0, 0, client.Right, client.Bottom);
            if (button.Width < 4 || button.Height < 4 || !area.Contains(button)) return;
            bool enabled = WindowEnabled(window);
            Point cursor;
            bool hot = PointerPosition(out cursor) && ToClient(window, ref cursor) && button.Contains(cursor.X, cursor.Y);
            Color face = !enabled ? Color.FromArgb(32, 36, 43) : (info.ButtonState & 8) != 0
                ? Color.FromArgb(52, 68, 82) : hot ? Color.FromArgb(62, 70, 81) : Color.FromArgb(40, 45, 53);
            IntPtr dc = AcquireClientDc(window);
            if (dc == IntPtr.Zero) return;
            try
            {
                using (var graphics = CreateGraphicsFromDc(dc))
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
            finally { ReleaseWindowDc(window, dc); }
        }

        /// <summary>Captures a native surface, remaps its pixels, and paints the result back when needed.</summary>
        /// <param name="window">Native window being themed.</param>
        /// <param name="client">Whether to process its client rectangle instead of its non-client caption.</param>
        /// <param name="suppliedDc">Existing device context, or zero to acquire and release one for the window.</param>
        /// <param name="hostedCaption">Whether the caption height is taken from the hosted child window.</param>
        /// <param name="preserveDarkClient">Whether an already-dark client surface should be left intact.</param>
        /// <param name="codeSurface">Whether to use the syntax-editor pixel mapping and dark-background guard.</param>
        internal static void Paint(IntPtr window, bool client, IntPtr suppliedDc, bool hostedCaption = false, bool preserveDarkClient = false, bool codeSurface = false)
        {
            if (painting) return;
            Rect bounds, inner;
            if (!WindowBounds(window, out bounds) || !ClientBounds(window, out inner)) return;
            var origin = new Point();
            if (!ToScreen(window, ref origin)) return;
            int width = client ? inner.Right : bounds.Right - bounds.Left;
            int height = client ? inner.Bottom : origin.Y - bounds.Top;
            if (hostedCaption)
            {
                Rect childBounds;
                IntPtr child = RelatedWindow(window, 5);
                if (child == IntPtr.Zero || !WindowBounds(child, out childBounds)) return;
                height = childBounds.Top - bounds.Top;
            }
            if (width <= 0 || height <= 0 || width > 16384 || height > 16384) return;
            IntPtr dc = suppliedDc != IntPtr.Zero ? suppliedDc : client ? AcquireClientDc(window) : AcquireWindowDc(window);
            if (dc == IntPtr.Zero) return;
            painting = true;
            try
            {
                using (var bitmap = CreateChromeBitmap(width, height))
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
                    using (var target = CreateGraphicsFromDc(dc)) target.DrawImageUnscaled(bitmap, 0, 0);
                }
            }
            catch (Exception error) { LoadLog.Write("Native chrome painting failed: " + error.Message); }
            finally { painting = false; if (suppliedDc == IntPtr.Zero) ReleaseWindowDc(window, dc); }
        }

        /// <summary>Stores the create property row bitmap used by VbeNativeChrome.</summary>
        internal static Func<int, int, Bitmap> CreatePropertyRowBitmap = (width, height) => new Bitmap(width, height, PixelFormat.Format32bppRgb);

        /// <summary>Remaps the captured pixels of one native property-list row in place.</summary>
        /// <param name="dc">Device context containing the row.</param>
        /// <param name="bounds">Row bounds in device-context coordinates.</param>
        internal static void PaintPropertyRow(IntPtr dc, VbeNativeTheme.NativeRect bounds)
        {
            int width = bounds.Right - bounds.Left, height = bounds.Bottom - bounds.Top;
            if (width <= 0 || height <= 0 || width > 16384 || height > 2048) return;
            try
            {
                using (var bitmap = CreatePropertyRowBitmap(width, height))
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
                    using (var target = CreateGraphicsFromDc(dc)) target.DrawImageUnscaled(bitmap, bounds.Left, bounds.Top);
                }
            }
            catch (Exception error) { LoadLog.Write("Native property row painting failed: " + error.Message); }
        }

        /// <summary>Checks four interior sample pixels to determine whether a captured surface already has a dark background.</summary>
        /// <param name="bitmap">Captured surface to inspect.</param><returns><see langword="true"/> when all four samples are dark.</returns>
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

        /// <summary>Finds the right edge of a light left margin in a captured code surface.</summary>
        /// <param name="bitmap">Captured code surface.</param><returns>Margin width in pixels, or zero when none is detected.</returns>
        private static int LightCodeMargin(Bitmap bitmap)
        {
            int width = 0;
            for (int x = 0; x < Math.Min(48, bitmap.Width / 4); x++)
                if ((bitmap.GetPixel(x, bitmap.Height / 2).ToArgb() & 0xffffff) == 0xf0f0f0) width = x + 1;
            return width;
        }
    }
}

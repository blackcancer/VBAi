using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;

namespace CodexVBE.Tests.Unit
{
    internal static class NativeChromeCanvas
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct BitmapInfo
        {
            internal uint Size;
            internal int Width, Height;
            internal ushort Planes, Bits;
            internal uint Compression, ImageSize;
            internal int XPixelsPerMeter, YPixelsPerMeter;
            internal uint ColorsUsed, ColorsImportant;
        }
        [DllImport("gdi32.dll")] private static extern IntPtr CreateCompatibleDC(IntPtr source);
        [DllImport("gdi32.dll")] private static extern IntPtr CreateDIBSection(IntPtr dc, ref BitmapInfo info, uint usage,
            out IntPtr bits, IntPtr section, uint offset);
        [DllImport("gdi32.dll")] private static extern IntPtr SelectObject(IntPtr dc, IntPtr item);
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr item);
        [DllImport("gdi32.dll")] private static extern bool DeleteDC(IntPtr dc);
        [DllImport("gdi32.dll")] private static extern bool GdiFlush();

        internal static Bitmap Paint(Color initial, Action<IntPtr> paint)
        {
            IntPtr dc = CreateCompatibleDC(IntPtr.Zero);
            if (dc == IntPtr.Zero) throw new InvalidOperationException("Synthetic memory DC creation failed.");
            IntPtr image = IntPtr.Zero, previous = IntPtr.Zero;
            try
            {
                var info = new BitmapInfo { Size = 40, Width = 24, Height = -16, Planes = 1, Bits = 32 };
                IntPtr bits;
                image = CreateDIBSection(dc, ref info, 0, out bits, IntPtr.Zero, 0);
                if (image == IntPtr.Zero || bits == IntPtr.Zero) throw new InvalidOperationException("Synthetic DIB creation failed.");
                previous = SelectObject(dc, image);
                Marshal.Copy(Enumerable.Repeat(initial.ToArgb(), 24 * 16).ToArray(), 0, bits, 24 * 16);
                paint(dc);
                GdiFlush();
                var pixels = new int[24 * 16];
                Marshal.Copy(bits, pixels, 0, pixels.Length);
                var result = new Bitmap(24, 16, PixelFormat.Format32bppRgb);
                var data = result.LockBits(new Rectangle(0, 0, 24, 16), ImageLockMode.WriteOnly, PixelFormat.Format32bppRgb);
                try
                {
                    for (int y = 0; y < 16; y++) Marshal.Copy(pixels, y * 24, IntPtr.Add(data.Scan0, y * data.Stride), 24);
                }
                finally { result.UnlockBits(data); }
                return result;
            }
            finally
            {
                if (previous != IntPtr.Zero) SelectObject(dc, previous);
                if (image != IntPtr.Zero) DeleteObject(image);
                DeleteDC(dc);
            }
        }
    }
}

using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeNativePropertyTabsTests
    {
        [StructLayout(LayoutKind.Sequential)] private struct Controls { internal uint Size, Classes; }
        [StructLayout(LayoutKind.Sequential)] private struct Rect { internal int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] private struct Point { internal int X, Y; }
        [StructLayout(LayoutKind.Sequential)] private struct TabItem
        {
            internal uint Mask, State, StateMask;
            internal IntPtr Text;
            internal int TextCapacity, Image;
            internal IntPtr Data;
        }
        [DllImport("comctl32.dll")] private static extern bool InitCommonControlsEx(ref Controls controls);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr CreateWindowEx(int extended, string className, string name, int style, int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr window);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")] private static extern int GetStyle(IntPtr window, int index);
        [DllImport("user32.dll", EntryPoint = "SendMessageW")] private static extern IntPtr Send(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", EntryPoint = "SendMessageW")] private static extern IntPtr SendItem(IntPtr window, uint message, IntPtr index, ref TabItem item);
        [DllImport("user32.dll", EntryPoint = "SendMessageW")] private static extern IntPtr SendRect(IntPtr window, uint message, IntPtr index, out Rect bounds);
        [DllImport("user32.dll")] private static extern bool EnableWindow(IntPtr window, bool enabled);
        [DllImport("user32.dll")] private static extern bool InvalidateRect(IntPtr window, IntPtr rectangle, bool erase);
        [DllImport("user32.dll")] private static extern bool GetUpdateRect(IntPtr window, out Rect bounds, bool erase);
        [DllImport("user32.dll")] private static extern uint GetGuiResources(IntPtr process, uint flags);
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr value);
        [DllImport("gdi32.dll")] private static extern uint SetTextColor(IntPtr dc, uint color);
        [DllImport("gdi32.dll")] private static extern uint GetTextColor(IntPtr dc);
        [DllImport("gdi32.dll")] private static extern IntPtr GetCurrentObject(IntPtr dc, uint kind);
        [DllImport("gdi32.dll")] private static extern bool SetViewportOrgEx(IntPtr dc, int x, int y, out Point previous);
        [DllImport("gdi32.dll")] private static extern bool GetViewportOrgEx(IntPtr dc, out Point origin);
        [DllImport("gdi32.dll")] private static extern int IntersectClipRect(IntPtr dc, int left, int top, int right, int bottom);

        [TestMethod]
        public void NativeSelectionDrivesFinalColorsWithoutChangingTheFontOrSelectionBehavior()
        {
            OnSta(() =>
            {
                using (var fixture = new Fixture())
                {
                    Assert.IsTrue(VbeNativePropertyTabs.CanRender(fixture.Window));
                    IntPtr nativeFont = Send(fixture.Window, 0x0031, IntPtr.Zero, IntPtr.Zero);
                    Rect first = fixture.Bounds(0), second = fixture.Bounds(1);
                    using (Bitmap before = fixture.Capture(0x0318, 4))
                    {
                        Assert.AreEqual(Color.FromArgb(86, 156, 214).ToArgb(), before.GetPixel(first.Left + 5, first.Top).ToArgb());
                        Send(fixture.Window, 0x130c, new IntPtr(1), IntPtr.Zero);
                        Assert.AreEqual(1, Send(fixture.Window, 0x130b, IntPtr.Zero, IntPtr.Zero).ToInt32());
                        using (Bitmap after = fixture.Capture(0x0318, 4))
                        {
                            Assert.AreEqual(Color.FromArgb(86, 156, 214).ToArgb(), after.GetPixel(second.Left + 5, second.Top).ToArgb());
                            Assert.AreNotEqual(before.GetPixel(first.Left + 5, first.Top), after.GetPixel(first.Left + 5, first.Top));
                            Assert.IsTrue(CountLightText(after) > 20, "Native tab labels must actually be rendered into the supplied DC.");
                        }
                    }
                    Assert.AreEqual(nativeFont, Send(fixture.Window, 0x0031, IntPtr.Zero, IntPtr.Zero));
                    EnableWindow(fixture.Window, false);
                    using (Bitmap disabled = fixture.Capture(0x0318, 4))
                        Assert.AreEqual(Color.FromArgb(120, 128, 139).ToArgb(), disabled.GetPixel(second.Left + 5, second.Top).ToArgb());
                    EnableWindow(fixture.Window, true);
                    // Native focus/selection messages are forwarded, not implemented by the renderer.
                    Send(fixture.Window, 0x1330, IntPtr.Zero, IntPtr.Zero);
                    Assert.AreEqual(0, Send(fixture.Window, 0x132f, IntPtr.Zero, IntPtr.Zero).ToInt32());
                }
            });
        }

        [TestMethod]
        public void PrintFlagsClipAndViewportAreRespectedAndCallerDcStateIsRestored()
        {
            OnSta(() =>
            {
                using (var fixture = new Fixture())
                using (Bitmap client = fixture.Capture(0x0318, 4))
                using (Bitmap full = fixture.Capture(0x0317, 4 | 2 | 16))
                using (Bitmap nonClientOnly = fixture.Capture(0x0317, 2))
                using (Bitmap hiddenVisibleOnly = fixture.Capture(0x0317, 1 | 4))
                {
                    AssertBitmapsEqual(client, full);
                    Assert.AreEqual(Color.Magenta.ToArgb(), nonClientOnly.GetPixel(8, 8).ToArgb());
                    Assert.AreEqual(Color.Magenta.ToArgb(), hiddenVisibleOnly.GetPixel(8, 8).ToArgb());
                    using (var bitmap = new Bitmap(340, 130, PixelFormat.Format32bppRgb))
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.Clear(Color.Magenta);
                        IntPtr dc = graphics.GetHdc();
                        try
                        {
                            SetViewportOrgEx(dc, 11, 13, out _);
                            IntersectClipRect(dc, 0, 0, 20, 20);
                            SetTextColor(dc, 0x00332211);
                            IntPtr font = GetCurrentObject(dc, 6);
                            Send(fixture.Window, 0x0318, dc, new IntPtr(4));
                            Assert.AreEqual(0x00332211u, GetTextColor(dc));
                            Assert.AreEqual(font, GetCurrentObject(dc, 6));
                            GetViewportOrgEx(dc, out Point origin);
                            Assert.AreEqual(11, origin.X);
                            Assert.AreEqual(13, origin.Y);
                        }
                        finally { graphics.ReleaseHdc(dc); }
                        Assert.AreEqual(Color.Magenta.ToArgb(), bitmap.GetPixel(1, 1).ToArgb());
                        Assert.AreNotEqual(Color.Magenta.ToArgb(), bitmap.GetPixel(12, 14).ToArgb());
                        Assert.AreEqual(Color.Magenta.ToArgb(), bitmap.GetPixel(40, 14).ToArgb());
                    }
                    Assert.IsTrue(fixture.Renderer.PrintCount >= 3);
                }
            });
        }

        [TestMethod]
        public void PaintConsumesTheInvalidRegionAndOnlyStateChangesRequestAnotherPaint()
        {
            OnSta(() =>
            {
                using (var fixture = new Fixture(visible: true))
                {
                    InvalidateRect(fixture.Window, IntPtr.Zero, false);
                    Assert.IsTrue(GetUpdateRect(fixture.Window, out _, false), "Invalidation must establish a real native update region.");
                    // Explicit delivery exercises BeginPaint/EndPaint in this hidden native
                    // control harness. The product uses normal invalidation and UpdateWindow.
                    Send(fixture.Window, 0x000f, IntPtr.Zero, IntPtr.Zero);
                    Assert.IsFalse(GetUpdateRect(fixture.Window, out _, false), "BeginPaint/EndPaint must validate the native region.");
                    Assert.IsTrue(fixture.Renderer.PaintCount > 0, "The direct paint handler must run.");
                    Rect bounds = fixture.Bounds(1);
                    IntPtr location = new IntPtr((bounds.Top + 4) << 16 | (bounds.Left + 4));
                    Send(fixture.Window, 0x0200, IntPtr.Zero, location);
                    Send(fixture.Window, 0x000f, IntPtr.Zero, IntPtr.Zero);
                    int painted = fixture.Renderer.PaintCount;
                    Send(fixture.Window, 0x0200, IntPtr.Zero, location);
                    Assert.IsFalse(GetUpdateRect(fixture.Window, out _, false), "Moving within the same tab must not continuously repaint.");
                    Assert.AreEqual(painted, fixture.Renderer.PaintCount);
                    using (Bitmap hover = fixture.Capture(0x0318, 4))
                        Assert.AreEqual(Color.FromArgb(52, 68, 82).ToArgb(), hover.GetPixel(bounds.Left + 2, bounds.Top + 3).ToArgb());
                }
            });
        }

        [TestMethod]
        public void UnsupportedTabsAndCrossThreadCallsAreRejectedAndRenderingReleasesItsResources()
        {
            OnSta(() =>
            {
                using (var fixture = new Fixture())
                {
                    bool accepted = true;
                    var other = new Thread(() => accepted = VbeNativePropertyTabs.CanRender(fixture.Window));
                    other.Start();
                    Assert.IsTrue(other.Join(2000));
                    Assert.IsFalse(accepted);
                    using (Bitmap warmup = fixture.Capture(0x0318, 4)) { }
                    uint before = GetGuiResources(Process.GetCurrentProcess().Handle, 0);
                    for (int i = 0; i < 100; i++) using (Bitmap image = fixture.Capture(0x0318, 4)) { }
                    uint after = GetGuiResources(Process.GetCurrentProcess().Handle, 0);
                    Assert.IsTrue(after <= before + 2, "Repeated rendering leaked GDI objects: " + before + " -> " + after);
                    fixture.Renderer.Dispose();
                    Assert.IsFalse(fixture.Renderer.TryHandleMessage(0x0318, IntPtr.Zero, new IntPtr(4), out _));
                }
                // The two alignment styles require TCS_FIXEDWIDTH; the native control
                // can normalize invalid combinations before the renderer sees them.
                foreach (int style in new[] { 0x0100, 0x0200, 0x2000, 0x0002, 0x0080, 0x0410, 0x0430, 0x4000 })
                    using (var fixture = new Fixture(style, attach: false))
                        Assert.IsFalse(VbeNativePropertyTabs.CanRender(fixture.Window),
                            "Unsupported native tab style was accepted: requested=0x" + style.ToString("X") +
                            ", actual=0x" + GetStyle(fixture.Window, -16).ToString("X"));
            });
        }

        private static int CountLightText(Bitmap bitmap)
        {
            int count = 0;
            for (int y = 3; y < 28; y++)
                for (int x = 0; x < bitmap.Width; x++)
                {
                    Color pixel = bitmap.GetPixel(x, y);
                    if (pixel.R > 140 && pixel.G > 140 && pixel.B > 140) count++;
                }
            return count;
        }

        private static void AssertBitmapsEqual(Bitmap first, Bitmap second)
        {
            for (int y = 0; y < first.Height; y++)
                for (int x = 0; x < first.Width; x++)
                    if (first.GetPixel(x, y) != second.GetPixel(x, y)) Assert.Fail("Different supplied-DC render at " + x + "," + y);
        }

        private sealed class Fixture : NativeWindow, IDisposable
        {
            private readonly Form parent;
            private readonly IntPtr font;
            internal IntPtr Window { get; }
            internal VbeNativePropertyTabs Renderer { get; }

            internal Fixture(int extraStyle = 0, bool attach = true, bool visible = false)
            {
                var controls = new Controls { Size = (uint)Marshal.SizeOf(typeof(Controls)), Classes = 8 };
                Assert.IsTrue(InitCommonControlsEx(ref controls));
                parent = new NonActivatingForm
                {
                    StartPosition = FormStartPosition.Manual,
                    Location = new System.Drawing.Point(-30000, -30000),
                    FormBorderStyle = FormBorderStyle.None,
                    ClientSize = new Size(300, 90),
                    ShowInTaskbar = false
                };
                Window = CreateWindowEx(0, "SysTabControl32", "Properties renderer fixture", 0x54001040 | extraStyle,
                    0, 0, 300, 90, parent.Handle, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                Assert.AreNotEqual(IntPtr.Zero, Window);
                font = SystemFonts.MessageBoxFont.ToHfont();
                Send(Window, 0x0030, font, IntPtr.Zero);
                int index = 0;
                foreach (string name in new[] { "Alphabetical", "Categorized" })
                {
                    IntPtr text = Marshal.StringToHGlobalUni(name);
                    try
                    {
                        var item = new TabItem { Mask = 1, Text = text };
                        Assert.AreEqual(index, SendItem(Window, 0x133e, new IntPtr(index), ref item).ToInt32());
                    }
                    finally { Marshal.FreeHGlobal(text); }
                    index++;
                }
                if (attach)
                {
                    Assert.IsTrue(VbeNativePropertyTabs.TryCreate(Window, out VbeNativePropertyTabs renderer));
                    Renderer = renderer;
                    AssignHandle(Window);
                }
                if (visible) parent.Show();
            }

            internal Rect Bounds(int index)
            {
                Assert.AreNotEqual(IntPtr.Zero, SendRect(Window, 0x130a, new IntPtr(index), out Rect bounds));
                return bounds;
            }

            internal Bitmap Capture(uint message, int flags)
            {
                var bitmap = new Bitmap(300, 90, PixelFormat.Format32bppRgb);
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.Magenta);
                    IntPtr dc = graphics.GetHdc();
                    try { Send(Window, message, dc, new IntPtr(flags)); }
                    finally { graphics.ReleaseHdc(dc); }
                }
                return bitmap;
            }

            protected override void WndProc(ref Message message)
            {
                if (Renderer != null && Renderer.TryHandleMessage((uint)message.Msg, message.WParam, message.LParam, out IntPtr result))
                {
                    message.Result = result;
                    return;
                }
                base.WndProc(ref message);
                Renderer?.AfterNativeMessage((uint)message.Msg, message.WParam, message.LParam);
            }

            public void Dispose()
            {
                Renderer?.Dispose();
                if (Handle != IntPtr.Zero) ReleaseHandle();
                DestroyWindow(Window);
                DeleteObject(font);
                parent.Dispose();
            }
        }

        private sealed class NonActivatingForm : Form
        {
            protected override bool ShowWithoutActivation => true;
        }

        private static void OnSta(Action action)
        {
            Exception failure = null;
            var thread = new Thread(() => { try { action(); } catch (Exception error) { failure = error; } });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            Assert.IsTrue(thread.Join(20000), "The native tab test did not finish.");
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}

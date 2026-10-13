using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Forms = System.Windows.Forms;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises actual HWND regions inside scrolling WPF viewports.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ChatDesignerHostViewportTests
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct Bounds { public int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")] private static extern int GetWindowRgn(IntPtr window, IntPtr region);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out Bounds bounds);
        [DllImport("gdi32.dll")] private static extern IntPtr CreateRectRgn(int left, int top, int right, int bottom);
        [DllImport("gdi32.dll")] private static extern int GetRgnBox(IntPtr region, out Bounds bounds);
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr region);

        /// <summary>Checks partial, empty and full native regions under scroll, resize and layout scale.</summary>
        [STATestMethod]
        public void NativeRegionNeverLeavesViewportWhenRowsScrollResizeAndScale()
        {
            using (var card = new Forms.UserControl { MinimumSize = new System.Drawing.Size(120, 320) })
            using (var host = new ChatDesignerHost(card))
            {
                var stack = new StackPanel();
                stack.Children.Add(new Border { Height = 110 }); stack.Children.Add(host); stack.Children.Add(new Border { Height = 450 });
                var scroll = new ScrollViewer { Content = stack, VerticalScrollBarVisibility = ScrollBarVisibility.Hidden, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
                var window = CreateWindow(scroll);
                try
                {
                    window.Show(); Pump(window);
                    foreach (double scale in new[] { 1d, 1.5d, 2d })
                    {
                        host.LayoutTransform = new ScaleTransform(scale, scale);
                        foreach (double height in new[] { 200d, 270d })
                        {
                            window.Height = height;
                            foreach (double offset in new[] { 0d, 80d, 180d, 480d, 760d, 0d })
                            {
                                scroll.ScrollToVerticalOffset(offset); Pump(window);
                                AssertNativeRegionContained(host, scroll);
                            }
                        }
                    }
                    stack.Children.Remove(host); Pump(window);
                    stack.Children.Insert(1, host); Pump(window);
                    AssertNativeRegionContained(host, scroll);
                }
                finally { window.Close(); }
            }
        }

        /// <summary>Checks intersections at negative screen coordinates and fractional DPI boundaries.</summary>
        [TestMethod]
        public void PixelBoundsRoundInwardAndRejectOffscreenOrEmptyIntersection()
        {
            var native = new Rect(-200, -100, 300, 400);
            Assert.AreEqual(new System.Drawing.Rectangle(0, 0, 300, 400), ChatDesignerHost.PixelClip(native, native));
            Assert.AreEqual(new System.Drawing.Rectangle(1, 51, 198, 149), ChatDesignerHost.PixelClip(native, new Rect(-199.25, -49.25, 199, 149.5)));
            Assert.AreEqual(System.Drawing.Rectangle.Empty, ChatDesignerHost.PixelClip(native, new Rect(150, 400, 10, 10)));
            Assert.AreEqual(System.Drawing.Rectangle.Empty, ChatDesignerHost.PixelClip(native, Rect.Empty));
            Assert.AreEqual(System.Drawing.Rectangle.Empty, ChatDesignerHost.PixelClip(native, new Rect(-199.8, -99.8, .4, .4)));
        }

        /// <summary>Checks that hosts without a scrolling ancestor retain their ordinary native layout.</summary>
        [STATestMethod]
        public void DetachedAndUnscrolledHostsKeepOrdinaryWindowsFormsHostLayout()
        {
            using (var card = new Forms.UserControl())
            using (var host = new ChatDesignerHost(card))
            {
                host.Measure(new Size(400, 500)); host.Arrange(new Rect(0, 0, 400, 500));
                var window = CreateWindow(host);
                try { window.Show(); Pump(window); Assert.AreEqual(0, ReadRegion(host.Handle, out _)); }
                finally { window.Close(); }
            }
        }

        private static Window CreateWindow(object content) => new Window
        {
            Content = content, Width = 420, Height = 250, ShowInTaskbar = false, ShowActivated = false,
            WindowStyle = WindowStyle.None,
            Left = SystemParameters.VirtualScreenLeft + SystemParameters.VirtualScreenWidth + 256,
            Top = SystemParameters.VirtualScreenTop
        };

        private static void Pump(Window window)
        {
            window.UpdateLayout(); Forms.Application.DoEvents(); window.UpdateLayout();
        }

        private static int ReadRegion(IntPtr window, out Bounds bounds)
        {
            IntPtr region = CreateRectRgn(0, 0, 0, 0);
            try { int type = GetWindowRgn(window, region); GetRgnBox(region, out bounds); return type; }
            finally { DeleteObject(region); }
        }

        private static void AssertNativeRegionContained(ChatDesignerHost host, ScrollViewer scroll)
        {
            Assert.AreNotEqual(IntPtr.Zero, host.Handle);
            int kind = ReadRegion(host.Handle, out Bounds clip);
            Assert.AreNotEqual(0, kind, "Every scrolling native row needs an explicit region.");
            if (kind == 1) return; // Empty region: the row is entirely outside the viewport.
            Assert.IsTrue(GetWindowRect(host.Handle, out Bounds native));
            DependencyObject viewport = FindViewport(scroll);
            Assert.IsNotNull(viewport);
            var element = (FrameworkElement)viewport;
            var start = element.PointToScreen(new Point());
            var end = element.PointToScreen(new Point(element.ActualWidth, element.ActualHeight));
            Assert.IsTrue(native.Left + clip.Left >= Math.Ceiling(start.X) && native.Top + clip.Top >= Math.Ceiling(start.Y));
            Assert.IsTrue(native.Left + clip.Right <= Math.Floor(end.X) && native.Top + clip.Bottom <= Math.Floor(end.Y), "A native message can cover content outside the history viewport.");
        }

        private static DependencyObject FindViewport(DependencyObject parent)
        {
            if (parent is ScrollContentPresenter) return parent;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var found = FindViewport(VisualTreeHelper.GetChild(parent, i)); if (found != null) return found;
            }
            return null;
        }
    }
}

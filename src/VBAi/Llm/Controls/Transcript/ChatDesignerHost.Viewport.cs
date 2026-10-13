using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace VBAi
{
    /// <summary>Clips native transcript windows to their WPF scrolling viewport.</summary>
    internal sealed partial class ChatDesignerHost
    {
        /// <summary>Native window whose last region was applied successfully.</summary>
        private IntPtr clipWindow;

        /// <summary>Last applied region in physical pixels relative to the native window.</summary>
        private Rectangle? clipBounds;

        /// <summary>Physical window bounds returned by User32.</summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct NativeBounds { internal int Left, Top, Right, Bottom; }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr window, out NativeBounds bounds);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int SetWindowRgn(IntPtr window, IntPtr region, bool redraw);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern IntPtr CreateRectRgn(int left, int top, int right, int bottom);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr value);

        /// <summary>Applies the intersection of all scrolling and explicitly clipped visual ancestors.</summary>
        /// <remarks>WPF clipping does not by itself bound a WindowsFormsHost HWND. Screen-space
        /// conversion keeps the region in physical pixels under scaling and mixed monitor DPI.</remarks>
        private void UpdateViewportClip()
        {
            IntPtr window = Handle;
            if (!IsLoaded || window == IntPtr.Zero || View == null || View.IsDisposed) return;
            if (!GetWindowRect(window, out NativeBounds native)) return;
            var windowBounds = new Rect(native.Left, native.Top, Math.Max(0, native.Right - native.Left), Math.Max(0, native.Bottom - native.Top));
            Rect visible = windowBounds;
            bool clipped = false;
            for (DependencyObject ancestor = VisualTreeHelper.GetParent(this); ancestor != null; ancestor = VisualTreeHelper.GetParent(ancestor))
            {
                if (!(ancestor is FrameworkElement element) || (!(element is ScrollContentPresenter) && !element.ClipToBounds)) continue;
                clipped = true;
                if (!element.IsVisible || element.ActualWidth <= 0 || element.ActualHeight <= 0)
                {
                    visible = Rect.Empty;
                    break;
                }
                var start = element.PointToScreen(new System.Windows.Point());
                var end = element.PointToScreen(new System.Windows.Point(element.ActualWidth, element.ActualHeight));
                visible.Intersect(new Rect(start, end));
                if (visible.IsEmpty) break;
            }
            Rectangle? bounds = clipped ? (Rectangle?)PixelClip(windowBounds, visible) : null;
            if (clipWindow == window && clipBounds == bounds) return;
            IntPtr region = IntPtr.Zero;
            if (bounds.HasValue)
            {
                Rectangle value = bounds.Value;
                region = CreateRectRgn(value.Left, value.Top, value.Right, value.Bottom);
                if (region == IntPtr.Zero) return;
            }
            if (SetWindowRgn(window, region, true) == 0)
            {
                if (region != IntPtr.Zero) DeleteObject(region);
                return;
            }
            // User32 owns a successfully assigned region; it must not be deleted by this host.
            clipWindow = window;
            clipBounds = bounds;
        }

        /// <summary>Rounds an intersected screen rectangle inward and offsets it into native window coordinates.</summary>
        /// <param name="window">Native window bounds in physical screen pixels.</param>
        /// <param name="visible">Visible intersection in the same physical screen coordinates.</param>
        /// <returns>An empty region outside the viewport, or a region wholly contained in it.</returns>
        internal static Rectangle PixelClip(Rect window, Rect visible)
        {
            if (visible.IsEmpty || window.IsEmpty) return Rectangle.Empty;
            visible.Intersect(window);
            if (visible.IsEmpty) return Rectangle.Empty;
            int left = (int)Math.Ceiling(visible.Left - window.Left);
            int top = (int)Math.Ceiling(visible.Top - window.Top);
            int right = (int)Math.Floor(visible.Right - window.Left);
            int bottom = (int)Math.Floor(visible.Bottom - window.Top);
            return right <= left || bottom <= top ? Rectangle.Empty : Rectangle.FromLTRB(left, top, right, bottom);
        }
    }
}

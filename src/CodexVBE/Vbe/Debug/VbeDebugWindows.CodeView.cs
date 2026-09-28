using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace CodexVBE
{
    internal static partial class VbeDebugWindows
    {
        [StructLayout(LayoutKind.Sequential)] internal struct ViewRect { public int Left, Top, Right, Bottom; }
        [DllImport("user32.dll", EntryPoint = "GetWindowRect")] private static extern bool NativeViewBounds(IntPtr hwnd, out ViewRect rect);
        internal delegate bool ViewBoundsReader(IntPtr hwnd, out ViewRect rect);
        /// <summary>Reads native window geometry; isolated in deterministic window tests.</summary>
        internal static ViewBoundsReader ViewBounds = NativeViewBounds;

        internal static object ChangeCodeView(string caption, bool procedure)
        {
            IntPtr root = FindVbeRoot();
            var windows = ChildWindows(root).Where(window => WindowText(window) == caption).ToArray();
            if (root == IntPtr.Zero || windows.Length != 1) throw new InvalidOperationException("The exact native code window was not identified.");
            IntPtr windowHandle = windows[0];
            for (IntPtr ancestor = windowHandle; ancestor != IntPtr.Zero; ancestor = ObjectBrowserParent(ancestor))
                if (!ObjectBrowserEnabled(ancestor)) throw new InvalidOperationException("The code window is disabled by a modal window.");
            var bars = new List<ViewRect>();
            var buttons = new List<Tuple<IntPtr, ViewRect>>();
            EnumChildWindows(windowHandle, (child, ignored) =>
            {
                if (!IsWindowVisible(child) || !ViewBounds(child, out ViewRect bounds)) return true;
                string kind = ClassName(child);
                if (kind == "ScrollBar" && bounds.Right-bounds.Left > bounds.Bottom-bounds.Top) bars.Add(bounds);
                if (kind == "ObtbarWndClass") buttons.Add(Tuple.Create(child,bounds));
                return true;
            }, IntPtr.Zero);
            if (bars.Count != 1) throw new InvalidOperationException("The code window must expose exactly one shared horizontal scrollbar.");
            ViewRect bar = bars[0];
            var matches = buttons.Where(item => item.Item2.Right == bar.Left && item.Item2.Top == bar.Top && item.Item2.Bottom == bar.Bottom).ToArray();
            if (matches.Length != 1) throw new InvalidOperationException("The native procedure/module view toolbar was not uniquely identified.");
            var target = matches[0];
            int width = target.Item2.Right-target.Item2.Left, height = target.Item2.Bottom-target.Item2.Top;
            if (height < 10 || height > 64 || width < height*1.5 || width > height*3 || !ObjectBrowserEnabled(target.Item1))
                throw new InvalidOperationException("Unrecognized native view-button geometry or disabled control.");
            int x = procedure ? width/4 : 3*width/4, y = height/2;
            IntPtr point = new IntPtr((y << 16) | x);
            // These owner-drawn buttons expose no UIA/MSAA action. Send mouse
            // messages only to their verified HWND using its local dimensions.
            SendMessageInt(target.Item1, 0x0201, new IntPtr(1), point);
            SendMessageInt(target.Item1, 0x0202, IntPtr.Zero, point);
            return new { Method = "Native view toolbar, local button center", Width = width, Height = height, X = x, Y = y };
        }
    }
}

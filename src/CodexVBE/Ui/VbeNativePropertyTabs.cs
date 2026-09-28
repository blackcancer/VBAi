using System;
using System.Runtime.InteropServices;
using System.Text;

namespace CodexVBE
{
    /// <summary>Draws the two standard Properties tabs using their native geometry and font.</summary>
    /// <remarks>
    /// The caller owns subclassing and restricts instances to the VBE Properties pane.
    /// The original HFONT, including its rasterization quality, is selected unchanged;
    /// glyph pixels are never captured or recolored. Unsupported layout changes return
    /// control to the native procedure, so this pilot does not promise dark overflow tabs.
    /// </remarks>
    internal sealed class VbeNativePropertyTabs : IDisposable
    {
        private const uint Paint = 0x000f, Erase = 0x0014, Print = 0x0317, PrintClient = 0x0318;
        private const int UnsupportedStyles = 0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0020 | 0x0080 | 0x0100 | 0x0200 | 0x2000 | 0x4000;
        private const int UnsupportedExtendedStyles = 0x00000001 | 0x00000100 | 0x00000200 | 0x00001000 | 0x00002000 | 0x00400000;
        private readonly IntPtr window;
        private readonly uint ownerThread;
        private bool disposed;
        private int hotItem = -1;
        private bool trackingMouse;
        internal int PaintCount { get; private set; }
        internal int PrintCount { get; private set; }

        [StructLayout(LayoutKind.Sequential)] internal struct Rect { internal int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] internal struct Point { internal int X, Y; }
        [StructLayout(LayoutKind.Sequential)] internal struct TabItem
        {
            internal uint Mask, State, StateMask;
            internal IntPtr Text;
            internal int TextCapacity, Image;
            internal IntPtr Data;
        }
        [StructLayout(LayoutKind.Sequential)] internal struct PaintState
        {
            internal IntPtr Dc;
            internal int Erase;
            internal Rect Bounds;
            internal int Restore, IncrementalUpdate;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)] internal byte[] Reserved;
        }
        [StructLayout(LayoutKind.Sequential)] internal struct TrackMouse
        {
            internal uint Size, Flags;
            internal IntPtr Window;
            internal uint HoverTime;
        }
        private sealed class Item
        {
            internal Rect Bounds;
            internal string Text;
            internal bool Highlighted;
        }
        private sealed class Snapshot
        {
            internal Rect Client;
            internal Item[] Items;
            internal IntPtr Font;
            internal int Selected, Focused;
            internal bool Enabled, ShowFocus;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")] private static extern int GetStyle(IntPtr window, int index);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetFocus();
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out Rect rectangle);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out Rect rectangle);
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr window, ref Point point);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
        [DllImport("user32.dll", EntryPoint = "SendMessageW")] private static extern IntPtr Send(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", EntryPoint = "SendMessageW")] private static extern IntPtr SendItem(IntPtr window, uint message, IntPtr index, ref TabItem item);
        [DllImport("user32.dll", EntryPoint = "SendMessageW")] private static extern IntPtr SendRect(IntPtr window, uint message, IntPtr index, out Rect rectangle);
        [DllImport("user32.dll")] private static extern IntPtr BeginPaint(IntPtr window, out PaintState state);
        [DllImport("user32.dll")] private static extern bool EndPaint(IntPtr window, ref PaintState state);
        [DllImport("user32.dll")] private static extern bool InvalidateRect(IntPtr window, IntPtr rectangle, bool erase);
        [DllImport("user32.dll")] private static extern bool UpdateWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool TrackMouseEvent(ref TrackMouse tracking);
        [DllImport("gdi32.dll")] private static extern int SaveDC(IntPtr dc);
        [DllImport("gdi32.dll")] private static extern bool RestoreDC(IntPtr dc, int saved);
        [DllImport("gdi32.dll")] private static extern int IntersectClipRect(IntPtr dc, int left, int top, int right, int bottom);
        [DllImport("gdi32.dll")] private static extern IntPtr SelectObject(IntPtr dc, IntPtr value);
        [DllImport("gdi32.dll")] private static extern IntPtr GetStockObject(int index);
        [DllImport("gdi32.dll")] private static extern uint SetDCBrushColor(IntPtr dc, uint color);
        [DllImport("gdi32.dll")] private static extern uint SetTextColor(IntPtr dc, uint color);
        [DllImport("gdi32.dll")] private static extern uint SetBkColor(IntPtr dc, uint color);
        [DllImport("gdi32.dll")] private static extern int SetBkMode(IntPtr dc, int mode);
        [DllImport("user32.dll")] private static extern int FillRect(IntPtr dc, ref Rect rectangle, IntPtr brush);
        [DllImport("user32.dll")] private static extern int FrameRect(IntPtr dc, ref Rect rectangle, IntPtr brush);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int DrawText(IntPtr dc, string text, int length, ref Rect bounds, uint flags);
        [DllImport("user32.dll")] private static extern bool DrawFocusRect(IntPtr dc, ref Rect bounds);

        internal delegate uint WindowThreadReader(IntPtr window, out uint process);
        internal delegate bool RectReader(IntPtr window, out Rect rectangle);
        internal delegate bool ScreenPointReader(IntPtr window, ref Point point);
        internal delegate IntPtr ItemReader(IntPtr window, uint message, IntPtr index, ref TabItem item);
        internal delegate IntPtr ItemRectReader(IntPtr window, uint message, IntPtr index, out Rect rectangle);
        internal delegate IntPtr PaintBeginner(IntPtr window, out PaintState state);
        internal delegate bool PaintEnder(IntPtr window, ref PaintState state);
        internal delegate bool MouseTracker(ref TrackMouse state);
        internal static Func<IntPtr, StringBuilder, int, int> ClassName = GetClassName;
        internal static Func<IntPtr, int, int> Style = GetStyle;
        internal static WindowThreadReader WindowThread = GetWindowThreadProcessId;
        internal static Func<uint> CurrentThread = GetCurrentThreadId;
        internal static Func<IntPtr, bool> ValidWindow = IsWindow, VisibleWindow = IsWindowVisible, EnabledWindow = IsWindowEnabled;
        internal static Func<IntPtr> Focus = GetFocus;
        internal static RectReader ClientBounds = GetClientRect, WindowBounds = GetWindowRect;
        internal static ScreenPointReader ScreenPoint = ClientToScreen;
        internal static Func<IntPtr, uint, IntPtr> RelatedWindow = GetWindow;
        internal static Func<IntPtr, uint, IntPtr, IntPtr, IntPtr> SendMessage = Send;
        internal static ItemReader ReadItem = SendItem;
        internal static ItemRectReader ReadItemRect = SendRect;
        internal static PaintBeginner StartPaint = BeginPaint;
        internal static PaintEnder FinishPaint = EndPaint;
        internal static Func<IntPtr, IntPtr, bool, bool> Invalidate = InvalidateRect;
        internal static Func<IntPtr, bool> Update = UpdateWindow;
        internal static MouseTracker Track = TrackMouseEvent;
        private VbeNativePropertyTabs(IntPtr window)
        {
            this.window = window;
            ownerThread = WindowThread(window, out _);
        }

        /// <summary>Accepts only two text-only horizontal native tabs on their owning thread.</summary>
        internal static bool CanRender(IntPtr window)
        {
            if (window == IntPtr.Zero || !ValidWindow(window) || WindowThread(window, out _) != CurrentThread()) return false;
            var name = new StringBuilder(64);
            ClassName(window, name, name.Capacity);
            if (name.ToString() != "SysTabControl32") return false;
            int style = Style(window, -16);
            // Window borders and RTL would require a distinct WM_PRINT coordinate contract.
            if ((style & (UnsupportedStyles | 0x00c00000)) != 0 || (Style(window, -20) & UnsupportedExtendedStyles) != 0) return false;
            if (SendMessage(window, 0x1304, IntPtr.Zero, IntPtr.Zero).ToInt32() != 2 ||
                SendMessage(window, 0x1302, IntPtr.Zero, IntPtr.Zero) != IntPtr.Zero ||
                SendMessage(window, 0x132c, IntPtr.Zero, IntPtr.Zero).ToInt32() > 1 || RelatedWindow(window, 5) != IntPtr.Zero) return false;
            Rect client, bounds;
            var origin = new Point();
            return ClientBounds(window, out client) && WindowBounds(window, out bounds) && ScreenPoint(window, ref origin) &&
                client.Right > 0 && client.Bottom > 0 && client.Right <= 16384 && client.Bottom <= 16384 &&
                origin.X == bounds.Left && origin.Y == bounds.Top && client.Right == bounds.Right - bounds.Left && client.Bottom == bounds.Bottom - bounds.Top;
        }

        /// <summary>Creates a renderer without changing styles, selection, focus, or fonts.</summary>
        internal static bool TryCreate(IntPtr window, out VbeNativePropertyTabs renderer)
        {
            renderer = CanRender(window) ? new VbeNativePropertyTabs(window) : null;
            return renderer != null;
        }

        /// <summary>Handles painting before the original window procedure; other messages remain native.</summary>
        internal bool TryHandleMessage(uint message, IntPtr wParam, IntPtr lParam, out IntPtr result)
        {
            result = IntPtr.Zero;
            if (disposed || ownerThread != CurrentThread() ||
                (message != Paint && message != Erase && message != Print && message != PrintClient)) return false;
            Snapshot state;
            if (!TryRead(out state)) return false;
            if (message == Erase)
            {
                // WM_PAINT supplies the complete background in the same transaction as its text.
                result = new IntPtr(1);
                return true;
            }
            if (message == Paint)
            {
                PaintCount++;
                PaintState paint;
                IntPtr dc = StartPaint(window, out paint);
                try { if (dc != IntPtr.Zero) Draw(dc, state); }
                finally { FinishPaint(window, ref paint); }
                return true;
            }
            if (wParam == IntPtr.Zero) return false;
            // Accepted controls have no non-client region or children. A non-client-only
            // request must leave the caller's DC untouched. Preserve its viewport and clip.
            long flags = lParam.ToInt64();
            if ((flags & 1) != 0 && !VisibleWindow(window)) return true;
            if (message == Print && (flags & (4 | 8)) == 0) return true;
            PrintCount++;
            if (message == Print && (flags & 4) == 0) DrawBackground(wParam, state.Client);
            else Draw(wParam, state);
            return true;
        }

        /// <summary>Repaints only this control after native state transitions or pointer changes.</summary>
        internal void AfterNativeMessage(uint message, IntPtr wParam, IntPtr lParam)
        {
            if (disposed || ownerThread != CurrentThread()) return;
            bool changed = false;
            if (message == 0x0200 && (Style(window, -16) & 0x0040) != 0)
            {
                int x = unchecked((short)lParam.ToInt64()), y = unchecked((short)(lParam.ToInt64() >> 16));
                int next = -1;
                for (int index = 0; index < 2; index++)
                {
                    Rect item;
                    if (ReadItemRect(window, 0x130a, new IntPtr(index), out item) != IntPtr.Zero &&
                        x >= item.Left && x < item.Right && y >= item.Top && y < item.Bottom) { next = index; break; }
                }
                changed = hotItem != next;
                hotItem = next;
                if (!trackingMouse)
                {
                    var track = new TrackMouse { Size = (uint)Marshal.SizeOf(typeof(TrackMouse)), Flags = 2, Window = window };
                    trackingMouse = Track(ref track);
                }
            }
            else if (message == 0x02a3)
            {
                changed = hotItem != -1;
                hotItem = -1;
                trackingMouse = false;
            }
            else changed = message == 0x0005 || message == 0x0007 || message == 0x0008 || message == 0x000a ||
                message == 0x0030 || message == 0x007d || message == 0x0128 || message == 0x031a ||
                message == 0x0100 || message == 0x0101 || message == 0x0201 || message == 0x0202 ||
                message == 0x02e0 || message == 0x130c || message == 0x1330 || message == 0x133d ||
                message == 0x1306 || message == 0x1307 || message == 0x133e || message == 0x1308 || message == 0x1309;
            if (changed && CanRender(window))
            {
                Invalidate(window, IntPtr.Zero, false);
                // Some native state messages draw synchronously. Complete our paint before
                // returning rather than posting a later screenshot recoloring pass.
                Update(window);
            }
        }

        private bool TryRead(out Snapshot state)
        {
            state = null;
            if (!CanRender(window)) return false;
            var snapshot = new Snapshot { Items = new Item[2] };
            ClientBounds(window, out snapshot.Client);
            IntPtr storage = Marshal.AllocHGlobal(1024 * sizeof(char));
            try
            {
                for (int index = 0; index < 2; index++)
                {
                    var item = new TabItem { Mask = 1 | 16, StateMask = 2, Text = storage, TextCapacity = 1024 };
                    Marshal.WriteInt16(storage, 0);
                    Rect bounds;
                    if (ReadItem(window, 0x133c, new IntPtr(index), ref item) == IntPtr.Zero ||
                        ReadItemRect(window, 0x130a, new IntPtr(index), out bounds) == IntPtr.Zero ||
                        bounds.Right <= bounds.Left || bounds.Bottom <= bounds.Top) return false;
                    // TCM_GETITEM may return a different text pointer. Copy it while the
                    // control is on its owning thread; never free the returned pointer.
                    string text = item.Text == IntPtr.Zero ? string.Empty : Marshal.PtrToStringUni(item.Text);
                    if (text.Length >= 1023) return false;
                    snapshot.Items[index] = new Item { Bounds = bounds, Text = text, Highlighted = (item.State & 2) != 0 };
                }
            }
            finally { Marshal.FreeHGlobal(storage); }
            snapshot.Selected = SendMessage(window, 0x130b, IntPtr.Zero, IntPtr.Zero).ToInt32();
            snapshot.Focused = SendMessage(window, 0x132f, IntPtr.Zero, IntPtr.Zero).ToInt32();
            snapshot.Font = SendMessage(window, 0x0031, IntPtr.Zero, IntPtr.Zero);
            if (snapshot.Font == IntPtr.Zero) snapshot.Font = GetStockObject(17);
            snapshot.Enabled = EnabledWindow(window);
            snapshot.ShowFocus = Focus() == window && (SendMessage(window, 0x0129, IntPtr.Zero, IntPtr.Zero).ToInt64() & 1) == 0;
            state = snapshot;
            return true;
        }

        private void Draw(IntPtr dc, Snapshot state)
        {
            int saved = SaveDC(dc);
            if (saved == 0) return;
            try
            {
                IntersectClipRect(dc, state.Client.Left, state.Client.Top, state.Client.Right, state.Client.Bottom);
                SelectObject(dc, state.Font);
                SetBkMode(dc, 1);
                Fill(dc, state.Client, Color(32, 36, 43));
                Rect frame = state.Client;
                frame.Top = Math.Max(state.Items[0].Bounds.Bottom, state.Items[1].Bounds.Bottom) - 1;
                Frame(dc, frame, Color(62, 70, 81));
                for (int index = 0; index < 2; index++)
                {
                    Item item = state.Items[index];
                    bool selected = index == state.Selected;
                    bool hot = state.Enabled && index == hotItem && (Style(window, -16) & 0x0040) != 0;
                    uint face = selected ? Color(40, 45, 53) : hot ? Color(52, 68, 82) : Color(32, 36, 43);
                    Fill(dc, item.Bounds, face);
                    Frame(dc, item.Bounds, Color(62, 70, 81));
                    if (selected)
                    {
                        Rect accent = item.Bounds;
                        accent.Bottom = Math.Min(accent.Bottom, accent.Top + 2);
                        Fill(dc, accent, state.Enabled ? Color(86, 156, 214) : Color(120, 128, 139));
                    }
                    Rect text = item.Bounds;
                    text.Left += 4; text.Right -= 4; text.Top += 2; text.Bottom -= 1;
                    SetBkColor(dc, face);
                    SetTextColor(dc, !state.Enabled ? Color(120, 128, 139) : item.Highlighted ? Color(86, 156, 214) : Color(226, 232, 240));
                    DrawText(dc, item.Text, item.Text.Length, ref text, 0x0001 | 0x0004 | 0x0020 | 0x0800 | 0x8000);
                    if (state.Enabled && state.ShowFocus && index == state.Focused)
                    {
                        text.Left -= 1; text.Right += 1;
                        DrawFocusRect(dc, ref text);
                    }
                }
            }
            finally { RestoreDC(dc, saved); }
        }

        private static void DrawBackground(IntPtr dc, Rect client)
        {
            int saved = SaveDC(dc);
            if (saved == 0) return;
            try { Fill(dc, client, Color(32, 36, 43)); }
            finally { RestoreDC(dc, saved); }
        }

        private static uint Color(int red, int green, int blue) => (uint)(red | green << 8 | blue << 16);
        private static void Fill(IntPtr dc, Rect bounds, uint color)
        {
            SetDCBrushColor(dc, color);
            FillRect(dc, ref bounds, GetStockObject(18));
        }
        private static void Frame(IntPtr dc, Rect bounds, uint color)
        {
            SetDCBrushColor(dc, color);
            FrameRect(dc, ref bounds, GetStockObject(18));
        }

        /// <summary>Stops mouse tracking; the native font and stock GDI objects remain owned by Windows.</summary>
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            if (trackingMouse && ownerThread == CurrentThread() && ValidWindow(window))
            {
                var track = new TrackMouse { Size = (uint)Marshal.SizeOf(typeof(TrackMouse)), Flags = 0x80000002, Window = window };
                Track(ref track);
            }
        }
    }
}

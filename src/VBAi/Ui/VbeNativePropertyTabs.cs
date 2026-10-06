using System;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi
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

        /// <summary>Paint, erase, print, and print-client messages handled by this renderer.</summary>
        private const uint Paint = 0x000f, Erase = 0x0014, Print = 0x0317, PrintClient = 0x0318;

        /// <summary>Window style bits that make direct horizontal-tab rendering unsupported.</summary>
        private const int UnsupportedStyles = 0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0020 | 0x0080 | 0x0100 | 0x0200 | 0x2000 | 0x4000;

        /// <summary>Extended window style bits that make direct tab rendering unsupported.</summary>
        private const int UnsupportedExtendedStyles = 0x00000001 | 0x00000100 | 0x00000200 | 0x00001000 | 0x00002000 | 0x00400000;

        /// <summary>Native tab-control handle rendered by this instance.</summary>
        private readonly IntPtr window;

        /// <summary>Thread identifier that owns the tab-control window.</summary>
        private readonly uint ownerThread;

        /// <summary>Whether this renderer has been disposed.</summary>
        private bool disposed;

        /// <summary>Current hot tab index, or -1 when the pointer is outside all tabs.</summary>
        private int hotItem = -1;

        /// <summary>Whether tracking mouse leave for the tab control is active.</summary>
        private bool trackingMouse;

        /// <summary>Gets the number of WM_PAINT messages rendered directly.</summary>
        /// <value>Direct paint count for this instance.</value>
        internal int PaintCount { get; private set; }

        /// <summary>Gets the number of WM_PRINT or WM_PRINTCLIENT messages rendered directly.</summary>
        /// <value>Direct print count for this instance.</value>
        internal int PrintCount { get; private set; }

        /// <summary>Native rectangle bounds in left, top, right, bottom order.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct Rect {

/// <summary>Left, top, right, and bottom edge coordinates.</summary>
internal int Left, Top, Right, Bottom; }

        /// <summary>Native point in the current coordinate space.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct Point {

/// <summary>X and Y coordinates.</summary>
internal int X, Y; }

        /// <summary>TCITEM-compatible structure used to read native tab text and state.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct TabItem
        {

            /// <summary>Fields requested, item state, and state mask.</summary>
            internal uint Mask, State, StateMask;

            /// <summary>Pointer to tab text storage.</summary>
            internal IntPtr Text;

            /// <summary>Text buffer capacity and image-list index.</summary>
            internal int TextCapacity, Image;

            /// <summary>Application-defined item data.</summary>
            internal IntPtr Data;
        }

        /// <summary>PAINTSTRUCT-compatible data for a native tab paint transaction.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct PaintState
        {

            /// <summary>Paint device context.</summary>
            internal IntPtr Dc;

            /// <summary>Whether background erasure is required.</summary>
            internal int Erase;

            /// <summary>Invalidated bounds.</summary>
            internal Rect Bounds;

            /// <summary>Native restore and incremental-update flags.</summary>
            internal int Restore, IncrementalUpdate;

            /// <summary>Reserved native padding bytes.</summary>
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)] internal byte[] Reserved;
        }

        /// <summary>TRACKMOUSEEVENT-compatible request for tab hover and leave notifications.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct TrackMouse
        {

            /// <summary>Structure size and tracking flags.</summary>
            internal uint Size, Flags;

            /// <summary>Window being tracked.</summary>
            internal IntPtr Window;

            /// <summary>Hover timeout requested from Windows.</summary>
            internal uint HoverTime;
        }

        /// <summary>One tab's bounds, text, and native highlight state.</summary>
        private sealed class Item
        {

            /// <summary>Tab bounds in client coordinates.</summary>
            internal Rect Bounds;

            /// <summary>Copied tab label.</summary>
            internal string Text;

            /// <summary>Whether the native tab state marks the item highlighted.</summary>
            internal bool Highlighted;
        }

        /// <summary>Snapshot of tab geometry, selection, font, focus, and enabled state for one paint.</summary>
        private sealed class Snapshot
        {

            /// <summary>Client rectangle and the two tab items.</summary>
            internal Rect Client;

            /// <summary>Snapshot of the two tab rows.</summary>
            internal Item[] Items;

            /// <summary>Native font handle retained by the control.</summary>
            internal IntPtr Font;

            /// <summary>Selected and focused tab indexes.</summary>
            internal int Selected, Focused;

            /// <summary>Control enabled state and whether the focus rectangle should be drawn.</summary>
            internal bool Enabled, ShowFocus;
        }

        /// <summary>Gets the native class name for a window.</summary><param name="window">Window handle.</param><param name="text">Output buffer.</param><param name="capacity">Buffer capacity.</param><returns>Characters copied.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);

        /// <summary>Reads a window style value.</summary><param name="window">Window handle.</param><param name="index">Style selector.</param><returns>Style bits.</returns>
        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")] private static extern int GetStyle(IntPtr window, int index);

        /// <summary>Gets the process and thread that created a window.</summary><param name="window">Window handle.</param><param name="process">Receives the process identifier.</param><returns>Thread identifier.</returns>
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);

        /// <summary>Gets the calling thread identifier.</summary><returns>Thread identifier.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

        /// <summary>Checks whether a window handle is valid.</summary><param name="window">Window handle.</param><returns>Whether it identifies an existing window.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);

        /// <summary>Checks whether a window is visible.</summary><param name="window">Window handle.</param><returns>Visibility state.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);

        /// <summary>Checks whether a window accepts input.</summary><param name="window">Window handle.</param><returns>Enabled state.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);

        /// <summary>Gets the window that currently owns keyboard focus.</summary><returns>Focused window handle, or zero.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetFocus();

        /// <summary>Gets the client rectangle of a window.</summary><param name="window">Window handle.</param><param name="rectangle">Receives the client bounds.</param><returns>Whether the rectangle was read.</returns>
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out Rect rectangle);

        /// <summary>Gets a window rectangle in screen coordinates.</summary><param name="window">Window handle.</param><param name="rectangle">Receives the screen bounds.</param><returns>Whether the rectangle was read.</returns>
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out Rect rectangle);

        /// <summary>Converts client coordinates into screen coordinates.</summary><param name="window">Source window.</param><param name="point">Coordinates to convert.</param><returns>Whether conversion succeeded.</returns>
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr window, ref Point point);

        /// <summary>Gets a related native window by relationship selector.</summary><param name="window">Starting window.</param><param name="command">Relationship selector.</param><returns>Related handle or zero.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);

        /// <summary>Sends a message to a native window.</summary><param name="window">Target handle.</param><param name="message">Message identifier.</param><param name="wParam">First message value.</param><param name="lParam">Second message value.</param><returns>Window-procedure result.</returns>
        [DllImport("user32.dll", EntryPoint = "SendMessageW")] private static extern IntPtr Send(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

        /// <summary>Sends a tab-control message that reads one tab item.</summary><param name="window">Tab-control handle.</param><param name="message">Message identifier.</param><param name="index">Zero-based tab index.</param><param name="item">Receives the tab fields.</param><returns>Native message result.</returns>
        [DllImport("user32.dll", EntryPoint = "SendMessageW")] private static extern IntPtr SendItem(IntPtr window, uint message, IntPtr index, ref TabItem item);

        /// <summary>Sends a tab-control message that reads one tab rectangle.</summary><param name="window">Tab-control handle.</param><param name="message">Message identifier.</param><param name="index">Zero-based tab index.</param><param name="rectangle">Receives the tab bounds.</param><returns>Native message result.</returns>
        [DllImport("user32.dll", EntryPoint = "SendMessageW")] private static extern IntPtr SendRect(IntPtr window, uint message, IntPtr index, out Rect rectangle);

        /// <summary>Begins a native paint transaction and obtains its drawing context.</summary><param name="window">Window being painted.</param><param name="state">Receives paint metadata.</param><returns>Paint device context.</returns>
        [DllImport("user32.dll")] private static extern IntPtr BeginPaint(IntPtr window, out PaintState state);

        /// <summary>Completes a paint transaction opened by BeginPaint.</summary><param name="window">Window being painted.</param><param name="state">Paint metadata returned by BeginPaint.</param><returns>Native completion result.</returns>
        [DllImport("user32.dll")] private static extern bool EndPaint(IntPtr window, ref PaintState state);

        /// <summary>Invalidates a window area for later repaint.</summary><param name="window">Window handle.</param><param name="rectangle">Area to invalidate, or zero for the full client area.</param><param name="erase">Whether to request background erasure.</param><returns>Whether invalidation succeeded.</returns>
        [DllImport("user32.dll")] private static extern bool InvalidateRect(IntPtr window, IntPtr rectangle, bool erase);

        /// <summary>Synchronously processes pending paint messages for a window.</summary><param name="window">Window handle.</param><returns>Native update result.</returns>
        [DllImport("user32.dll")] private static extern bool UpdateWindow(IntPtr window);

        /// <summary>Requests mouse hover or leave notifications for a window.</summary><param name="tracking">Tracking request and window handle.</param><returns>Whether tracking was registered.</returns>
        [DllImport("user32.dll")] private static extern bool TrackMouseEvent(ref TrackMouse tracking);

        /// <summary>Saves the current GDI device-context state.</summary><param name="dc">Device context.</param><returns>Saved state identifier, or zero.</returns>
        [DllImport("gdi32.dll")] private static extern int SaveDC(IntPtr dc);

        /// <summary>Restores a GDI device context to a saved state.</summary><param name="dc">Device context.</param><param name="saved">Saved state identifier.</param><returns>Whether restoration succeeded.</returns>
        [DllImport("gdi32.dll")] private static extern bool RestoreDC(IntPtr dc, int saved);

        /// <summary>Intersects the device-context clip with a rectangle.</summary><param name="dc">Device context.</param><param name="left">Left clip edge.</param><param name="top">Top clip edge.</param><param name="right">Right clip edge.</param><param name="bottom">Bottom clip edge.</param><returns>Clip-region result.</returns>
        [DllImport("gdi32.dll")] private static extern int IntersectClipRect(IntPtr dc, int left, int top, int right, int bottom);

        /// <summary>Selects a GDI object into a device context.</summary><param name="dc">Device context.</param><param name="value">Font or other GDI object handle.</param><returns>Previously selected object.</returns>
        [DllImport("gdi32.dll")] private static extern IntPtr SelectObject(IntPtr dc, IntPtr value);

        /// <summary>Gets a stock GDI object such as the DC brush.</summary><param name="index">Stock-object selector.</param><returns>Stock object handle.</returns>
        [DllImport("gdi32.dll")] private static extern IntPtr GetStockObject(int index);

        /// <summary>Sets the color of the stock DC brush.</summary><param name="dc">Device context.</param><param name="color">COLORREF value.</param><returns>Previous brush color.</returns>
        [DllImport("gdi32.dll")] private static extern uint SetDCBrushColor(IntPtr dc, uint color);

        /// <summary>Sets the text color used for subsequent drawing.</summary><param name="dc">Device context.</param><param name="color">COLORREF value.</param><returns>Previous text color.</returns>
        [DllImport("gdi32.dll")] private static extern uint SetTextColor(IntPtr dc, uint color);

        /// <summary>Sets the text background color used for subsequent drawing.</summary><param name="dc">Device context.</param><param name="color">COLORREF value.</param><returns>Previous background color.</returns>
        [DllImport("gdi32.dll")] private static extern uint SetBkColor(IntPtr dc, uint color);

        /// <summary>Sets the background fill mode for text drawing.</summary><param name="dc">Device context.</param><param name="mode">Opaque or transparent background mode.</param><returns>Previous mode.</returns>
        [DllImport("gdi32.dll")] private static extern int SetBkMode(IntPtr dc, int mode);

        /// <summary>Fills a rectangle with a GDI brush.</summary><param name="dc">Device context.</param><param name="rectangle">Rectangle to fill.</param><param name="brush">Brush handle.</param><returns>Native fill result.</returns>
        [DllImport("user32.dll")] private static extern int FillRect(IntPtr dc, ref Rect rectangle, IntPtr brush);

        /// <summary>Frames a rectangle with a GDI brush.</summary><param name="dc">Device context.</param><param name="rectangle">Rectangle to frame.</param><param name="brush">Brush handle.</param><returns>Native frame result.</returns>
        [DllImport("user32.dll")] private static extern int FrameRect(IntPtr dc, ref Rect rectangle, IntPtr brush);

        /// <summary>Draws a Unicode label within a rectangle.</summary><param name="dc">Device context.</param><param name="text">Text to draw.</param><param name="length">Text length.</param><param name="bounds">Text bounds, updated with any measured extent.</param><param name="flags">Alignment and formatting flags.</param><returns>Text height or formatted text result.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int DrawText(IntPtr dc, string text, int length, ref Rect bounds, uint flags);

        /// <summary>Draws the native focus rectangle for a tab item.</summary><param name="dc">Device context.</param><param name="bounds">Focus bounds.</param><returns>Native drawing result.</returns>
        [DllImport("user32.dll")] private static extern bool DrawFocusRect(IntPtr dc, ref Rect bounds);

        /// <summary>Defines the window thread reader callback.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="process">uint that supplies the process for this operation.</param>
        /// <returns>uint produced by the operation for operation on vbe native property tabs.</returns>
        internal delegate uint WindowThreadReader(IntPtr window, out uint process);

        /// <summary>Defines the rect reader callback.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="rectangle">rect that supplies the rectangle for this operation.</param>
        /// <returns>Boolean indicating the result of the check for operation on vbe native property tabs.</returns>
        internal delegate bool RectReader(IntPtr window, out Rect rectangle);

        /// <summary>Defines the screen point reader callback.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="point">point that supplies the point for this operation.</param>
        /// <returns>Boolean indicating the result of the check for operation on vbe native property tabs.</returns>
        internal delegate bool ScreenPointReader(IntPtr window, ref Point point);

        /// <summary>Defines the item reader callback.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="message">uint that supplies the message for this operation.</param>
        /// <param name="index">Native handle that supplies the index for this operation.</param>
        /// <param name="item">tab item that supplies the item for this operation.</param>
        /// <returns>int ptr produced by the operation for operation on vbe native property tabs.</returns>
        internal delegate IntPtr ItemReader(IntPtr window, uint message, IntPtr index, ref TabItem item);

        /// <summary>Defines the item rect reader callback.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="message">uint that supplies the message for this operation.</param>
        /// <param name="index">Native handle that supplies the index for this operation.</param>
        /// <param name="rectangle">rect that supplies the rectangle for this operation.</param>
        /// <returns>int ptr produced by the operation for operation on vbe native property tabs.</returns>
        internal delegate IntPtr ItemRectReader(IntPtr window, uint message, IntPtr index, out Rect rectangle);

        /// <summary>Defines the paint beginner callback.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="state">paint state that supplies the state for this operation.</param>
        /// <returns>int ptr produced by the operation for operation on vbe native property tabs.</returns>
        internal delegate IntPtr PaintBeginner(IntPtr window, out PaintState state);

        /// <summary>Defines the paint ender callback.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="state">paint state that supplies the state for this operation.</param>
        /// <returns>Boolean indicating the result of the check for operation on vbe native property tabs.</returns>
        internal delegate bool PaintEnder(IntPtr window, ref PaintState state);

        /// <summary>Defines the mouse tracker callback.</summary>
        /// <param name="state">track mouse that supplies the state for this operation.</param>
        /// <returns>Boolean indicating the result of the check for operation on vbe native property tabs.</returns>
        internal delegate bool MouseTracker(ref TrackMouse state);

        /// <summary>Maintains the class name state for vbe native property tabs.</summary>
        internal static Func<IntPtr, StringBuilder, int, int> ClassName = GetClassName;

        /// <summary>Maintains the style state for vbe native property tabs.</summary>
        internal static Func<IntPtr, int, int> Style = GetStyle;

        /// <summary>Maintains the window thread state for vbe native property tabs.</summary>
        internal static WindowThreadReader WindowThread = GetWindowThreadProcessId;

        /// <summary>Maintains the current thread state for vbe native property tabs.</summary>
        internal static Func<uint> CurrentThread = GetCurrentThreadId;

        /// <summary>Maintains the valid window and visible window and enabled window state for vbe native property tabs.</summary>
        internal static Func<IntPtr, bool> ValidWindow = IsWindow, VisibleWindow = IsWindowVisible, EnabledWindow = IsWindowEnabled;

        /// <summary>Maintains the focus state for vbe native property tabs.</summary>
        internal static Func<IntPtr> Focus = GetFocus;

        /// <summary>Maintains the client bounds and window bounds state for vbe native property tabs.</summary>
        internal static RectReader ClientBounds = GetClientRect, WindowBounds = GetWindowRect;

        /// <summary>Maintains the screen point state for vbe native property tabs.</summary>
        internal static ScreenPointReader ScreenPoint = ClientToScreen;

        /// <summary>Maintains the related window state for vbe native property tabs.</summary>
        internal static Func<IntPtr, uint, IntPtr> RelatedWindow = GetWindow;

        /// <summary>Maintains the send message state for vbe native property tabs.</summary>
        internal static Func<IntPtr, uint, IntPtr, IntPtr, IntPtr> SendMessage = Send;

        /// <summary>Maintains the read item state for vbe native property tabs.</summary>
        internal static ItemReader ReadItem = SendItem;

        /// <summary>Maintains the read item rect state for vbe native property tabs.</summary>
        internal static ItemRectReader ReadItemRect = SendRect;

        /// <summary>Maintains the start paint state for vbe native property tabs.</summary>
        internal static PaintBeginner StartPaint = BeginPaint;

        /// <summary>Maintains the finish paint state for vbe native property tabs.</summary>
        internal static PaintEnder FinishPaint = EndPaint;

        /// <summary>Maintains the invalidate state for vbe native property tabs.</summary>
        internal static Func<IntPtr, IntPtr, bool, bool> Invalidate = InvalidateRect;

        /// <summary>Maintains the update state for vbe native property tabs.</summary>
        internal static Func<IntPtr, bool> Update = UpdateWindow;

        /// <summary>Maintains the track state for vbe native property tabs.</summary>
        internal static MouseTracker Track = TrackMouseEvent;

        /// <summary>Creates a renderer tied to the thread that owns the native tab control.</summary>
        /// <param name="window">Native tab-control handle.</param>
        private VbeNativePropertyTabs(IntPtr window)
        {
            this.window = window;
            ownerThread = WindowThread(window, out _);
        }

        /// <summary>Accepts only two text-only horizontal native tabs on their owning thread.</summary>
        /// <param name="window">Candidate native tab-control handle.</param>
        /// <returns><see langword="true"/> when its styles, geometry, and thread meet the renderer contract.</returns>
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
        /// <param name="window">Candidate native tab-control handle.</param><param name="renderer">Receives the renderer when supported; otherwise null.</param>
        /// <returns><see langword="true"/> when the handle is supported and a renderer was created.</returns>
        internal static bool TryCreate(IntPtr window, out VbeNativePropertyTabs renderer)
        {
            renderer = CanRender(window) ? new VbeNativePropertyTabs(window) : null;
            return renderer != null;
        }

        /// <summary>Handles painting before the original window procedure; other messages remain native.</summary>
        /// <param name="message">Windows message identifier.</param><param name="wParam">First message value, including the print device context.</param>
        /// <param name="lParam">Second message value, including WM_PRINT flags.</param><param name="result">Receives the handled message result.</param>
        /// <returns><see langword="true"/> when the renderer handled the message.</returns>
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
        /// <param name="message">Native message that has just completed.</param><param name="wParam">Message-specific first value.</param>
        /// <param name="lParam">Message-specific second value; mouse coordinates for pointer messages.</param>
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

        /// <summary>Reads native tab geometry and state into a snapshot while on the owning UI thread.</summary>
        /// <param name="state">Receives the snapshot when the control remains within the supported contract.</param>
        /// <returns><see langword="true"/> when all required tab data was read successfully.</returns>
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

        /// <summary>Draws the background, tab borders, labels, selection accent, and focus cue.</summary>
        /// <param name="dc">Device context to paint into.</param><param name="state">Current tab snapshot.</param>
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

        /// <summary>Fills the tab client rectangle without drawing tabs or labels.</summary>
        /// <param name="dc">Device context to paint into.</param><param name="client">Client bounds.</param>
        private static void DrawBackground(IntPtr dc, Rect client)
        {
            int saved = SaveDC(dc);
            if (saved == 0) return;
            try { Fill(dc, client, Color(32, 36, 43)); }
            finally { RestoreDC(dc, saved); }
        }

        /// <summary>Packs RGB channel values into the COLORREF byte order used by GDI.</summary>
        /// <param name="red">Red channel from 0 to 255.</param><param name="green">Green channel from 0 to 255.</param><param name="blue">Blue channel from 0 to 255.</param>
        /// <returns>Packed RGB color.</returns>
        private static uint Color(int red, int green, int blue) => (uint)(red | green << 8 | blue << 16);

        /// <summary>Fills a rectangle using the stock DC brush after setting its color.</summary>
        /// <param name="dc">Target device context.</param><param name="bounds">Rectangle to fill.</param><param name="color">COLORREF fill color.</param>
        private static void Fill(IntPtr dc, Rect bounds, uint color)
        {
            SetDCBrushColor(dc, color);
            FillRect(dc, ref bounds, GetStockObject(18));
        }

        /// <summary>Frames a rectangle using the stock DC brush after setting its color.</summary>
        /// <param name="dc">Target device context.</param><param name="bounds">Rectangle to frame.</param><param name="color">COLORREF frame color.</param>
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

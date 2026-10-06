using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Shared rounded input outlines; native editing, scrolling and accessibility are preserved.</summary>
    internal static class UiInputFrame
    {

        /// <summary>Acquires the device context for the entire native window.</summary>
        /// <param name="window">Target native window handle.</param>
        /// <returns>A device-context handle covering the window and its non-client border, or zero on failure.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetWindowDC(IntPtr window);

        /// <summary>Releases a device context obtained for the native window.</summary>
        /// <param name="window">Target native window handle.</param>
        /// <param name="dc">Native device context.</param>
        /// <returns>Nonzero when the device context was released.</returns>
        [DllImport("user32.dll")] private static extern int ReleaseDC(IntPtr window, IntPtr dc);

        /// <summary>The chat selector seven-DIP corner radius, shared by fields and lists.</summary>
        internal const int CornerRadius = 7;

        /// <summary>Builds a closed rounded path at the requested DPI.</summary>
        /// <param name="bounds">Available drawing rectangle.</param>
        /// <param name="dpi">Display density used to scale logical dimensions.</param>
        /// <returns>A new rounded vector path; the caller owns and must dispose it.</returns>
        internal static GraphicsPath Outline(Rectangle bounds, int dpi)
        {
            var path = new GraphicsPath();
            int diameter = Math.Max(1, Math.Min(2 * CornerRadius * dpi / 96, Math.Min(bounds.Width, bounds.Height)));
            path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure(); return path;
        }

        /// <summary>Masks outer corners and draws the focused or idle input boundary.</summary>
        /// <param name="graphics">Drawing context; ownership remains with the caller.</param>
        /// <param name="control">Control whose native palette and geometry are used.</param>
        internal static void DrawOutline(Graphics graphics, Control control)
        {
            var saved = graphics.Save();
            try
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var path = Outline(new Rectangle(0, 0, control.Width - 1, control.Height - 1), control.DeviceDpi))
                {
                    // Only mask the outside corners. Leave the native client area, text,
                    // scrollbars, IME and accessibility untouched.
                    using (var outside = new Region(new Rectangle(0, 0, control.Width, control.Height)))
                    {
                        outside.Exclude(path);
                        using (var fill = new SolidBrush(control.Parent?.BackColor ?? UiTheme.Background)) graphics.FillRegion(fill, outside);
                    }
                    using (var pen = new Pen(Border(control))) graphics.DrawPath(pen, path);
                }
            }
            finally { graphics.Restore(saved); }
        }

        /// <summary>Selects the input boundary color from its focus state.</summary>
        /// <param name="control">Control whose native palette and geometry are used.</param>
        /// <returns>Theme border color, switched to the focus border while the control contains focus.</returns>
        internal static Color Border(Control control) => control.ContainsFocus ? UiTheme.FocusBorderFor(control) : UiTheme.BorderFor(control);

        /// <summary>Draws an input outline for native paint, focus and print messages.</summary>
        /// <param name="control">Control whose native palette and geometry are used.</param>
        /// <param name="style">Current native border style; no custom outline is drawn for <see cref="BorderStyle.None"/>.</param>
        /// <param name="message">Window message that triggered painting; only non-client, paint, focus, and print messages are handled.</param>
        internal static void Paint(Control control, BorderStyle style, Message message)
        {
            if (style == BorderStyle.None || !control.IsHandleCreated || control.Width < 2 || control.Height < 2) return;
            if (message.Msg != 0x85 && message.Msg != 0xF && message.Msg != 7 && message.Msg != 8 && message.Msg != 0xA && message.Msg != 0x317) return;
            bool printing = message.Msg == 0x317 && message.WParam != IntPtr.Zero;
            IntPtr dc = printing ? message.WParam : GetWindowDC(control.Handle);
            if (dc == IntPtr.Zero) return;
            try
            {
                using (var graphics = Graphics.FromHdc(dc)) DrawOutline(graphics, control);
            }
            finally { if (!printing) ReleaseDC(control.Handle, dc); }
        }
    }

    /// <summary>Native text input with the common field border.</summary>
    [ToolboxItem(true)]
    public class UiTextBox : TextBox
    {

        /// <summary>Creates a text field.</summary>
        public UiTextBox() { BorderStyle = BorderStyle.FixedSingle; }

        /// <summary>Sends a native edit-control message to the text box window.</summary>
        /// <param name="window">Target native window handle.</param>
        /// <param name="message">Edit-control message identifier.</param>
        /// <param name="wParam">First message-specific value.</param>
        /// <param name="lParam">Second message-specific value.</param>
        /// <returns>The native window-procedure result.</returns>
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

        /// <inheritdoc/>
        /// <summary>Applies DPI-scaled horizontal text insets after the native edit handle is created.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            int inset = 6 * DeviceDpi / 96;
            SendMessage(Handle, 0xD3, new IntPtr(3), new IntPtr(inset | (inset << 16)));
        }

        /// <inheritdoc/>
        /// <summary>Preserves native message handling and repaints the input surface and outline.</summary>
        /// <param name="m">Native input or paint message passed to the base edit control first.</param>
        protected override void WndProc(ref Message m) { base.WndProc(ref m); UiInputFrame.Paint(this, BorderStyle, m); }
    }

    /// <summary>Native rich text field with the common field border.</summary>
    [ToolboxItem(true)]
    public class UiRichTextBox : RichTextBox
    {

        /// <summary>Creates a rich text field.</summary>
        public UiRichTextBox() { BorderStyle = BorderStyle.FixedSingle; }

        /// <inheritdoc/>
        /// <summary>Preserves native message handling and repaints the input surface and outline.</summary>
        /// <param name="m">Native rich-edit message passed to the base control before its custom border is painted.</param>
        protected override void WndProc(ref Message m) { base.WndProc(ref m); UiInputFrame.Paint(this, BorderStyle, m); }
    }

    /// <summary>Native list with the common field border.</summary>
    [ToolboxItem(true)]
    public class UiListBox : ListBox
    {

        /// <summary>Suppresses recursive cache repair while synchronizing a list selection changed by an external UIA caller.</summary>
        private bool synchronizingExternalSelection;

        /// <summary>Identifies whether the current window message arrived from another process.</summary>
        /// <param name="reserved">Reserved native parameter, passed as zero for this query.</param>
        /// <returns>Native flags describing the sending message; zero means no cross-thread message is pending.</returns>
        [DllImport("user32.dll")]
        private static extern uint InSendMessageEx(IntPtr reserved);

        /// <summary>Sends a native list-box message and returns its control-specific result.</summary>
        /// <param name="window">List-box HWND.</param><param name="message">LB_* message identifier.</param>
        /// <param name="wParam">First message-specific value.</param><param name="lParam">Second message-specific value.</param>
        /// <returns>Native message result.</returns>
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

        /// <summary>Creates a list.</summary>
        public UiListBox() { BorderStyle = BorderStyle.FixedSingle; }

        /// <inheritdoc/>
        /// <summary>Preserves native message handling and repaints the input surface and outline.</summary>
        /// <param name="m">Native list-box message; externally sent selection changes are reflected in WinForms' selected-item cache.</param>
        protected override void WndProc(ref Message m)
        {
            // UIA's native list provider sends LB_SETCURSEL from another
            // thread without the LBN_SELCHANGE notification that WinForms
            // needs to invalidate its SelectedItems cache. Same-thread
            // managed setters already raise this event themselves.
            bool externalSelection = !synchronizingExternalSelection && m.Msg == 0x186 && SelectionMode == SelectionMode.One &&
                InSendMessageEx(IntPtr.Zero) != 0;
            int previous = externalSelection ? SelectedIndex : -1;
            bool externalMultiSelection = !synchronizingExternalSelection && m.Msg == 0x185 &&
                (SelectionMode == SelectionMode.MultiSimple || SelectionMode == SelectionMode.MultiExtended) &&
                InSendMessageEx(IntPtr.Zero) != 0 && m.LParam.ToInt64() >= 0 && m.LParam.ToInt64() < Items.Count;
            int multiIndex = externalMultiSelection ? m.LParam.ToInt32() : -1;
            // LB_GETSEL reads the native bit before and after LB_SETSEL; the
            // materialized SelectedItems cache can still hold the old bit.
            int previousMultiState = externalMultiSelection ? SendMessage(Handle, 0x187, new IntPtr(multiIndex), IntPtr.Zero).ToInt32() : -1;
            base.WndProc(ref m);
            if (externalSelection && previous != SelectedIndex)
            {
                // The public managed selection APIs update the cached item flags
                // as well as raising the event. Raising the event alone leaves
                // .NET Framework's already materialized cache unchanged.
                synchronizingExternalSelection = true;
                try
                {
                    int selected = SelectedIndex;
                    if (selected < 0) ClearSelected();
                    else SetSelected(selected, true);
                }
                finally { synchronizingExternalSelection = false; }
            }
            if (externalMultiSelection && previousMultiState >= 0)
            {
                int selected = SendMessage(Handle, 0x187, new IntPtr(multiIndex), IntPtr.Zero).ToInt32();
                if (selected >= 0 && selected != previousMultiState)
                {
                    synchronizingExternalSelection = true;
                    try { SetSelected(multiIndex, selected != 0); }
                    finally { synchronizingExternalSelection = false; }
                }
            }
            UiInputFrame.Paint(this, BorderStyle, m);
        }
    }

    /// <summary>Native checked list with the common field border.</summary>
    [ToolboxItem(true)]
    public class UiCheckedListBox : CheckedListBox
    {

        /// <summary>Creates a checked list.</summary>
        public UiCheckedListBox() { BorderStyle = BorderStyle.FixedSingle; }

        /// <inheritdoc/>
        /// <summary>Preserves native message handling and repaints the input surface and outline.</summary>
        /// <param name="m">Native checked-list message passed to the base control before its custom border is painted.</param>
        protected override void WndProc(ref Message m) { base.WndProc(ref m); UiInputFrame.Paint(this, BorderStyle, m); }
    }

    /// <summary>Native data grid with the common field border.</summary>
    [ToolboxItem(true)]
    public class UiDataGridView : DataGridView
    {

        /// <summary>Creates a data grid.</summary>
        public UiDataGridView() { BorderStyle = BorderStyle.FixedSingle; }

        /// <inheritdoc/>
        /// <summary>Preserves native message handling and repaints the input surface and outline.</summary>
        /// <param name="m">Native grid message passed to the base control before its custom border is painted.</param>
        protected override void WndProc(ref Message m) { base.WndProc(ref m); UiInputFrame.Paint(this, BorderStyle, m); }
    }

    /// <summary>Shared selector for chat, Git and settings; editable dropdowns keep their native edit control.</summary>
    [ToolboxItem(true)]
    public class UiComboBox : ComboBox
    {

        /// <summary>Carries the native rect values passed between operations.</summary>
        [StructLayout(LayoutKind.Sequential)] private struct NativeRect {

/// <summary>Edges of a native client, item, or button rectangle, in device coordinates.</summary>
public int Left, Top, Right, Bottom; }

        /// <summary>Carries the combo info values passed between operations.</summary>
        [StructLayout(LayoutKind.Sequential)] private struct ComboInfo
        {

            /// <summary>Size of this interop structure in bytes, initialized before querying the combo box.</summary>
            public int Size;

/// <summary>Native bounds of the displayed item and dropdown button.</summary>
public NativeRect Item, Button;

/// <summary>Native state flags for the dropdown button.</summary>
public int ButtonState;

            /// <summary>HWNDs for the combo itself and its optional edit and list children.</summary>
            public IntPtr Combo, Edit, List;
        }

        /// <summary>Reads the native selector and editable child bounds.</summary>
        /// <param name="combo">Combo-box HWND whose child handles and bounds are requested.</param>
        /// <param name="info">Structure initialized with its native size; receives child handles and item/button rectangles.</param>
        /// <returns><see langword="true"/> when Windows fills the structure.</returns>
        [DllImport("user32.dll")] private static extern bool GetComboBoxInfo(IntPtr combo, ref ComboInfo info);

        /// <summary>Cached GDI brush returned for native edit/list background-color messages; disposed on color change and control disposal.</summary>
        private IntPtr inputBrush;

        /// <summary>Creates a GDI brush for the native input background.</summary>
        /// <param name="color">Win32 COLORREF value used as the brush fill.</param>
        /// <returns>Owned brush handle, or zero if allocation fails.</returns>
        [DllImport("gdi32.dll")] private static extern IntPtr CreateSolidBrush(int color);

        /// <summary>Releases a previously allocated GDI object.</summary>
        /// <param name="value">GDI brush handle allocated by this control.</param>
        /// <returns><see langword="true"/> when Windows deletes the object.</returns>
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr value);

        /// <summary>Sets the native device context background color.</summary>
        /// <param name="dc">Native device context.</param>
        /// <param name="color">COLORREF background value.</param>
        /// <returns>Previous background COLORREF.</returns>
        [DllImport("gdi32.dll")] private static extern int SetBkColor(IntPtr dc, int color);

        /// <summary>Sets the native device context text color.</summary>
        /// <param name="dc">Native device context.</param>
        /// <param name="color">COLORREF text value.</param>
        /// <returns>Previous text COLORREF.</returns>
        [DllImport("gdi32.dll")] private static extern int SetTextColor(IntPtr dc, int color);

        /// <inheritdoc/>
        /// <summary>Releases the cached native brush when its background color changes.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnBackColorChanged(EventArgs e)
        {
            if (inputBrush != IntPtr.Zero) { DeleteObject(inputBrush); inputBrush = IntPtr.Zero; }
            base.OnBackColorChanged(e);
        }

        /// <inheritdoc/>
        /// <summary>Releases the resources owned by this control before base disposal.</summary>
        /// <param name="disposing">Whether managed resources must also be released.</param>
        protected override void Dispose(bool disposing)
        {
            if (inputBrush != IntPtr.Zero) { DeleteObject(inputBrush); inputBrush = IntPtr.Zero; }
            base.Dispose(disposing);
        }

        /// <summary>Creates a selector with consistent list spacing.</summary>
        public UiComboBox() { FlatStyle = FlatStyle.Standard; DrawMode = DrawMode.OwnerDrawFixed; ItemHeight = Math.Max(Font.Height, 22 * DeviceDpi / 96); }

        /// <inheritdoc/>
        /// <summary>Updates the native item height for the current font and DPI.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnFontChanged(EventArgs e) { base.OnFontChanged(e); ItemHeight = Math.Max(Font.Height, 22 * DeviceDpi / 96); }

        /// <inheritdoc/>
        /// <summary>Repaints the focus outline after the control gains focus.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }

        /// <inheritdoc/>
        /// <summary>Repaints the focus outline after the control loses focus.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }

        /// <inheritdoc/>
        /// <summary>Repaints the current selection after its index changes.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnSelectedIndexChanged(EventArgs e) { base.OnSelectedIndexChanged(e); Invalidate(); }

        /// <inheritdoc/>
        /// <summary>Draws one owner-drawn item with selection, RTL alignment, ellipsis, and focus state.</summary>
        /// <param name="e">Item bounds, index, state flags, and drawing context supplied by WinForms.</param>
        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            bool selected = (e.State & DrawItemState.Selected) != 0;
            using (var fill = new SolidBrush(selected ? SystemColors.Highlight : BackColor)) e.Graphics.FillRectangle(fill, e.Bounds);
            var bounds = Rectangle.Inflate(e.Bounds, -6 * DeviceDpi / 96, 0);
            TextRenderer.DrawText(e.Graphics, GetItemText(Items[e.Index]), Font, bounds, selected ? SystemColors.HighlightText : ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | (RightToLeft == RightToLeft.Yes ? TextFormatFlags.RightToLeft | TextFormatFlags.Right : 0));
            e.DrawFocusRectangle();
        }

        /// <inheritdoc/>
        /// <summary>Preserves native message handling and repaints the input surface and outline.</summary>
        /// <param name="m">Native paint or control-color message processed before custom selector rendering.</param>
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if ((m.Msg == 0x133 || m.Msg == 0x138) && m.WParam != IntPtr.Zero)
            {
                // The editable child also needs the palette when disabled; otherwise
                // Windows paints a light textbox inside the dark selector.
                if (inputBrush == IntPtr.Zero) inputBrush = CreateSolidBrush(ColorTranslator.ToWin32(BackColor));
                SetBkColor(m.WParam, ColorTranslator.ToWin32(BackColor));
                SetTextColor(m.WParam, ColorTranslator.ToWin32(Enabled ? ForeColor : SystemColors.GrayText));
                m.Result = inputBrush;
                return;
            }

            if ((m.Msg != 0xF && m.Msg != 0x317 && m.Msg != 0x318) || Width < 2 || Height < 2) return;
            using (var graphics = m.Msg == 0xF || m.WParam == IntPtr.Zero ? Graphics.FromHwnd(Handle) : Graphics.FromHdc(m.WParam))
            {
                bool rtl = RightToLeft == RightToLeft.Yes;
                int arrowWidth = SystemInformation.VerticalScrollBarWidth;
                var arrow = new Rectangle(rtl ? 1 : Width - arrowWidth - 1, 1, arrowWidth, Height - 2);
                using (var fill = new SolidBrush(BackColor))
                {
                    if (DropDownStyle == ComboBoxStyle.DropDownList)
                    {
                        graphics.FillRectangle(fill, ClientRectangle);
                        var text = new Rectangle(rtl ? arrowWidth + 4 : 6, 0, Math.Max(1, Width - arrowWidth - 10), Height);
                        TextRenderer.DrawText(graphics, SelectedIndex < 0 ? Text : GetItemText(SelectedItem), Font, text, Enabled ? ForeColor : SystemColors.GrayText,
                            TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | (rtl ? TextFormatFlags.RightToLeft | TextFormatFlags.Right : 0));
                    }
                    if (DropDownStyle != ComboBoxStyle.Simple) graphics.FillRectangle(fill, arrow);
                }
                if (DropDownStyle != ComboBoxStyle.Simple)
                {
                    int x = arrow.Left + arrow.Width / 2, y = arrow.Top + arrow.Height / 2;
                    using (var pen = new Pen(Enabled ? ForeColor : SystemColors.GrayText))
                        graphics.DrawLines(pen, new[] { new Point(x - 3, y - 1), new Point(x, y + 2), new Point(x + 3, y - 1) });
                }
                if (DropDownStyle != ComboBoxStyle.DropDownList)
                {
                    var info = new ComboInfo { Size = Marshal.SizeOf(typeof(ComboInfo)) };
                    if (GetComboBoxInfo(Handle, ref info))
                    using (var erase = new Pen(BackColor, 3))
                        graphics.DrawRectangle(erase, info.Item.Left - 2, info.Item.Top - 2,
                            info.Item.Right - info.Item.Left + 4, info.Item.Bottom - info.Item.Top + 4);
                }
                UiInputFrame.DrawOutline(graphics, this);
            }
        }
    }
}

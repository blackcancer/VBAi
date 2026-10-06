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
        /// <returns>int ptr produced by the operation for get window dc on ui input frame.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetWindowDC(IntPtr window);

        /// <summary>Releases a device context obtained for the native window.</summary>
        /// <param name="window">Target native window handle.</param>
        /// <param name="dc">Native device context.</param>
        /// <returns>int produced by the operation for release dc on ui input frame.</returns>
        [DllImport("user32.dll")] private static extern int ReleaseDC(IntPtr window, IntPtr dc);

        /// <summary>The chat selector seven-DIP corner radius, shared by fields and lists.</summary>
        internal const int CornerRadius = 7;

        /// <summary>Builds a closed rounded path at the requested DPI.</summary>
        /// <param name="bounds">Available drawing rectangle.</param>
        /// <param name="dpi">Display density used to scale logical dimensions.</param>
        /// <returns>graphics path produced by the operation for outline on ui input frame.</returns>
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
        /// <returns>color produced by the operation for border on ui input frame.</returns>
        internal static Color Border(Control control) => control.ContainsFocus ? UiTheme.FocusBorderFor(control) : UiTheme.BorderFor(control);

        /// <summary>Draws an input outline for native paint, focus and print messages.</summary>
        /// <param name="control">Control whose native palette and geometry are used.</param>
        /// <param name="style">border style that supplies the style for this operation.</param>
        /// <param name="message">message that supplies the message for this operation.</param>
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

        /// <summary>Handles send message for ui text box.</summary>
        /// <param name="window">Target native window handle.</param>
        /// <param name="message">int that supplies the message for this operation.</param>
        /// <param name="wParam">Native handle that supplies the w param for this operation.</param>
        /// <param name="lParam">Native handle that supplies the l param for this operation.</param>
        /// <returns>int ptr produced by the operation for send message on ui text box.</returns>
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

        /// <inheritdoc/>
        /// <summary>Handles on handle created for ui text box.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            int inset = 6 * DeviceDpi / 96;
            SendMessage(Handle, 0xD3, new IntPtr(3), new IntPtr(inset | (inset << 16)));
        }

        /// <inheritdoc/>
        /// <summary>Preserves native message handling and repaints the input surface and outline.</summary>
        /// <param name="m">message that supplies the m for this operation.</param>
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
        /// <param name="m">message that supplies the m for this operation.</param>
        protected override void WndProc(ref Message m) { base.WndProc(ref m); UiInputFrame.Paint(this, BorderStyle, m); }
    }

    /// <summary>Native list with the common field border.</summary>
    [ToolboxItem(true)]
    public class UiListBox : ListBox
    {

        /// <summary>Maintains the synchronizing external selection state for ui list box.</summary>
        private bool synchronizingExternalSelection;

        /// <summary>Handles in send message ex for ui list box.</summary>
        /// <param name="reserved">Native handle that supplies the reserved for this operation.</param>
        /// <returns>uint produced by the operation for in send message ex on ui list box.</returns>
        [DllImport("user32.dll")]
        private static extern uint InSendMessageEx(IntPtr reserved);

        /// <summary>Handles send message for ui list box.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="message">int that supplies the message for this operation.</param>
        /// <param name="wParam">Native handle that supplies the w param for this operation.</param>
        /// <param name="lParam">Native handle that supplies the l param for this operation.</param>
        /// <returns>int ptr produced by the operation for send message on ui list box.</returns>
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

        /// <summary>Creates a list.</summary>
        public UiListBox() { BorderStyle = BorderStyle.FixedSingle; }

        /// <inheritdoc/>
        /// <summary>Preserves native message handling and repaints the input surface and outline.</summary>
        /// <param name="m">message that supplies the m for this operation.</param>
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
        /// <param name="m">message that supplies the m for this operation.</param>
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
        /// <param name="m">message that supplies the m for this operation.</param>
        protected override void WndProc(ref Message m) { base.WndProc(ref m); UiInputFrame.Paint(this, BorderStyle, m); }
    }

    /// <summary>Shared selector for chat, Git and settings; editable dropdowns keep their native edit control.</summary>
    [ToolboxItem(true)]
    public class UiComboBox : ComboBox
    {

        /// <summary>Carries the native rect values passed between operations.</summary>
        [StructLayout(LayoutKind.Sequential)] private struct NativeRect {

/// <summary>Maintains the left and top and right and bottom state for native rect.</summary>
public int Left, Top, Right, Bottom; }

        /// <summary>Carries the combo info values passed between operations.</summary>
        [StructLayout(LayoutKind.Sequential)] private struct ComboInfo
        {

            /// <summary>Maintains the size state for combo info.</summary>
            public int Size;

/// <summary>Maintains the item and button state for combo info.</summary>
public NativeRect Item, Button;

/// <summary>Maintains the button state state for combo info.</summary>
public int ButtonState;

            /// <summary>Maintains the combo and edit and list state for combo info.</summary>
            public IntPtr Combo, Edit, List;
        }

        /// <summary>Reads the native selector and editable child bounds.</summary>
        /// <param name="combo">Native handle that supplies the combo for this operation.</param>
        /// <param name="info">combo info that supplies the info for this operation.</param>
        /// <returns>Boolean indicating the result of the check for get combo box info on ui combo box.</returns>
        [DllImport("user32.dll")] private static extern bool GetComboBoxInfo(IntPtr combo, ref ComboInfo info);

        /// <summary>Maintains the input brush state for ui combo box.</summary>
        private IntPtr inputBrush;

        /// <summary>Creates a GDI brush for the native input background.</summary>
        /// <param name="color">int that supplies the color for this operation.</param>
        /// <returns>int ptr produced by the operation for create solid brush on ui combo box.</returns>
        [DllImport("gdi32.dll")] private static extern IntPtr CreateSolidBrush(int color);

        /// <summary>Releases a previously allocated GDI object.</summary>
        /// <param name="value">Native handle that supplies the value for this operation.</param>
        /// <returns>Boolean indicating the result of the check for delete object on ui combo box.</returns>
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr value);

        /// <summary>Sets the native device context background color.</summary>
        /// <param name="dc">Native device context.</param>
        /// <param name="color">int that supplies the color for this operation.</param>
        /// <returns>int produced by the operation for set bk color on ui combo box.</returns>
        [DllImport("gdi32.dll")] private static extern int SetBkColor(IntPtr dc, int color);

        /// <summary>Sets the native device context text color.</summary>
        /// <param name="dc">Native device context.</param>
        /// <param name="color">int that supplies the color for this operation.</param>
        /// <returns>int produced by the operation for set text color on ui combo box.</returns>
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
        /// <summary>Handles on draw item for ui combo box.</summary>
        /// <param name="e">Native event data.</param>
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
        /// <param name="m">message that supplies the m for this operation.</param>
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

using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Shared rounded input outlines; native editing, scrolling and accessibility are preserved.</summary>
    internal static class UiInputFrame
    {
        /// <summary>Acquires the device context for the entire native window.</summary>
        /// <param name="window">Target native window handle.</param>
        /// <returns>The result produced by this operation.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetWindowDC(IntPtr window);
        /// <summary>Releases a device context obtained for the native window.</summary>
        /// <param name="window">Target native window handle.</param>
        /// <param name="dc">Native device context.</param>
        /// <returns>The result produced by this operation.</returns>
        [DllImport("user32.dll")] private static extern int ReleaseDC(IntPtr window, IntPtr dc);
        /// <summary>The chat selector seven-DIP corner radius, shared by fields and lists.</summary>
        internal const int CornerRadius = 7;
        /// <summary>Builds a closed rounded path at the requested DPI.</summary>
        /// <param name="bounds">Available drawing rectangle.</param>
        /// <param name="dpi">Display density used to scale logical dimensions.</param>
        /// <returns>The result produced by this operation.</returns>
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
        /// <returns>The result produced by this operation.</returns>
        internal static Color Border(Control control) => control.ContainsFocus ? UiTheme.FocusBorderFor(control) : UiTheme.BorderFor(control);
        /// <summary>Draws an input outline for native paint, focus and print messages.</summary>
        /// <param name="control">Control whose native palette and geometry are used.</param>
        /// <param name="style">The style used by this operation.</param>
        /// <param name="message">The message used by this operation.</param>
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
        /// <summary>Performs the send message operation for UiTextBox.</summary>
        /// <param name="window">Target native window handle.</param>
        /// <param name="message">The message used by this operation.</param>
        /// <param name="wParam">The w param used by this operation.</param>
        /// <param name="lParam">The l param used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);
        /// <inheritdoc/>
        /// <summary>Performs the on handle created operation for UiTextBox.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            int inset = 6 * DeviceDpi / 96;
            SendMessage(Handle, 0xD3, new IntPtr(3), new IntPtr(inset | (inset << 16)));
        }
        /// <inheritdoc/>
        /// <summary>Preserves native message handling and repaints the input surface and outline.</summary>
        /// <param name="m">The m used by this operation.</param>
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
        /// <param name="m">The m used by this operation.</param>
        protected override void WndProc(ref Message m) { base.WndProc(ref m); UiInputFrame.Paint(this, BorderStyle, m); }
    }
    /// <summary>Native list with the common field border.</summary>
    [ToolboxItem(true)]
    public class UiListBox : ListBox
    {
        /// <summary>Creates a list.</summary>
        public UiListBox() { BorderStyle = BorderStyle.FixedSingle; }
        /// <inheritdoc/>
        /// <summary>Preserves native message handling and repaints the input surface and outline.</summary>
        /// <param name="m">The m used by this operation.</param>
        protected override void WndProc(ref Message m) { base.WndProc(ref m); UiInputFrame.Paint(this, BorderStyle, m); }
    }
    /// <summary>Native checked list with the common field border.</summary>
    [ToolboxItem(true)]
    public class UiCheckedListBox : CheckedListBox
    {
        /// <summary>Creates a checked list.</summary>
        public UiCheckedListBox() { BorderStyle = BorderStyle.FixedSingle; }
        /// <inheritdoc/>
        /// <summary>Preserves native message handling and repaints the input surface and outline.</summary>
        /// <param name="m">The m used by this operation.</param>
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
        /// <param name="m">The m used by this operation.</param>
        protected override void WndProc(ref Message m) { base.WndProc(ref m); UiInputFrame.Paint(this, BorderStyle, m); }
    }
    /// <summary>Shared selector for chat, Git and settings; editable dropdowns keep their native edit control.</summary>
    [ToolboxItem(true)]
    public class UiComboBox : ComboBox
    {
        /// <summary>Represents native rect data.</summary>
        [StructLayout(LayoutKind.Sequential)] private struct NativeRect { /// <summary>Stores the left,top,right,bottom used by NativeRect.</summary>
public int Left, Top, Right, Bottom; }
        /// <summary>Represents combo info data.</summary>
        [StructLayout(LayoutKind.Sequential)] private struct ComboInfo
        {
            /// <summary>Stores the size used by ComboInfo.</summary>
            public int Size; /// <summary>Stores the item,button used by ComboInfo.</summary>
public NativeRect Item, Button; /// <summary>Stores the button state used by ComboInfo.</summary>
public int ButtonState;
            /// <summary>Stores the combo,edit,list used by ComboInfo.</summary>
            public IntPtr Combo, Edit, List;
        }
        /// <summary>Reads the native selector and editable child bounds.</summary>
        /// <param name="combo">The combo used by this operation.</param>
        /// <param name="info">The info used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        [DllImport("user32.dll")] private static extern bool GetComboBoxInfo(IntPtr combo, ref ComboInfo info);
        /// <summary>Stores the input brush used by UiComboBox.</summary>
        private IntPtr inputBrush;
        /// <summary>Creates a GDI brush for the native input background.</summary>
        /// <param name="color">The color used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        [DllImport("gdi32.dll")] private static extern IntPtr CreateSolidBrush(int color);
        /// <summary>Releases a previously allocated GDI object.</summary>
        /// <param name="value">The value used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr value);
        /// <summary>Sets the native device context background color.</summary>
        /// <param name="dc">Native device context.</param>
        /// <param name="color">The color used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        [DllImport("gdi32.dll")] private static extern int SetBkColor(IntPtr dc, int color);
        /// <summary>Sets the native device context text color.</summary>
        /// <param name="dc">Native device context.</param>
        /// <param name="color">The color used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
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
        /// <summary>Performs the on draw item operation for UiComboBox.</summary>
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
        /// <param name="m">The m used by this operation.</param>
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

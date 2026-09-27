using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace CodexVBE
{
    internal static class ChatControlPainting
    {
        public static GraphicsPath Rounded(Rectangle bounds, int radius)
        {
            var path = new GraphicsPath();
            int diameter = Math.Max(1, Math.Min(radius * 2, Math.Min(bounds.Width, bounds.Height)));
            path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure(); return path;
        }
    }

    // Native WinForms controls: standard properties, events and accessibility remain designer-editable.
    [ToolboxItem(true)]
    public class ChatActionButton : Button
    {
        private bool hovered, pressed;
        public ChatActionButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            FlatStyle = FlatStyle.Flat; FlatAppearance.BorderSize = 0;
            BackColor = Color.FromArgb(241, 245, 249); ForeColor = Color.FromArgb(51, 65, 85);
        }
        protected override void OnMouseEnter(EventArgs e) { hovered = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hovered = pressed = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { pressed = true; Invalidate(); base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { pressed = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            if (Width < 2 || Height < 2) return;
            using (var background = new SolidBrush(Parent?.BackColor ?? SystemColors.Control)) e.Graphics.FillRectangle(background, ClientRectangle);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var fill = !Enabled ? Color.FromArgb(241, 245, 249) : pressed ? ControlPaint.Dark(BackColor, .05f) : hovered ? ControlPaint.Light(BackColor, .12f) : BackColor;
            using (var path = ChatControlPainting.Rounded(new Rectangle(0, 0, Width - 1, Height - 1), 7))
            using (var brush = new SolidBrush(fill))
            {
                e.Graphics.FillPath(brush, path);
                if (Focused && ShowFocusCues) using (var pen = new Pen(Color.FromArgb(96, 165, 250))) e.Graphics.DrawPath(pen, path);
            }
            TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, Enabled ? ForeColor : SystemColors.GrayText,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }
    }

    [ToolboxItem(true)]
    public class ChatChoiceBox : ComboBox
    {
        public ChatChoiceBox()
        {
            DropDownStyle = ComboBoxStyle.DropDownList;
            DrawMode = DrawMode.OwnerDrawFixed; ItemHeight = 22;
            BackColor = Color.White; ForeColor = Color.FromArgb(51, 65, 85);
        }
        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            bool selected = (e.State & DrawItemState.Selected) != 0;
            using (var brush = new SolidBrush(selected ? Color.FromArgb(239, 246, 255) : BackColor)) e.Graphics.FillRectangle(brush, e.Bounds);
            var bounds = new Rectangle(e.Bounds.X + 7, e.Bounds.Y, Math.Max(1, e.Bounds.Width - 14), e.Bounds.Height);
            TextRenderer.DrawText(e.Graphics, GetItemText(Items[e.Index]), Font, bounds, ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
        }
        protected override void OnSelectedIndexChanged(EventArgs e) { base.OnSelectedIndexChanged(e); Invalidate(); }
        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if ((m.Msg == 0x000F || m.Msg == 0x0317 || m.Msg == 0x0318) && Width > 2 && Height > 2)
            {
                using (var graphics = m.Msg == 0x000F ? Graphics.FromHwnd(Handle) : Graphics.FromHdc(m.WParam))
                {
                    using (var background = new SolidBrush(Parent?.BackColor ?? SystemColors.Control)) graphics.FillRectangle(background, ClientRectangle);
                    graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    using (var path = ChatControlPainting.Rounded(new Rectangle(0, 0, Width - 1, Height - 1), 7))
                    using (var fill = new SolidBrush(BackColor))
                    using (var border = new Pen(Focused ? Color.FromArgb(96, 165, 250) : Color.FromArgb(226, 232, 240)))
                    { graphics.FillPath(fill, path); graphics.DrawPath(border, path); }
                    TextRenderer.DrawText(graphics, SelectedIndex < 0 ? Text : GetItemText(SelectedItem), Font,
                        new Rectangle(9, 0, Math.Max(1, Width - 32), Height), Enabled ? ForeColor : SystemColors.GrayText,
                        TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
                    using (var pen = new Pen(Color.FromArgb(100, 116, 139), 1.3f))
                        graphics.DrawLines(pen, new[] { new Point(Width - 19, Height / 2 - 2), new Point(Width - 15, Height / 2 + 2), new Point(Width - 11, Height / 2 - 2) });
                }
            }
        }
    }

    [ToolboxItem(true)]
    public class ChatComposerPanel : TableLayoutPanel
    {
        public ChatComposerPanel() { DoubleBuffered = true; BackColor = Color.White; }
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            if (Width < 2 || Height < 2) return;
            using (var background = new SolidBrush(Parent?.BackColor ?? SystemColors.Control)) e.Graphics.FillRectangle(background, ClientRectangle);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = ChatControlPainting.Rounded(new Rectangle(0, 0, Width - 1, Height - 1), 12))
            using (var fill = new SolidBrush(BackColor))
            using (var pen = new Pen(Color.FromArgb(203, 213, 225)))
            { e.Graphics.FillPath(fill, path); e.Graphics.DrawPath(pen, path); }
        }
    }
}

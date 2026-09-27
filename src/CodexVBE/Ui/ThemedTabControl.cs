using System.Drawing;
using System.Windows.Forms;
namespace CodexVBE
{
    public sealed class ThemedTabControl : TabControl
    {
        public ThemedTabControl() { SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true); }
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(UiTheme.Background);
            for (int i = 0; i < TabPages.Count; i++)
            {
                var bounds = GetTabRect(i); bool selected = i == SelectedIndex;
                using (var brush = new SolidBrush(selected ? UiTheme.Surface : UiTheme.Background)) e.Graphics.FillRectangle(brush, bounds);
                Color ink = TabPages[i].Enabled ? ForeColor : (UiTheme.Dark ? Color.FromArgb(148, 163, 184) : SystemColors.GrayText);
                TextRenderer.DrawText(e.Graphics, TabPages[i].Text, Font, bounds, ink, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
                if (selected) using (var pen = new Pen(UiTheme.Dark ? Color.FromArgb(96, 165, 250) : Color.RoyalBlue, 2)) e.Graphics.DrawLine(pen, bounds.Left + 3, bounds.Bottom - 2, bounds.Right - 3, bounds.Bottom - 2);
                if (selected && Focused) ControlPaint.DrawFocusRectangle(e.Graphics, Rectangle.Inflate(bounds, -3, -3), ink, UiTheme.Surface);
            }
        }
        protected override void OnSelectedIndexChanged(System.EventArgs e) { base.OnSelectedIndexChanged(e); Invalidate(); }
    }
    public sealed class ThemedButton : Button
    {
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (!Enabled && UiTheme.Dark)
            {
                var content = Rectangle.Inflate(ClientRectangle, -2, -2);
                using (var brush = new SolidBrush(BackColor)) e.Graphics.FillRectangle(brush, content);
                TextRenderer.DrawText(e.Graphics, Text, Font, content, Color.FromArgb(148, 163, 184), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            }
        }
    }
    public sealed class ThemedComboBox : ComboBox
    {
        protected override void WndProc(ref Message message)
        {
            base.WndProc(ref message);
            if ((message.Msg == 0x000F || message.Msg == 0x0318) && UiTheme.Dark && !UiTheme.HighContrast())
            {
                using (var graphics = message.Msg == 0x0318 ? Graphics.FromHdc(message.WParam) : CreateGraphics())
                {
                    int width = SystemInformation.VerticalScrollBarWidth;
                    var button = new Rectangle(RightToLeft == RightToLeft.Yes ? 1 : Width - width - 1, 1, width, Height - 2);
                    using (var brush = new SolidBrush(UiTheme.Surface)) graphics.FillRectangle(brush, button);
                    int x = button.Left + button.Width / 2, y = button.Top + button.Height / 2;
                    using (var brush = new SolidBrush(Enabled ? UiTheme.Foreground : SystemColors.GrayText))
                        graphics.FillPolygon(brush, new[] { new Point(x - 4, y - 2), new Point(x + 4, y - 2), new Point(x, y + 2) });
                    using (var pen = new Pen(Color.FromArgb(75, 85, 99))) graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
                }
            }
        }
    }

}

using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
namespace CodexVBE
{
    /// <summary>Native Designer-editable tabs using the chat's rounded surfaces and focus states.</summary>
    public sealed class ThemedTabControl : TabControl
    {
        private int hoveredTab = -1;
        private bool closeHovered;
        /// <summary>Displays a close command on each document tab.</summary>
        [DefaultValue(false)]
        public bool ShowCloseButtons { get; set; }
        /// <summary>Requests closure of a document without selecting a different tab.</summary>
        public event EventHandler<TabControlEventArgs> CloseRequested;
        /// <summary>Creates the shared tab appearance without replacing the native tab model.</summary>
        public ThemedTabControl()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            ItemSize = new Size(0, 36);
            Padding = new Point(12, 6);
        }
        private Rectangle CloseBounds(int index)
        {
            var bounds = GetTabRect(index); int size = 22 * DeviceDpi / 96;
            return new Rectangle(bounds.Right - size - 6 * DeviceDpi / 96, bounds.Top + (bounds.Height - size) / 2, size, size);
        }
        /// <inheritdoc/>
        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (ShowCloseButtons && e.Button == MouseButtons.Left)
                for (int i = 0; i < TabCount; i++)
                    if (CloseBounds(i).Contains(e.Location)) { CloseRequested?.Invoke(this, new TabControlEventArgs(TabPages[i], i, TabControlAction.Deselecting)); return; }
            base.OnMouseDown(e);
        }
        /// <inheritdoc/>
        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            int index = -1;
            for (int i = 0; i < TabCount; i++) if (GetTabRect(i).Contains(e.Location)) { index = i; break; }
            bool close = index >= 0 && ShowCloseButtons && CloseBounds(index).Contains(e.Location);
            if (hoveredTab != index || closeHovered != close) { hoveredTab = index; closeHovered = close; Invalidate(); }
            Cursor = close ? Cursors.Hand : Cursors.Default;
        }
        /// <inheritdoc/>
        protected override void OnMouseLeave(EventArgs e) { hoveredTab = -1; closeHovered = false; Cursor = Cursors.Default; Invalidate(); base.OnMouseLeave(e); }
        /// <inheritdoc/>
        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        /// <inheritdoc/>
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }
        /// <inheritdoc/>
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(UiTheme.BackgroundFor(this));
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            int inset = Math.Max(2, 3 * DeviceDpi / 96);
            for (int i = 0; i < TabPages.Count; i++)
            {
                var bounds = GetTabRect(i);
                var face = Rectangle.Inflate(bounds, -inset, -inset);
                if (face.Width < 2 || face.Height < 2) continue;
                bool selected = i == SelectedIndex, hovered = i == hoveredTab && TabPages[i].Enabled;
                if (selected || hovered)
                using (var path = UiInputFrame.Outline(face, DeviceDpi))
                {
                    using (var fill = new SolidBrush(UiTheme.SurfaceFor(this))) e.Graphics.FillPath(fill, path);
                    using (var pen = new Pen(selected && Focused && ShowFocusCues ? UiTheme.FocusBorderFor(this) : UiTheme.BorderFor(this))) e.Graphics.DrawPath(pen, path);
                }
                Color ink = TabPages[i].Enabled ? UiTheme.ForegroundFor(this) : SystemColors.GrayText;
                var text = Rectangle.Inflate(face, -6 * DeviceDpi / 96, 0);
                if (ShowCloseButtons)
                {
                    var close = CloseBounds(i);
                    text.Width = Math.Max(0, close.Left - text.Left - 2);
                    if (closeHovered && hovered)
                    using (var path = UiInputFrame.Outline(close, DeviceDpi))
                    using (var fill = new SolidBrush(UiTheme.BackgroundFor(this))) e.Graphics.FillPath(fill, path);
                    UiCommandIcons.Draw(e.Graphics, UiSymbol.Close, close, ink, DeviceDpi);
                }
                TextRenderer.DrawText(e.Graphics, TabPages[i].Text, Font, text, ink,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            }
        }
        /// <inheritdoc/>
        protected override void OnSelectedIndexChanged(EventArgs e) { base.OnSelectedIndexChanged(e); Invalidate(); }
    }
    /// <summary>Bouton WinForms qui adapte le texte désactivé au thème sombre.</summary>
    public sealed class ThemedButton : UiActionButton { }
    /// <summary>ComboBox WinForms dont la flèche et le contour sont repeints en thème sombre.</summary>
    public sealed class ThemedComboBox : UiComboBox { }
}

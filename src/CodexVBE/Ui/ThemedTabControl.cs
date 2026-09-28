using System.Drawing;
using System.Windows.Forms;
namespace CodexVBE
{
    /// <summary>Dessine les onglets WinForms avec les couleurs et indicateurs de focus du thème actif.</summary>
    public sealed class ThemedTabControl : TabControl
    {
        /// <summary>Affiche une croix de fermeture sur chaque onglet.</summary>
        [System.ComponentModel.DefaultValue(false)]
        public bool ShowCloseButtons { get; set; }
        /// <summary>Demande la fermeture de l'onglet désigné.</summary>
        public event System.EventHandler<TabControlEventArgs> CloseRequested;
        private Rectangle CloseBounds(int index)
        { var r = GetTabRect(index); return new Rectangle(r.Right - 21, r.Top + (r.Height - 16) / 2, 16, 16); }
        /// <summary>Route le clic de fermeture sans changer les autres onglets.</summary>
        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (ShowCloseButtons && e.Button == MouseButtons.Left)
                for (int i = 0; i < TabCount; i++)
                    if (CloseBounds(i).Contains(e.Location)) { CloseRequested?.Invoke(this, new TabControlEventArgs(TabPages[i], i, TabControlAction.Deselecting)); return; }
            base.OnMouseDown(e);
        }
        /// <summary>Active le dessin personnalisé et le double buffering des onglets.</summary>
        public ThemedTabControl() { SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true); }
        /// <summary>Dessine le fond, les onglets, leurs états désactivés, la sélection et le focus.</summary>
        /// <param name="e">Contexte graphique et zone de peinture fournis par WinForms.</param>
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(UiTheme.Background);
            for (int i = 0; i < TabPages.Count; i++)
            {
                var bounds = GetTabRect(i); bool selected = i == SelectedIndex;
                using (var brush = new SolidBrush(selected ? UiTheme.Surface : UiTheme.Background)) e.Graphics.FillRectangle(brush, bounds);
                Color ink = TabPages[i].Enabled ? ForeColor : (UiTheme.Dark ? Color.FromArgb(148, 163, 184) : SystemColors.GrayText);
                var textBounds = bounds;
                if (ShowCloseButtons)
                {
                    textBounds.Width -= 24;
                    var close = CloseBounds(i);
                    using (var pen = new Pen(ink, 1.5F))
                    { e.Graphics.DrawLine(pen, close.Left + 4, close.Top + 4, close.Right - 4, close.Bottom - 4); e.Graphics.DrawLine(pen, close.Left + 4, close.Bottom - 4, close.Right - 4, close.Top + 4); }
                }
                TextRenderer.DrawText(e.Graphics, TabPages[i].Text, Font, textBounds, ink, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
                if (selected) using (var pen = new Pen(UiTheme.Dark ? Color.FromArgb(96, 165, 250) : Color.RoyalBlue, 2)) e.Graphics.DrawLine(pen, bounds.Left + 3, bounds.Bottom - 2, bounds.Right - 3, bounds.Bottom - 2);
                if (selected && Focused) ControlPaint.DrawFocusRectangle(e.Graphics, Rectangle.Inflate(bounds, -3, -3), ink, UiTheme.Surface);
            }
        }
        /// <summary>Transmet le changement de sélection puis invalide le contrôle pour redessiner l’onglet actif.</summary>
        /// <param name="e">Données de l’événement WinForms.</param>
        protected override void OnSelectedIndexChanged(System.EventArgs e) { base.OnSelectedIndexChanged(e); Invalidate(); }
    }
    /// <summary>Bouton WinForms qui adapte le texte désactivé au thème sombre.</summary>
    public sealed class ThemedButton : Button
    {
        /// <summary>Dessine le bouton natif puis renforce le contraste de son texte désactivé en thème sombre.</summary>
        /// <param name="e">Contexte graphique et zone de peinture fournis par WinForms.</param>
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
    /// <summary>ComboBox WinForms dont la flèche et le contour sont repeints en thème sombre.</summary>
    public sealed class ThemedComboBox : ComboBox
    {
        /// <summary>Traite le message Win32 puis repeint le bouton de liste en thème sombre hors mode contraste élevé.</summary>
        /// <param name="message">Message Win32 reçu par le contrôle.</param>
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

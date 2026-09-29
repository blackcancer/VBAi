using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Fonctions de dessin partagées par les contrôles visuels de conversation.</summary>
    internal static class ChatControlPainting
    {
        /// <summary>Construit un chemin graphique rectangulaire aux coins arrondis.</summary>
        /// <param name="bounds">Rectangle à arrondir.</param>
        /// <param name="radius">Rayon nominal des coins, limité aux dimensions disponibles.</param>
        /// <returns>Chemin fermé représentant le rectangle arrondi.</returns>
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
    /// <summary>Bouton WinForms peint avec un fond arrondi et des états de survol/focus.</summary>
    [ToolboxItem(true)]
    public class ChatActionButton : UiActionButton { }

    /// <summary>ComboBox propriétaire dessinée pour les choix de l’interface conversation.</summary>
    [ToolboxItem(true)]
    public class ChatChoiceBox : UiComboBox
    {
        /// <summary>Creates a choice-only selector with the shared input appearance.</summary>
        public ChatChoiceBox() { DropDownStyle = ComboBoxStyle.DropDownList; }
    }

    /// <summary>Panneau à fond blanc et bordure arrondie pour la zone de composition.</summary>
    [ToolboxItem(true)]
    public class ChatComposerPanel : TableLayoutPanel
    {
        /// <summary>Whether a boundary is drawn around this surface.</summary>
        /// <value>The current value represented by this member.</value>
        [Category("Appearance"), DefaultValue(true)]
        public bool ShowBorder { get; set; } = true;
        /// <summary>Crée le panneau de composition avec double tampon et fond blanc.</summary>
        public ChatComposerPanel() { DoubleBuffered = true; BackColor = Color.White; }
        /// <summary>Dessine le fond arrondi sans effacer la surface par défaut.</summary>
        /// <param name="e">Données de l’événement graphique ou pointeur.</param>
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            if (Width < 2 || Height < 2) return;
            using (var background = new SolidBrush(Parent?.BackColor ?? SystemColors.Control)) e.Graphics.FillRectangle(background, ClientRectangle);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = ChatControlPainting.Rounded(new Rectangle(0, 0, Width - 1, Height - 1), 12 * DeviceDpi / 96))
            using (var fill = new SolidBrush(BackColor))
            using (var pen = new Pen(UiTheme.BorderFor(this)))
            { e.Graphics.FillPath(fill, path); if (ShowBorder) e.Graphics.DrawPath(pen, path); }
        }
    }
}

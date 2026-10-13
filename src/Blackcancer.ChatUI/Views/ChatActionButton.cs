// SPDX-License-Identifier: MPL-2.0
// Native accessible button renderer adapted from SolidAi/VBAi; symbol glyphs are host-independent strings.
using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>A compact accessible command with a Designer-editable symbol and caption.</summary>
    [ToolboxItem(true)]
    public class ChatActionButton : Button
    {
        /// <summary>Stores the symbol used by ChatActionButton.</summary>
        private string symbol = string.Empty;
        /// <summary>Stores the icon only,primary,hovered,pressed used by ChatActionButton.</summary>
        private bool iconOnly, primary, hovered, pressed;
        /// <summary>Stores the caption tip used by ChatActionButton.</summary>
        private readonly ToolTip captionTip = new ToolTip { ShowAlways = true };
        /// <summary>Stores the symbol font used by ChatActionButton.</summary>
        private static readonly string SymbolFont = FindSymbolFont();
        /// <summary>Creates a native keyboard-accessible command.</summary>
        public ChatActionButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            FlatStyle = FlatStyle.Flat; FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
        }
        /// <summary>Symbol displayed beside the caption or on its own.</summary>
        /// <value>The current value represented by this member.</value>
        [Category("Appearance"), DefaultValue("")]
        public string Symbol { get => symbol; set { symbol = value ?? string.Empty; Invalidate(); } }
        /// <summary>Hides the painted caption while retaining its tooltip and accessible name.</summary>
        /// <value>The current value represented by this member.</value>
        [Category("Appearance"), DefaultValue(false)]
        public bool IconOnly { get => iconOnly; set { iconOnly = value; UpdateCaption(); Invalidate(); } }
        /// <summary>Emphasizes the main action in its group.</summary>
        /// <value>The current value represented by this member.</value>
        [Category("Appearance"), DefaultValue(false)]
        public bool Primary { get => primary; set { primary = value; Invalidate(); } }
        /// <summary>Chooses the available Windows icon font for symbols without a bundled SVG.</summary>
        /// <returns>The result produced by this operation.</returns>
        private static string FindSymbolFont()
        {
            using (var font = new Font("Segoe Fluent Icons", 12))
                return ChooseSymbolFont(font.Name);
        }
        /// <summary>Selects the legacy symbol font when Windows substitutes the requested Fluent font.</summary>
        /// <param name="resolvedName">Font name resolved by the native font subsystem.</param>
        /// <returns>The installed Fluent font or the legacy Windows symbol font.</returns>
        private static string ChooseSymbolFont(string resolvedName) => resolvedName == "Segoe Fluent Icons" ? resolvedName : "Segoe MDL2 Assets";
        /// <summary>Updates the tooltip when the caption is hidden.</summary>
        private void UpdateCaption()
        {
            if (captionTip != null) captionTip.SetToolTip(this, iconOnly ? Text : "");
        }
        /// <inheritdoc/>
        /// <summary>Refreshes the tooltip and invalidates the caption after its text changes.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnTextChanged(EventArgs e) { base.OnTextChanged(e); UpdateCaption(); Invalidate(); }
        /// <inheritdoc/>
        /// <summary>Measures the caption and icon without accumulating width across layout passes.</summary>
        /// <param name="proposedSize">The proposed size used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        public override Size GetPreferredSize(Size proposedSize)
        {
            if (IconOnly && Symbol != string.Empty) return new Size(32 * DeviceDpi / 96, 30 * DeviceDpi / 96);
            // Measuring the caption directly avoids feeding the current auto-sized width
            // back into the next layout pass and adding the symbol width repeatedly.
            var caption = TextRenderer.MeasureText(Text, Font);
            int extra = (Symbol == string.Empty ? 20 : 44) * DeviceDpi / 96;
            return new Size(Math.Max(MinimumSize.Width, caption.Width + Padding.Horizontal + extra),
                Math.Max(MinimumSize.Height, Math.Max(caption.Height + Padding.Vertical + 10 * DeviceDpi / 96, 30 * DeviceDpi / 96)));
        }
        /// <inheritdoc/>
        /// <summary>Records the hover state and requests a repaint.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnMouseEnter(EventArgs e) { hovered = true; Invalidate(); base.OnMouseEnter(e); }
        /// <inheritdoc/>
        /// <summary>Clears hover state and requests a repaint.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnMouseLeave(EventArgs e) { hovered = pressed = false; Invalidate(); base.OnMouseLeave(e); }
        /// <inheritdoc/>
        /// <summary>Handles native mouse presses and updates the command state.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnMouseDown(MouseEventArgs e) { pressed = e.Button == MouseButtons.Left; Invalidate(); base.OnMouseDown(e); }
        /// <inheritdoc/>
        /// <summary>Releases the pressed state and requests a repaint.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnMouseUp(MouseEventArgs e) { pressed = false; Invalidate(); base.OnMouseUp(e); }
        /// <inheritdoc/>
        /// <summary>Draws the control using its current palette, selection and focus state.</summary>
        /// <param name="e">Native event data.</param>
        protected override void OnPaint(PaintEventArgs e)
        {
            if (Width < 2 || Height < 2) return;
            bool contrast = SystemInformation.HighContrast;
            Color background = Parent?.BackColor ?? BackColor;
            Color foreground = Enabled ? ForeColor : SystemColors.GrayText;
            Color fill = background;
            if (Primary && Enabled)
            {
                // Resolve from the actual surface so Designer previews and runtime agree.
                bool darkSurface = background.GetBrightness() < .5f;
                fill = contrast ? SystemColors.Highlight
                    : darkSurface ? (pressed ? Color.FromArgb(192, 197, 205) : hovered ? Color.FromArgb(248, 249, 251) : Color.FromArgb(232, 235, 239))
                    : pressed ? Color.FromArgb(207, 213, 222) : hovered ? Color.FromArgb(220, 225, 232) : Color.FromArgb(234, 237, 242);
                foreground = contrast ? SystemColors.HighlightText : darkSurface ? Color.FromArgb(25, 28, 33) : Color.FromArgb(47, 55, 67);
            }
            else if (Enabled && (hovered || pressed))
            {
                fill = contrast ? SystemColors.Highlight : background.GetBrightness() < .5f ? Color.FromArgb(53, 58, 68) : Color.FromArgb(229, 233, 240);
                if (contrast) foreground = SystemColors.HighlightText;
            }
            using (var brush = new SolidBrush(background)) e.Graphics.FillRectangle(brush, ClientRectangle);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            int r = Math.Min(10 * DeviceDpi / 96, Math.Min(Width - 1, Height - 1));
            using (var path = new GraphicsPath())
            {
                path.AddArc(0, 0, r, r, 180, 90); path.AddArc(Width - r - 1, 0, r, r, 270, 90);
                path.AddArc(Width - r - 1, Height - r - 1, r, r, 0, 90); path.AddArc(0, Height - r - 1, r, r, 90, 90); path.CloseFigure();
                using (var brush = new SolidBrush(fill)) e.Graphics.FillPath(brush, path);
                if (Primary && Enabled && !contrast && background.GetBrightness() >= .5f)
                    using (var pen = new Pen(Color.FromArgb(198, 205, 216))) e.Graphics.DrawPath(pen, path);
                if (Focused && ShowFocusCues) using (var pen = new Pen(contrast ? SystemColors.WindowText : Color.FromArgb(90, 147, 241), 2)) e.Graphics.DrawPath(pen, path);
            }
            var textBounds = ClientRectangle;
            if (Symbol != string.Empty)
            {
                int width = IconOnly ? Width : 30 * DeviceDpi / 96;
                int groupWidth = IconOnly ? Width : Math.Min(Width, TextRenderer.MeasureText(Text, Font).Width + width);
                int start = (Width - groupWidth) / 2;
                var iconBounds = new Rectangle(RightToLeft == RightToLeft.Yes && !IconOnly ? start + groupWidth - width : start, 0, width, Height);
                    using (var font = new Font(Symbol[0] >= '\uE000' ? SymbolFont : Font.FontFamily.Name, 12F))
                        TextRenderer.DrawText(e.Graphics, Symbol, font, iconBounds, foreground, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
                if (IconOnly) return;
                textBounds = new Rectangle(start + (RightToLeft == RightToLeft.Yes ? 0 : width), 0, Math.Max(0, groupWidth - width), Height);
            }
            TextRenderer.DrawText(e.Graphics, Text, Font, textBounds, foreground, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }
        /// <inheritdoc/>
        /// <summary>Releases the resources owned by this control before base disposal.</summary>
        /// <param name="disposing">Whether managed resources must also be released.</param>
        protected override void Dispose(bool disposing) { if (disposing) captionTip.Dispose(); base.Dispose(disposing); }
    }
}

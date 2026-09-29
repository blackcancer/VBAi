using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Shared Windows command symbols, serialized by the WinForms Designer.</summary>
    public enum UiSymbol
    {
        /// <summary>No symbol.</summary>
        None = 0,
        /// <summary>Add.</summary>
        Add = 0xE710,
        /// <summary>More commands.</summary>
        More = 0xE712,
        /// <summary>Close or cancel.</summary>
        Close = 0xE711,
        /// <summary>Settings.</summary>
        Settings = 0xE713,
        /// <summary>History.</summary>
        History = 0xE81C,
        /// <summary>Attach context.</summary>
        Attach = 0xE723,
        /// <summary>Code.</summary>
        Code = 0xE943,
        /// <summary>Refresh.</summary>
        Refresh = 0xE72C,
        /// <summary>Save.</summary>
        Save = 0xE74E,
        /// <summary>Delete.</summary>
        Delete = 0xE74D,
        /// <summary>Edit.</summary>
        Edit = 0xE70F,
        /// <summary>Copy.</summary>
        Copy = 0xE8C8,
        /// <summary>Pin.</summary>
        Pin = 0xE718,
        /// <summary>Upload or send.</summary>
        Upload = 0xE898,
        /// <summary>Download.</summary>
        Download = 0xE896,
        /// <summary>Stop.</summary>
        Stop = 0xE71A,
        /// <summary>Play or resume.</summary>
        Play = 0xE768,
        /// <summary>Undo or restore.</summary>
        Undo = 0xE7A7,
        /// <summary>Inspect.</summary>
        Inspect = 0xE890,
        /// <summary>Search.</summary>
        Search = 0xE721,
        /// <summary>Accept or verify.</summary>
        Check = 0xE73E,
        /// <summary>Open folder.</summary>
        Folder = 0xE8B7,
        /// <summary>Email.</summary>
        Mail = 0xE715,
        /// <summary>Previous.</summary>
        Previous = 0xE76B,
        /// <summary>Next.</summary>
        Next = 0xE76C
    }

    /// <summary>A compact accessible command with a Designer-editable symbol and caption.</summary>
    [ToolboxItem(true)]
    public class UiActionButton : Button
    {
        /// <summary>Stores the symbol used by UiActionButton.</summary>
        private UiSymbol symbol;
        /// <summary>Stores the icon only,primary,hovered,pressed used by UiActionButton.</summary>
        private bool iconOnly, primary, hovered, pressed;
        /// <summary>Stores the caption tip used by UiActionButton.</summary>
        private readonly ToolTip captionTip = new ToolTip { ShowAlways = true };
        /// <summary>Stores the symbol font used by UiActionButton.</summary>
        private static readonly string SymbolFont = FindSymbolFont();
        /// <summary>Creates a native keyboard-accessible command.</summary>
        public UiActionButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            FlatStyle = FlatStyle.Flat; FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
        }
        /// <summary>Symbol displayed beside the caption or on its own.</summary>
        /// <value>The current value represented by this member.</value>
        [Category("Appearance"), DefaultValue(UiSymbol.None)]
        public UiSymbol Symbol { get => symbol; set { symbol = value; Invalidate(); } }
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
            if (IconOnly && Symbol != UiSymbol.None) return new Size(32 * DeviceDpi / 96, 30 * DeviceDpi / 96);
            // Measuring the caption directly avoids feeding the current auto-sized width
            // back into the next layout pass and adding the symbol width repeatedly.
            var caption = TextRenderer.MeasureText(Text, Font);
            int extra = (Symbol == UiSymbol.None ? 20 : 44) * DeviceDpi / 96;
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
            bool contrast = UiTheme.HighContrast();
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
            if (Symbol != UiSymbol.None)
            {
                int width = IconOnly ? Width : 30 * DeviceDpi / 96;
                int groupWidth = IconOnly ? Width : Math.Min(Width, TextRenderer.MeasureText(Text, Font).Width + width);
                int start = (Width - groupWidth) / 2;
                var iconBounds = new Rectangle(RightToLeft == RightToLeft.Yes && !IconOnly ? start + groupWidth - width : start, 0, width, Height);
                if (!UiCommandIcons.Draw(e.Graphics, Symbol, iconBounds, foreground, DeviceDpi))
                    using (var font = new Font(SymbolFont, 12F)) TextRenderer.DrawText(e.Graphics, char.ConvertFromUtf32((int)Symbol), font, iconBounds, foreground, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
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

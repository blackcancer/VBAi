using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
namespace CodexVBE
{
    internal sealed class ChatDesignerHost : WindowsFormsHost
    {
        internal UserControl View => Child as UserControl;
        internal ChatDesignerHost(UserControl view)
        {
            Child = view;
            view.SizeChanged += (s,e) => InvalidateMeasure();
        }
        protected override Size MeasureOverride(Size constraint)
        {
            if (View == null || View.IsDisposed) return new Size();
            int width = (int)Math.Max(40, double.IsInfinity(constraint.Width) ? 500 : constraint.Width);
            if (View.Width != width || View.MaximumSize.Width != width) {
                View.MinimumSize = System.Drawing.Size.Empty;
                View.MaximumSize = new System.Drawing.Size(width,0);
                View.MinimumSize = new System.Drawing.Size(width,0);
                View.Width = width;
            }
            View.PerformLayout();
            var size = View.GetPreferredSize(new System.Drawing.Size(width, 0));
            return new Size(width, Math.Max(24, Math.Min(32000, size.Height)));
        }
    }
}

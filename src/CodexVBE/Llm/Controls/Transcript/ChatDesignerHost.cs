using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
namespace CodexVBE
{
    /// <summary>Hosts a transcript view in the WinForms Designer and measures its preferred height.</summary>
    internal sealed class ChatDesignerHost : WindowsFormsHost
    {
        /// <summary>Gets the view.</summary>
        /// <value>The current value represented by this member.</value>
        internal UserControl View => Child as UserControl;
        /// <summary>Initializes a ChatDesignerHost instance with the supplied state.</summary>
        /// <param name="view">The view used by this operation.</param>
        internal ChatDesignerHost(UserControl view)
        {
            Child = view;
            view.SizeChanged += (s,e) => InvalidateMeasure();
        }
        /// <summary>Measures the hosted transcript view within the available designer width.</summary>
        /// <param name="constraint">The constraint used by this operation.</param>
        /// <returns>The measured size required by the hosted transcript view.</returns>
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

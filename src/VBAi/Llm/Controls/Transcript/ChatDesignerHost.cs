using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
namespace VBAi
{

    /// <summary>Hosts a transcript view in the WinForms Designer and measures its preferred height.</summary>
    internal sealed class ChatDesignerHost : WindowsFormsHost
    {

        /// <summary>Gets the hosted transcript control when the native child is a WinForms UserControl.</summary>
        /// <value>The child UserControl, or <see langword="null"/> before assignment or for another child type.</value>
        internal UserControl View => Child as UserControl;

        /// <summary>Hosts the transcript view and invalidates WPF measurement when its native size changes.</summary>
        /// <param name="view">WinForms transcript control to host; it must be non-null.</param>
        internal ChatDesignerHost(UserControl view)
        {
            Child = view;
            view.SizeChanged += (s, e) => InvalidateMeasure();
        }

        /// <summary>Measures the hosted transcript view within the available designer width.</summary>
        /// <param name="constraint">Available WPF size in device-independent units; infinite width uses a 500-unit fallback.</param>
        /// <returns>The measured size required by the hosted transcript view.</returns>
        protected override Size MeasureOverride(Size constraint)
        {
            if (View == null || View.IsDisposed) return new Size();
            double width = Math.Max(40, double.IsInfinity(constraint.Width) ? 500 : constraint.Width);
            // WindowsFormsHost owns DIP/pixel conversion, DPI scaling and child arrangement.
            // A width change is remeasured through SizeChanged after native arrangement.
            var size = base.MeasureOverride(new Size(width, double.PositiveInfinity));
            return new Size(width, Math.Max(24, Math.Min(32000, size.Height)));
        }
    }
}

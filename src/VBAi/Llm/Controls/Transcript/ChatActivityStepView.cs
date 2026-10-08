using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Displays one chat activity step with its state and expandable details.</summary>
    public sealed partial class ChatActivityStepView : ChatDesignerView
    {

        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatActivityStepView()
        {
            InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);
            WatchMeasurements(this);
        }

        /// <summary>Retains bounded width-specific measurements until this row's actual layout inputs change.</summary>
        private readonly Dictionary<Size, Size> preferredSizes = new Dictionary<Size, Size>();

        /// <summary>Tracks the fixed Designer subtree and any controls subsequently added to it.</summary>
        private readonly HashSet<Control> measurementControls = new HashSet<Control>();

        /// <summary>Reuses stable row measurements while a sibling streams, without freezing this row's content or dimensions.</summary>
        /// <param name="proposedSize">The parent's width and height constraint in native pixels.</param>
        /// <returns>The same constrained preferred size supplied by WinForms for the current row inputs.</returns>
        public override Size GetPreferredSize(Size proposedSize)
        {
            if (preferredSizes == null || IsDisposed || Disposing) return base.GetPreferredSize(proposedSize);
            if (preferredSizes.TryGetValue(proposedSize, out var size)) return size;
            size = base.GetPreferredSize(proposedSize);
            if (preferredSizes.Count >= 8) preferredSizes.Clear();
            preferredSizes[proposedSize] = size;
            return size;
        }

        /// <summary>Invalidates measurements before a local layout consumes changed text, font, constraints or child state.</summary>
        /// <param name="e">The control and property that caused the layout.</param>
        protected override void OnLayout(LayoutEventArgs e)
        {
            preferredSizes?.Clear();
            base.OnLayout(e);
        }

        /// <summary>Observes descendant inputs even when their owning layouts are temporarily suspended.</summary>
        /// <param name="control">The subtree whose changes affect this row's preferred height.</param>
        private void WatchMeasurements(Control control)
        {
            if (!measurementControls.Add(control)) return;
            control.Layout += InvalidateMeasurement;
            control.SizeChanged += InvalidateMeasurement;
            control.TextChanged += InvalidateMeasurement;
            control.FontChanged += InvalidateMeasurement;
            control.VisibleChanged += InvalidateMeasurement;
            control.PaddingChanged += InvalidateMeasurement;
            control.MarginChanged += InvalidateMeasurement;
            control.AutoSizeChanged += InvalidateMeasurement;
            control.DockChanged += InvalidateMeasurement;
            control.ControlAdded += MeasurementControlAdded;
            control.ControlRemoved += MeasurementControlRemoved;
            foreach (Control child in control.Controls) WatchMeasurements(child);
        }

        /// <summary>Detaches removed children so later changes cannot retain or invalidate this row.</summary>
        /// <param name="control">Subtree being detached from the row.</param>
        private void UnwatchMeasurements(Control control)
        {
            if (!measurementControls.Remove(control)) return;
            control.Layout -= InvalidateMeasurement;
            control.SizeChanged -= InvalidateMeasurement;
            control.TextChanged -= InvalidateMeasurement;
            control.FontChanged -= InvalidateMeasurement;
            control.VisibleChanged -= InvalidateMeasurement;
            control.PaddingChanged -= InvalidateMeasurement;
            control.MarginChanged -= InvalidateMeasurement;
            control.AutoSizeChanged -= InvalidateMeasurement;
            control.DockChanged -= InvalidateMeasurement;
            control.ControlAdded -= MeasurementControlAdded;
            control.ControlRemoved -= MeasurementControlRemoved;
            foreach (Control child in control.Controls) UnwatchMeasurements(child);
        }

        /// <summary>Drops obsolete measurements whenever a descendant changes a layout input.</summary>
        private void InvalidateMeasurement(object sender, EventArgs e) => preferredSizes.Clear();

        /// <summary>Admits dynamically inserted descendants into the measurement invalidation tree.</summary>
        private void MeasurementControlAdded(object sender, ControlEventArgs e)
        {
            preferredSizes.Clear(); WatchMeasurements(e.Control);
        }

        /// <summary>Releases removed descendants and invalidates the former child arrangement.</summary>
        private void MeasurementControlRemoved(object sender, ControlEventArgs e)
        {
            preferredSizes.Clear(); UnwatchMeasurements(e.Control);
        }
    }
}

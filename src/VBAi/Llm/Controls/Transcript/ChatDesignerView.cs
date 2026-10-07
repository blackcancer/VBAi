using System;
using System.Collections.Generic;
using System.Windows.Forms;
namespace VBAi
{

    /// <summary>WinForms base that keeps variable vertical rows aligned to their Designer container.</summary>
    public class ChatDesignerView : UserControl
    {

        /// <summary>Prevents duplicate resize subscriptions on nested top-down flow panels.</summary>
        private readonly HashSet<FlowLayoutPanel> watched = new HashSet<FlowLayoutPanel>();

        /// <summary>Connects resize behavior to Designer-created containers.</summary>
        /// <param name="e">Added control.</param>
        protected override void OnControlAdded(ControlEventArgs e)
        {
            base.OnControlAdded(e); Watch(e.Control);
        }

        /// <summary>Registers a control whose preferred height should trigger row remeasurement.</summary>
        /// <param name="parent">New control subtree to inspect for top-down FlowLayoutPanel containers.</param>
        private void Watch(Control parent)
        {
            if (parent is FlowLayoutPanel flow && flow.FlowDirection == FlowDirection.TopDown && watched.Add(flow))
            {
                int lastWidth = -1;
                flow.SizeChanged += (s, e) =>
                {
                    if (lastWidth == flow.ClientSize.Width) return;
                    lastWidth = flow.ClientSize.Width; ResizeRows(flow);
                };
                flow.ControlAdded += (s, e) => ResizeRows(flow);
            }
            foreach (Control child in parent.Controls) Watch(child);
        }

        /// <summary>Resizes the designer host rows to fit their current transcript controls.</summary>
        /// <param name="flow">Top-down panel whose child controls must fit its current client width.</param>
        private static void ResizeRows(FlowLayoutPanel flow)
        {
            if (flow.ClientSize.Width <= 0) return;
            foreach (Control child in flow.Controls)
            {
                int width = Math.Max(20, flow.ClientSize.Width - flow.Padding.Horizontal - child.Margin.Horizontal);
                if (child.Width == width) continue;
                int minimumHeight = child.MinimumSize.Height;
                child.MinimumSize = System.Drawing.Size.Empty;
                child.MaximumSize = new System.Drawing.Size(width, 0);
                child.MinimumSize = new System.Drawing.Size(width, minimumHeight);
                child.Width = width;
            }
        }
    }
}

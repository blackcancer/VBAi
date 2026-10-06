using System;
using System.Collections.Generic;
using System.Windows.Forms;
namespace VBAi
{

    /// <summary>WinForms base that keeps variable vertical rows aligned to their Designer container.</summary>
    public class ChatDesignerView : UserControl
    {

        /// <summary>Maintains the watched state for chat designer view.</summary>
        private readonly HashSet<FlowLayoutPanel> watched = new HashSet<FlowLayoutPanel>();

        /// <summary>Connects resize behavior to Designer-created containers.</summary>
        /// <param name="e">Added control.</param>
        protected override void OnControlAdded(ControlEventArgs e)
        {
            base.OnControlAdded(e); Watch(e.Control);
        }

        /// <summary>Registers a control whose preferred height should trigger row remeasurement.</summary>
        /// <param name="parent">control that supplies the parent for this operation.</param>
        private void Watch(Control parent)
        {
            if (parent is FlowLayoutPanel flow && flow.FlowDirection == FlowDirection.TopDown && watched.Add(flow)) {
                int lastWidth = -1;
                flow.SizeChanged += (s,e) => {
                    if (lastWidth == flow.ClientSize.Width) return;
                    lastWidth = flow.ClientSize.Width; ResizeRows(flow);
                };
                flow.ControlAdded += (s,e) => ResizeRows(flow);
            }
            foreach (Control child in parent.Controls) Watch(child);
        }

        /// <summary>Resizes the designer host rows to fit their current transcript controls.</summary>
        /// <param name="flow">flow layout panel that supplies the flow for this operation.</param>
        private static void ResizeRows(FlowLayoutPanel flow)
        {
            if (flow.ClientSize.Width <= 0) return;
            foreach (Control child in flow.Controls) {
                int width = Math.Max(20, flow.ClientSize.Width - flow.Padding.Horizontal - child.Margin.Horizontal);
                if (child.Width == width) continue;
                int minimumHeight = child.MinimumSize.Height;
                child.MinimumSize = System.Drawing.Size.Empty;
                child.MaximumSize = new System.Drawing.Size(width,0);
                child.MinimumSize = new System.Drawing.Size(width, minimumHeight);
                child.Width = width;
            }
        }
    }
}

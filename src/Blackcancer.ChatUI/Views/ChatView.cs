// SPDX-License-Identifier: MPL-2.0
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Vertical transcript rows retain the width assigned by their container.</summary>
    public class ChatView : UserControl
    {
        private readonly HashSet<FlowLayoutPanel> watched = new HashSet<FlowLayoutPanel>();
        public ChatView() { DoubleBuffered = true; }
        public override Size GetPreferredSize(Size proposedSize)
        {
            Size preferred = base.GetPreferredSize(proposedSize);
            int width = proposedSize.Width > 0 && proposedSize.Width < int.MaxValue ? proposedSize.Width : Math.Max(1, Width);
            return new Size(width, preferred.Height);
        }
        protected override void OnControlAdded(ControlEventArgs e) { base.OnControlAdded(e); Watch(e.Control); }
        private void Watch(Control control)
        {
            var flow = control as FlowLayoutPanel;
            if (flow != null && flow.FlowDirection == FlowDirection.TopDown && watched.Add(flow))
            {
                int lastWidth = -1;
                flow.SizeChanged += (s, e) => { if (lastWidth == flow.ClientSize.Width) return; lastWidth = flow.ClientSize.Width; ResizeRows(flow); };
                flow.ControlAdded += (s, e) => { Watch(e.Control); ResizeRows(flow); };
            }
            foreach (Control child in control.Controls) Watch(child);
        }
        private static void ResizeRows(FlowLayoutPanel flow)
        {
            if (flow.ClientSize.Width <= 0) return;
            foreach (Control child in flow.Controls)
            {
                int width = Math.Max(1, flow.ClientSize.Width - flow.Padding.Horizontal - child.Margin.Horizontal);
                if (child.Width == width && child.MinimumSize.Width == width && child.MaximumSize.Width == width) continue;
                int minimumHeight = child.MinimumSize.Height;
                child.MinimumSize = Size.Empty; child.MaximumSize = new Size(width, 0); child.MinimumSize = new Size(width, minimumHeight); child.Width = width;
            }
        }
    }
}

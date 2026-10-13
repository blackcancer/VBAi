// SPDX-License-Identifier: MPL-2.0
using System.ComponentModel;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Buffered transcript; selecting text scrolls its editor without jumping the conversation.</summary>
    [ToolboxItem(true)]
    public sealed class ChatTranscriptView : FlowLayoutPanel
    {
        private int previousViewportHeight;
        private bool resizing;
        public ChatTranscriptView() { DoubleBuffered = true; }
        protected override Point ScrollToControl(Control activeControl) => AutoScrollPosition;
        protected override void OnSizeChanged(EventArgs e)
        {
            bool follow = previousViewportHeight > 0
                && -AutoScrollPosition.Y + previousViewportHeight >= DisplayRectangle.Height - 24;
            base.OnSizeChanged(e);
            previousViewportHeight = ClientSize.Height;
            if (!follow || resizing) return;
            resizing = true;
            try
            {
                PerformLayout();
                AutoScrollPosition = new Point(0, Math.Max(0, DisplayRectangle.Height - ClientSize.Height));
            }
            finally { resizing = false; }
        }
    }
}

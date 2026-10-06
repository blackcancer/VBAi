using System;
using System.Windows.Forms;
namespace VBAi
{

    /// <summary>Displays starter prompts when the chat has no messages.</summary>
    public sealed partial class ChatWelcomeView : ChatDesignerView
    {

        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatWelcomeView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }

        /// <summary>Measures wrapped labels at the native width assigned by the transcript host.</summary>
        /// <param name="proposedSize">Available native pixel dimensions after WindowsFormsHost DPI conversion.</param>
        /// <returns>The complete Designer table height, including all starter actions.</returns>
        public override System.Drawing.Size GetPreferredSize(System.Drawing.Size proposedSize)
        {
            if (layout == null) return base.GetPreferredSize(proposedSize);
            int width = proposedSize.Width > 0 ? proposedSize.Width : Width;
            var content = layout.GetPreferredSize(new System.Drawing.Size(Math.Max(1, width - Padding.Horizontal), 0));
            return new System.Drawing.Size(content.Width + Padding.Horizontal, content.Height + Padding.Vertical);
        }
    }
}

namespace VBAi
{

    /// <summary>Displays an actionable link in the chat transcript.</summary>
    public sealed partial class ChatLinkView : ChatDesignerView
    {

        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatLinkView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this); }

        /// <summary>Provides the token width to horizontal reference lists without a host-imposed width.</summary>
        /// <param name="proposedSize">Available space.</param>
        /// <returns>Size needed by the Designer button and its margins.</returns>
        public override System.Drawing.Size GetPreferredSize(System.Drawing.Size proposedSize)
        {
            if (link == null) return base.GetPreferredSize(proposedSize);
            var size = link.GetPreferredSize(System.Drawing.Size.Empty);
            return new System.Drawing.Size(size.Width + layout.Padding.Horizontal + link.Margin.Horizontal,
                size.Height + layout.Padding.Vertical + link.Margin.Vertical);
        }
    }
}

namespace VBAi
{

    /// <summary>Displays a transcript message with its content, metadata, and available actions.</summary>
    public sealed partial class ChatMessageView : ChatDesignerView
    {

        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatMessageView()
        {
            InitializeComponent();
            TrackMetadataVisibility(attachments);
            TrackMetadataVisibility(references);
            TrackMetadataVisibility(targets);
            UiText.Apply(this, components);
            UiTheme.Apply(this);
        }

        /// <summary>Excludes empty metadata rows from layout and updates their visibility as children change.</summary>
        /// <param name="panel">The attachment, reference or intervention-target row owned by this card.</param>
        private static void TrackMetadataVisibility(System.Windows.Forms.FlowLayoutPanel panel)
        {
            System.Windows.Forms.ControlEventHandler refresh = (sender, args) =>
            {
                if (!panel.Disposing && !panel.IsDisposed)
                    panel.Visible = panel.Controls.Count > 0;
            };
            panel.ControlAdded += refresh;
            panel.ControlRemoved += refresh;
            panel.Visible = panel.Controls.Count > 0;
        }

        /// <summary>Gets the action button that copies this message&apos;s content.</summary>
        /// <value>Button that copies this message&apos;s displayed content.</value>
        internal ChatActionButton CopyButton => copy;

        /// <summary>Gets the action button that starts a new conversation from this message.</summary>
        /// <value>Button that starts a new chat from this message.</value>
        internal ChatActionButton ForkButton => fork;
    }
}

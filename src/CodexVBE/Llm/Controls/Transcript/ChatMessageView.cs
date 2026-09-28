using System;
using System.Windows.Forms;
namespace CodexVBE
{
    /// <summary>Designer-editable ChatMessageView layout.</summary>
    public sealed partial class ChatMessageView : ChatDesignerView
    {
        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatMessageView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
        internal ChatActionButton CopyButton => copy;
        internal ChatActionButton ForkButton => fork;
    }
}

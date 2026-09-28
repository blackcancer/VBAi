using System;
using System.Windows.Forms;
namespace CodexVBE
{
        /// <summary>Displays a transcript message with its content, metadata, and available actions.</summary>
public sealed partial class ChatMessageView : ChatDesignerView
    {
        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatMessageView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
        /// <summary>Gets the action button that copies this message&apos;s content.</summary>
/// <value>Button that copies this message&apos;s displayed content.</value>
internal ChatActionButton CopyButton => copy;
        /// <summary>Gets the action button that starts a new conversation from this message.</summary>
/// <value>Button that starts a new chat from this message.</value>
internal ChatActionButton ForkButton => fork;
    }
}

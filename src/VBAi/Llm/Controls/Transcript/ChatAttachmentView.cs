using System;
using System.Windows.Forms;
namespace VBAi
{

    /// <summary>Displays an attachment included in a chat message.</summary>
    public sealed partial class ChatAttachmentView : ChatDesignerView
    {

        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatAttachmentView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
    }
}

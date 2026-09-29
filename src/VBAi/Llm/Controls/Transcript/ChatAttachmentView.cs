using System;
using System.Windows.Forms;
namespace VBAi
{
    /// <summary>Designer-editable ChatAttachmentView layout.</summary>
    public sealed partial class ChatAttachmentView : ChatDesignerView
    {
        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatAttachmentView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
    }
}

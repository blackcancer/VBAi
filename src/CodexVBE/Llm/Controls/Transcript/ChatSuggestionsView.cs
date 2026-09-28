using System;
using System.Windows.Forms;
namespace CodexVBE
{
    /// <summary>Designer-editable ChatSuggestionsView layout.</summary>
    public sealed partial class ChatSuggestionsView : ChatDesignerView
    {
        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatSuggestionsView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
    }
}

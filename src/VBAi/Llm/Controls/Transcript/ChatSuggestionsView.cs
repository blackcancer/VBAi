using System;
using System.Windows.Forms;
namespace VBAi
{

    /// <summary>Displays navigation and reference suggestions for the current chat draft.</summary>
    public sealed partial class ChatSuggestionsView : ChatDesignerView
    {

        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatSuggestionsView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
    }
}

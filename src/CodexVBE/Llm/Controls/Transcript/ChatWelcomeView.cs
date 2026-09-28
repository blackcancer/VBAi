using System;
using System.Windows.Forms;
namespace CodexVBE
{
        /// <summary>Displays starter prompts when the chat has no messages.</summary>
public sealed partial class ChatWelcomeView : ChatDesignerView
    {
        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatWelcomeView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);  }
    }
}

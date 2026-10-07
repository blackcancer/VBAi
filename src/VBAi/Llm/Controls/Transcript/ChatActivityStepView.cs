namespace VBAi
{

    /// <summary>Displays one chat activity step with its state and expandable details.</summary>
    public sealed partial class ChatActivityStepView : ChatDesignerView
    {

        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatActivityStepView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this); }
    }
}

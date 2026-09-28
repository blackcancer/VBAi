using System;
using System.Windows.Forms;
namespace CodexVBE
{
    /// <summary>Designer-editable ChatDisclosureView layout.</summary>
    [System.ComponentModel.Designer(typeof(ChatDisclosureDesigner))]
    public sealed partial class ChatDisclosureView : ChatDesignerView
    {
        /// <summary>Creates the fixed controls from the WinForms Designer.</summary>
        public ChatDisclosureView() { InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this); toggle.Click += (s,e) => Expanded = !Expanded; UpdateExpansion(); }
                /// <summary>Gets the flow panel used to add section specific controls in the Designer.</summary>
/// <value>Flow panel for controls serialized into this disclosure section by the Designer.</value>
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Content)]
        public System.Windows.Forms.FlowLayoutPanel ContentPanel => body;
        /// <summary>Stores the expanded used by ChatDisclosureView.</summary>
        private bool expanded;
        /// <summary>Stores the title used by ChatDisclosureView.</summary>
        private string title = "Details";
        /// <summary>Raised when the section is expanded or collapsed.</summary>
        public event EventHandler ExpansionChanged;
                /// <summary>Section caption, editable in the Designer.</summary>
/// <value>The current value represented by this member.</value>
        [System.ComponentModel.Category("Appearance")]
        public string Title { get => title; set { title = value; UpdateExpansion(); } }
                /// <summary>Whether the section body is expanded.</summary>
/// <value>The current value represented by this member.</value>
        [System.ComponentModel.DefaultValue(false), System.ComponentModel.Category("Appearance")]
        public bool Expanded { get => expanded; set { if (expanded == value) return; expanded = value; UpdateExpansion(); ExpansionChanged?.Invoke(this, EventArgs.Empty); } }
        /// <summary>Synchronizes the disclosure caption and body visibility with the expanded state.</summary>
        private void UpdateExpansion() { if (toggle == null) return; toggle.Text = (expanded ? "▾ " : "▸ ") + UiText.Get(title); body.Visible = expanded; }
    }
}

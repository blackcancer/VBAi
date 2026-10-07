namespace VBAi
{

    /// <summary>Displays one chat activity step with its state and expandable details.</summary>
    public sealed partial class ChatActivityStepView
    {

        /// <summary>Container that owns the disposable components created by the WinForms Designer.</summary>
        private System.ComponentModel.IContainer components;

        /// <summary>ToolTip component used to show full text for transcript controls.</summary>
        private System.Windows.Forms.ToolTip toolTips;

        /// <summary>Flow layout panel that contains this transcript view&apos;s child controls.</summary>
        private System.Windows.Forms.TableLayoutPanel layout;

        /// <summary>Expandable section containing the activity state and detailed tool result.</summary>
        internal ChatDisclosureView section;

        /// <summary>Displays the detailed text returned for this tool activity step.</summary>
        internal ChatTextContentView detail;

        /// <summary>Displays the current state of this tool activity step.</summary>
        internal System.Windows.Forms.Label state;

        /// <summary>Releases the Designer components.</summary>
        /// <param name="disposing">Whether to release managed resources.</param>
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }

        /// <summary>Creates and configures the chat activity step view controls serialized by the WinForms Designer.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.detail = new ChatTextContentView(); this.detail.Name = "detail"; this.detail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.section = new ChatDisclosureView();
            this.section.ContentPanel.Controls.Add(this.detail);
            this.section.Name = "section";
            this.section.Dock = System.Windows.Forms.DockStyle.Fill;
            this.section.Title = "Agent activity";
            this.state = new System.Windows.Forms.Label();
            this.state.Name = "state";
            this.state.AutoSize = true;
            this.state.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            this.state.UseCompatibleTextRendering = false;
            this.state.Text = "In progress";
            this.SuspendLayout(); this.layout.SuspendLayout();
            this.layout.AutoSize = true;
            this.layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.layout.Dock = System.Windows.Forms.DockStyle.Top;
            this.layout.ColumnCount = 2;
            this.layout.RowCount = 1;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            this.state.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.state.Margin = new System.Windows.Forms.Padding(3, 10, 3, 0);
            this.layout.Controls.Add(this.state, 1, 0);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.section, 0, 0);

            this.layout.Name = "layout";
            this.layout.Padding = new System.Windows.Forms.Padding(4, 0, 0, 0);
            this.Controls.Add(this.layout);
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "ChatActivityStepView";
            this.Size = new System.Drawing.Size(500, 120);
            this.layout.ResumeLayout(false); this.layout.PerformLayout(); this.ResumeLayout(false); this.PerformLayout();
        }
    }
}

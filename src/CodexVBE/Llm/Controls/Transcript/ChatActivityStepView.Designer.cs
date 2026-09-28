namespace CodexVBE
{
    public sealed partial class ChatActivityStepView
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.ToolTip toolTips;
        private System.Windows.Forms.TableLayoutPanel layout;
        internal ChatDisclosureView section;
        internal ChatTextContentView detail;
        internal System.Windows.Forms.Label state;
        /// <summary>Releases the Designer components.</summary>
        /// <param name="disposing">Whether to release managed resources.</param>
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }
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
            this.state.Dock = System.Windows.Forms.DockStyle.Fill;
            this.state.Text = "In progress";
            this.SuspendLayout(); this.layout.SuspendLayout();
            this.layout.AutoSize = true;
            this.layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.layout.Dock = System.Windows.Forms.DockStyle.Top;
            this.layout.ColumnCount = 1;
            this.layout.RowCount = 2;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.Controls.Add(this.state, 0, 0);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.section, 0, 1);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Name = "layout";
            this.layout.Padding = new System.Windows.Forms.Padding(10);
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

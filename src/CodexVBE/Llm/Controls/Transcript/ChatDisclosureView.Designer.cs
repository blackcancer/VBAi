namespace CodexVBE
{
    /// <summary>Displays transcript details in a section that can expand or collapse.</summary>
    public sealed partial class ChatDisclosureView
    {
        /// <summary>Container that owns the disposable components created by the WinForms Designer.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>ToolTip component used to show full text for transcript controls.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Flow layout panel that contains this transcript view&apos;s child controls.</summary>
        private System.Windows.Forms.TableLayoutPanel layout;
        /// <summary>Button that expands or collapses the section body.</summary>
        internal ChatActionButton toggle;
        /// <summary>Panel that contains the controls shown while the section is expanded.</summary>
        internal System.Windows.Forms.FlowLayoutPanel body;
        /// <summary>Releases the Designer components.</summary>
        /// <param name="disposing">Whether to release managed resources.</param>
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }
        /// <summary>Creates and configures the chat disclosure view controls serialized by the WinForms Designer.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.toggle = new ChatActionButton();
            this.toggle.Name = "toggle";
            this.toggle.AutoSize = true;
            this.toggle.Text = "Details";
            this.body = new System.Windows.Forms.FlowLayoutPanel();
            this.body.Name = "body";
            this.body.AutoSize = true;
            this.body.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.body.Dock = System.Windows.Forms.DockStyle.Fill;
            this.body.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.body.WrapContents = false;
            this.SuspendLayout(); this.layout.SuspendLayout();
            this.layout.AutoSize = true;
            this.layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.layout.Dock = System.Windows.Forms.DockStyle.Top;
            this.layout.ColumnCount = 1;
            this.layout.RowCount = 2;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.Controls.Add(this.toggle, 0, 0);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.body, 0, 1);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Name = "layout";
            this.layout.Padding = new System.Windows.Forms.Padding(10);
            this.toolTips.SetToolTip(this.toggle, "Show or hide details");
            this.Controls.Add(this.layout);
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "ChatDisclosureView";
            this.Size = new System.Drawing.Size(500, 80);
            this.layout.ResumeLayout(false); this.layout.PerformLayout(); this.ResumeLayout(false); this.PerformLayout();
        }
    }
}

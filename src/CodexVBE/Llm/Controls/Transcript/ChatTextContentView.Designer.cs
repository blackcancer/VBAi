namespace CodexVBE
{
    public sealed partial class ChatTextContentView
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.ToolTip toolTips;
        private System.Windows.Forms.TableLayoutPanel layout;
        internal System.Windows.Forms.RichTextBox content;
        internal System.Windows.Forms.ContextMenuStrip copyMenu;
        internal System.Windows.Forms.ToolStripMenuItem copySelection;
        internal System.Windows.Forms.ToolStripMenuItem copyCode;
        /// <summary>Releases the Designer components.</summary>
        /// <param name="disposing">Whether to release managed resources.</param>
        protected override void Dispose(bool disposing) { if (disposing) { DisposeTextResources(); components?.Dispose(); } base.Dispose(disposing); }
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.copyMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.copyMenu.Name = "copyMenu";
            this.copySelection = new System.Windows.Forms.ToolStripMenuItem("Copy");
            this.copyCode = new System.Windows.Forms.ToolStripMenuItem("Copy code");
            this.copyMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { this.copySelection, this.copyCode });
            this.content = new System.Windows.Forms.RichTextBox();
            this.content.ReadOnly = true;
            this.content.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.content.DetectUrls = false;
            this.content.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None;
            this.content.ContextMenuStrip = this.copyMenu;
            this.content.Name = "content";
            this.content.Dock = System.Windows.Forms.DockStyle.Top;
            this.content.Height = 40;
            this.SuspendLayout(); this.layout.SuspendLayout();
            this.layout.AutoSize = true;
            this.layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.layout.Dock = System.Windows.Forms.DockStyle.Top;
            this.layout.ColumnCount = 1;
            this.layout.RowCount = 1;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.Controls.Add(this.content, 0, 0);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Name = "layout";
            this.layout.Padding = new System.Windows.Forms.Padding(0);
            this.Controls.Add(this.layout);
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "ChatTextContentView";
            this.Size = new System.Drawing.Size(500, 44);
            this.layout.ResumeLayout(false); this.layout.PerformLayout(); this.ResumeLayout(false); this.PerformLayout();
        }
    }
}

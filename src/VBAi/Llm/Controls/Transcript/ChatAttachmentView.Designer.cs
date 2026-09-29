namespace VBAi
{
    /// <summary>Displays an attachment included in a chat message.</summary>
    public sealed partial class ChatAttachmentView
    {
        /// <summary>Container that owns the disposable components created by the WinForms Designer.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>ToolTip component used to show full text for transcript controls.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Flow layout panel that contains this transcript view&apos;s child controls.</summary>
        private System.Windows.Forms.TableLayoutPanel layout;
        /// <summary>Stores the section used by ChatAttachmentView.</summary>
        internal ChatDisclosureView section;
        /// <summary>Displays the attachment name and descriptive text.</summary>
        internal ChatTextContentView text;
        /// <summary>Opens the attached file when the user activates its action.</summary>
        internal ChatActionButton open;
        /// <summary>Releases the Designer components.</summary>
        /// <param name="disposing">Whether to release managed resources.</param>
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }
        /// <summary>Creates and configures the chat attachment view controls serialized by the WinForms Designer.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.text = new ChatTextContentView(); this.text.Name = "text"; this.text.Dock = System.Windows.Forms.DockStyle.Fill;
            this.open = new ChatActionButton(); this.open.Name = "open"; this.open.Text = "Open in the VBE"; this.open.AutoSize = true;
            this.section = new ChatDisclosureView();
            this.section.ContentPanel.Controls.Add(this.text); this.section.ContentPanel.Controls.Add(this.open);
            this.section.Name = "section";
            this.section.Dock = System.Windows.Forms.DockStyle.Fill;
            this.section.Title = "Context";
            this.SuspendLayout(); this.layout.SuspendLayout();
            this.layout.AutoSize = true;
            this.layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.layout.Dock = System.Windows.Forms.DockStyle.Top;
            this.layout.ColumnCount = 1;
            this.layout.RowCount = 1;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.Controls.Add(this.section, 0, 0);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Name = "layout";
            this.layout.Padding = new System.Windows.Forms.Padding(4);
            this.toolTips.SetToolTip(this.open, "Open in the VBE");
            this.Controls.Add(this.layout);
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "ChatAttachmentView";
            this.Size = new System.Drawing.Size(500, 120);
            this.layout.ResumeLayout(false); this.layout.PerformLayout();
            this.open.Symbol = VBAi.UiSymbol.Code;
            this.ResumeLayout(false); this.PerformLayout();
        }
    }
}

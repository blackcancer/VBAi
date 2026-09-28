namespace CodexVBE
{
    /// <summary>Displays selectable transcript text and Markdown with clickable references and code copying.</summary>
    public sealed partial class ChatTextContentView
    {
        /// <summary>Container that owns the disposable components created by the WinForms Designer.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>ToolTip component used to show full text for transcript controls.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Flow layout panel that contains this transcript view&apos;s child controls.</summary>
        private System.Windows.Forms.TableLayoutPanel layout;
        /// <summary>Stores the content used by ChatTextContentView.</summary>
        internal CodexVBE.UiRichTextBox content;
        /// <summary>Stores the copy menu used by ChatTextContentView.</summary>
        internal System.Windows.Forms.ContextMenuStrip copyMenu;
        /// <summary>Stores the copy selection used by ChatTextContentView.</summary>
        internal System.Windows.Forms.ToolStripMenuItem copySelection;
        /// <summary>Stores the copy code used by ChatTextContentView.</summary>
        internal System.Windows.Forms.ToolStripMenuItem copyCode;
        /// <summary>Releases the Designer components.</summary>
        /// <param name="disposing">Whether to release managed resources.</param>
        protected override void Dispose(bool disposing) { if (disposing) { DisposeTextResources(); components?.Dispose(); } base.Dispose(disposing); }
        /// <summary>Creates and configures the chat text content view controls serialized by the WinForms Designer.</summary>
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
            this.content = new CodexVBE.UiRichTextBox();
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

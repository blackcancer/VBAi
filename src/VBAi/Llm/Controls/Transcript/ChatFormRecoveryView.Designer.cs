namespace VBAi
{
    /// <summary>Displays recoverable changes made to a VBA form.</summary>
    public sealed partial class ChatFormRecoveryView
    {
        /// <summary>Container that owns the disposable components created by the WinForms Designer.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>ToolTip component used to show full text for transcript controls.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Flow layout panel that contains this transcript view&apos;s child controls.</summary>
        private ChatComposerPanel layout;
        /// <summary>Displays the form recovery prompt.</summary>
        internal System.Windows.Forms.Label title;
        /// <summary>Displays the number of recoverable form changes.</summary>
        internal System.Windows.Forms.Label count;
        /// <summary>Provides the action that restores the recovered form changes.</summary>
        internal ChatActionButton recover;
        /// <summary>Releases the Designer components.</summary>
        /// <param name="disposing">Whether to release managed resources.</param>
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }
        /// <summary>Creates and configures the chat form recovery view controls serialized by the WinForms Designer.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.layout = new ChatComposerPanel();
            this.title = new System.Windows.Forms.Label();
            this.title.Name = "title";
            this.title.AutoSize = true;
            this.title.Dock = System.Windows.Forms.DockStyle.Fill;
            this.title.Text = "Designer";
            this.count = new System.Windows.Forms.Label();
            this.count.Name = "count";
            this.count.AutoSize = true;
            this.count.Dock = System.Windows.Forms.DockStyle.Fill;
            this.count.Text = "Controls cut";
            this.recover = new ChatActionButton();
            this.recover.Name = "recover";
            this.recover.AutoSize = true;
            this.recover.Text = "Restore cut controls";
            this.SuspendLayout(); this.layout.SuspendLayout();
            this.layout.AutoSize = true;
            this.layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.layout.Dock = System.Windows.Forms.DockStyle.Top;
            this.layout.ColumnCount = 1;
            this.layout.RowCount = 3;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.Controls.Add(this.title, 0, 0);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.count, 0, 1);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.recover, 0, 2);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Name = "layout";
            this.layout.Padding = new System.Windows.Forms.Padding(10);
            this.toolTips.SetToolTip(this.recover, "Restore names, position, size and tab order. Replaces the clipboard. Refuses intervening changes.");
            this.Controls.Add(this.layout);
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "ChatFormRecoveryView";
            this.Size = new System.Drawing.Size(500, 120);
            this.layout.ResumeLayout(false); this.layout.PerformLayout();
            this.recover.Symbol = VBAi.UiSymbol.Undo;
            this.recover.IconOnly = true;
            this.recover.AutoSize = false;
            this.recover.Size = new System.Drawing.Size(32, 30);
            this.ResumeLayout(false); this.PerformLayout();
        }
    }
}

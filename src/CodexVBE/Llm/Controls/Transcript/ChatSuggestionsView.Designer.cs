namespace CodexVBE
{
    /// <summary>Displays navigation and reference suggestions for the current chat draft.</summary>
    public sealed partial class ChatSuggestionsView
    {
        /// <summary>Container that owns the disposable components created by the WinForms Designer.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>ToolTip component used to show full text for transcript controls.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Flow layout panel that contains this transcript view&apos;s child controls.</summary>
        private System.Windows.Forms.TableLayoutPanel layout;
        /// <summary>Contains the navigation targets offered for the current composer text.</summary>
        internal CodexVBE.UiListBox targets;
        /// <summary>Displays the current reference suggestion status.</summary>
        internal System.Windows.Forms.Label status;
        /// <summary>Releases the Designer components.</summary>
        /// <param name="disposing">Whether to release managed resources.</param>
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }
        /// <summary>Creates and configures the chat suggestions view controls serialized by the WinForms Designer.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.targets = new CodexVBE.UiListBox();
            this.targets.Name = "targets";
            this.targets.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.targets.ItemHeight = 30;
            this.targets.Dock = System.Windows.Forms.DockStyle.Fill;
            this.targets.Height = 180;
            this.targets.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.status = new System.Windows.Forms.Label();
            this.status.Name = "status";
            this.status.AutoSize = true;
            this.status.Dock = System.Windows.Forms.DockStyle.Fill;
            this.status.Text = "No matching target";
            this.SuspendLayout(); this.layout.SuspendLayout();
            this.layout.AutoSize = true;
            this.layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.layout.Dock = System.Windows.Forms.DockStyle.Top;
            this.layout.ColumnCount = 1;
            this.layout.RowCount = 2;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.Controls.Add(this.targets, 0, 0);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.status, 0, 1);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Name = "layout";
            this.layout.Padding = new System.Windows.Forms.Padding(10);
            this.Controls.Add(this.layout);
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "ChatSuggestionsView";
            this.Size = new System.Drawing.Size(500, 220);
            this.layout.ResumeLayout(false); this.layout.PerformLayout(); this.ResumeLayout(false); this.PerformLayout();
        }
    }
}

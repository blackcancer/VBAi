namespace CodexVBE
{
    public sealed partial class ChatContextChipView
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.FlowLayoutPanel layout;
        private CodexVBE.ChatActionButton open;
        private CodexVBE.ChatActionButton remove;
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants du modèle Designer.</summary>
        protected override void Dispose(bool disposing) { if (disposing && components != null) components.Dispose(); base.Dispose(disposing); }
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.layout = new System.Windows.Forms.FlowLayoutPanel();
            this.open = new CodexVBE.ChatActionButton();
            this.remove = new CodexVBE.ChatActionButton();
            this.layout.SuspendLayout();
            this.SuspendLayout();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            // open
            this.open.AutoSize = true;
            this.open.Cursor = System.Windows.Forms.Cursors.Hand;
            this.open.Location = new System.Drawing.Point(2, 2);
            this.open.Margin = new System.Windows.Forms.Padding(2);
            this.open.Name = "open";
            this.open.Padding = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.open.Size = new System.Drawing.Size(112, 26);
            this.open.TabIndex = 0;
            this.open.Text = "#Reference";
            this.open.Click += new System.EventHandler(this.Open_Click);
            // remove
            this.remove.AutoSize = true;
            this.remove.Cursor = System.Windows.Forms.Cursors.Hand;
            this.remove.Location = new System.Drawing.Point(118, 2);
            this.remove.Margin = new System.Windows.Forms.Padding(2);
            this.remove.Name = "remove";
            this.remove.Padding = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.remove.Size = new System.Drawing.Size(28, 26);
            this.remove.TabIndex = 1;
            this.remove.Text = "×";
            this.remove.Click += new System.EventHandler(this.Remove_Click);
            // layout
            this.layout.AutoSize = true;
            this.layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.layout.Controls.Add(this.open);
            this.layout.Controls.Add(this.remove);
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.Location = new System.Drawing.Point(0, 0);
            this.layout.Margin = new System.Windows.Forms.Padding(0);
            this.layout.Name = "layout";
            this.layout.Size = new System.Drawing.Size(150, 30);
            this.layout.TabIndex = 0;
            this.layout.WrapContents = false;
            // ChatContextChipView
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.layout);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "ChatContextChipView";
            this.Size = new System.Drawing.Size(150, 30);
            this.layout.ResumeLayout(false);
            this.layout.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

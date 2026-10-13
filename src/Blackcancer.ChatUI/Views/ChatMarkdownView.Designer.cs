// SPDX-License-Identifier: MPL-2.0
namespace Blackcancer.ChatUI.Views
{
    partial class ChatMarkdownView
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.FlowLayoutPanel blocks;
        private System.Windows.Forms.ContextMenuStrip sourceMenu;
        private System.Windows.Forms.ToolStripMenuItem copySource;
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.blocks = new System.Windows.Forms.FlowLayoutPanel();
            this.sourceMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.copySource = new System.Windows.Forms.ToolStripMenuItem();
            this.sourceMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // blocks
            // 
            this.blocks.AutoSize = true;
            this.blocks.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.blocks.ContextMenuStrip = this.sourceMenu;
            this.blocks.Dock = System.Windows.Forms.DockStyle.Top;
            this.blocks.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.blocks.Location = new System.Drawing.Point(0, 0);
            this.blocks.Margin = new System.Windows.Forms.Padding(0);
            this.blocks.Name = "blocks";
            this.blocks.Size = new System.Drawing.Size(500, 0);
            this.blocks.TabIndex = 0;
            this.blocks.WrapContents = false;
            // 
            // sourceMenu
            // 
            this.sourceMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.copySource});
            this.sourceMenu.Name = "sourceMenu";
            this.sourceMenu.Size = new System.Drawing.Size(163, 26);
            // 
            // copySource
            // 
            this.copySource.Name = "copySource";
            this.copySource.Size = new System.Drawing.Size(162, 22);
            this.copySource.Text = "Copy Markdown";
            // 
            // ChatMarkdownView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ContextMenuStrip = this.sourceMenu;
            this.Controls.Add(this.blocks);
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "ChatMarkdownView";
            this.Size = new System.Drawing.Size(500, 0);
            this.sourceMenu.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}

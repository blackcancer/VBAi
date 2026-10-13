// SPDX-License-Identifier: MPL-2.0
namespace Blackcancer.ChatUI.Views
{
    partial class ChatTextView
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.TableLayoutPanel layout;
        private System.Windows.Forms.RichTextBox content;
        private System.Windows.Forms.ContextMenuStrip copyMenu;
        private System.Windows.Forms.ToolStripMenuItem copySelection;
        protected override void Dispose(bool disposing) { if (disposing) { components?.Dispose(); DisposeFonts(); } base.Dispose(disposing); }
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            layout = new System.Windows.Forms.TableLayoutPanel();
            content = new System.Windows.Forms.RichTextBox();
            copyMenu = new System.Windows.Forms.ContextMenuStrip(components);
            copySelection = new System.Windows.Forms.ToolStripMenuItem();
            layout.SuspendLayout(); copyMenu.SuspendLayout(); SuspendLayout();
            layout.Name = "layout"; layout.AutoSize = true; layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            layout.Dock = System.Windows.Forms.DockStyle.Top; layout.ColumnCount = 1; layout.RowCount = 1;
            layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            layout.Controls.Add(content, 0, 0); layout.Margin = new System.Windows.Forms.Padding(0);
            content.Name = "content"; content.BorderStyle = System.Windows.Forms.BorderStyle.None;
            content.Dock = System.Windows.Forms.DockStyle.Top; content.ReadOnly = true; content.DetectUrls = false;
            content.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical; content.HideSelection = false;
            content.Margin = new System.Windows.Forms.Padding(0, 3, 0, 3); content.Size = new System.Drawing.Size(500, 24);
            content.ContextMenuStrip = copyMenu;
            copyMenu.Name = "copyMenu"; copyMenu.Items.Add(copySelection); copySelection.Name = "copySelection"; copySelection.Text = "Copy";
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font; AutoSize = true; AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            Controls.Add(layout); Name = "ChatTextView"; Size = new System.Drawing.Size(500, 30); Margin = new System.Windows.Forms.Padding(0);
            layout.ResumeLayout(false); layout.PerformLayout(); copyMenu.ResumeLayout(false); ResumeLayout(false); PerformLayout();
        }
    }
}

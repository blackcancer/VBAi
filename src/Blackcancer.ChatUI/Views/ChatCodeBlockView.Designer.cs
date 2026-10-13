// SPDX-License-Identifier: MPL-2.0
namespace Blackcancer.ChatUI.Views
{
    partial class ChatCodeBlockView
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.TableLayoutPanel layout, header;
        private System.Windows.Forms.Button toggle, copy;
        private ChatTextView code;
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container(); layout = new System.Windows.Forms.TableLayoutPanel(); header = new System.Windows.Forms.TableLayoutPanel();
            toggle = new System.Windows.Forms.Button(); copy = new System.Windows.Forms.Button(); code = new ChatTextView();
            layout.SuspendLayout(); header.SuspendLayout(); SuspendLayout();
            layout.Name = "layout"; layout.AutoSize = true; layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink; layout.Dock = System.Windows.Forms.DockStyle.Top;
            layout.ColumnCount = 1; layout.RowCount = 2; layout.Margin = new System.Windows.Forms.Padding(0);
            layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize)); layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            layout.Controls.Add(header, 0, 0); layout.Controls.Add(code, 0, 1);
            header.Name = "header"; header.AutoSize = true; header.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink; header.Dock = System.Windows.Forms.DockStyle.Top;
            header.ColumnCount = 2; header.RowCount = 1; header.Margin = new System.Windows.Forms.Padding(0);
            header.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F)); header.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            header.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize)); header.Controls.Add(toggle, 0, 0); header.Controls.Add(copy, 1, 0);
            toggle.Name = "toggle"; toggle.AutoSize = true; toggle.Dock = System.Windows.Forms.DockStyle.Fill; toggle.FlatStyle = System.Windows.Forms.FlatStyle.Flat; toggle.FlatAppearance.BorderSize = 0; toggle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft; toggle.Text = "▾ Code";
            copy.Name = "copy"; copy.AutoSize = true; copy.FlatStyle = System.Windows.Forms.FlatStyle.Flat; copy.Text = "Copy";
            code.Name = "code"; code.Dock = System.Windows.Forms.DockStyle.Top; code.Margin = new System.Windows.Forms.Padding(4, 0, 4, 4);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font; AutoSize = true; AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink; Controls.Add(layout);
            Name = "ChatCodeBlockView"; Size = new System.Drawing.Size(500, 70); Margin = new System.Windows.Forms.Padding(0, 4, 0, 4);
            header.ResumeLayout(false); header.PerformLayout(); layout.ResumeLayout(false); layout.PerformLayout(); ResumeLayout(false); PerformLayout();
        }
    }
}

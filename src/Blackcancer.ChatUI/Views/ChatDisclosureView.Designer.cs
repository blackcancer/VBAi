// SPDX-License-Identifier: MPL-2.0
namespace Blackcancer.ChatUI.Views
{
    partial class ChatDisclosureView
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.TableLayoutPanel layout;
        private System.Windows.Forms.Button toggle;
        private System.Windows.Forms.FlowLayoutPanel body;
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container(); layout = new System.Windows.Forms.TableLayoutPanel();
            toggle = new System.Windows.Forms.Button(); body = new System.Windows.Forms.FlowLayoutPanel();
            layout.SuspendLayout(); SuspendLayout();
            layout.Name = "layout"; layout.AutoSize = true; layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            layout.Dock = System.Windows.Forms.DockStyle.Top; layout.ColumnCount = 1; layout.RowCount = 2;
            layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize)); layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            layout.Controls.Add(toggle, 0, 0); layout.Controls.Add(body, 0, 1);
            toggle.Name = "toggle"; toggle.AutoSize = true; toggle.FlatStyle = System.Windows.Forms.FlatStyle.Flat; toggle.FlatAppearance.BorderSize = 0; toggle.Text = "▸ Details";
            toggle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft; toggle.Dock = System.Windows.Forms.DockStyle.Top;
            body.Name = "body"; body.AutoSize = true; body.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            body.Dock = System.Windows.Forms.DockStyle.Top; body.FlowDirection = System.Windows.Forms.FlowDirection.TopDown; body.WrapContents = false; body.Margin = new System.Windows.Forms.Padding(0); body.Visible = false;
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font; AutoSize = true; AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            Controls.Add(layout); Name = "ChatDisclosureView"; Size = new System.Drawing.Size(500, 32); Margin = new System.Windows.Forms.Padding(0);
            layout.ResumeLayout(false); layout.PerformLayout(); ResumeLayout(false); PerformLayout();
        }
    }
}

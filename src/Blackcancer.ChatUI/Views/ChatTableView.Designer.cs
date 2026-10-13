// SPDX-License-Identifier: MPL-2.0
namespace Blackcancer.ChatUI.Views
{
    partial class ChatTableView
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.TableLayoutPanel layout;
        private System.Windows.Forms.DataGridView grid;
        private System.Windows.Forms.Button copy;
        private System.Windows.Forms.Button expand;
        private System.Windows.Forms.FlowLayoutPanel actions;
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }
        private void InitializeComponent()
        {
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(ChatTableView));
            components = new System.ComponentModel.Container(); layout = new System.Windows.Forms.TableLayoutPanel(); grid = new System.Windows.Forms.DataGridView(); copy = new System.Windows.Forms.Button();
            expand = new System.Windows.Forms.Button(); actions = new System.Windows.Forms.FlowLayoutPanel();
            layout.SuspendLayout(); ((System.ComponentModel.ISupportInitialize)grid).BeginInit(); SuspendLayout();
            layout.Name = "layout"; layout.AutoSize = true; layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink; layout.Dock = System.Windows.Forms.DockStyle.Top;
            layout.ColumnCount = 1; layout.RowCount = 2; layout.Margin = new System.Windows.Forms.Padding(0);
            layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize)); layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            layout.Controls.Add(actions, 0, 0); layout.Controls.Add(grid, 0, 1);
            actions.Name = "actions"; actions.AutoSize = true; actions.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            actions.Dock = System.Windows.Forms.DockStyle.Fill; actions.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft; actions.WrapContents = true; actions.Margin = new System.Windows.Forms.Padding(0);
            actions.Controls.Add(expand); actions.Controls.Add(copy);
            expand.Name = "expand"; resources.ApplyResources(expand, "expand"); expand.AutoSize = true; expand.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            copy.Name = "copy"; copy.Text = "Copy table"; copy.AutoSize = true; copy.Anchor = System.Windows.Forms.AnchorStyles.Right; copy.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            grid.Name = "grid"; grid.Dock = System.Windows.Forms.DockStyle.Top; grid.ReadOnly = true; grid.VirtualMode = true;
            grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false; grid.AllowUserToOrderColumns = false; grid.AllowUserToResizeRows = false;
            grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.None; grid.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.None;
            grid.RowHeadersVisible = false; grid.EnableHeadersVisualStyles = false; grid.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            grid.ScrollBars = System.Windows.Forms.ScrollBars.Both; grid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            grid.MultiSelect = true; grid.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            grid.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            grid.Margin = new System.Windows.Forms.Padding(0); grid.Size = new System.Drawing.Size(500, 100);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font; AutoSize = true; AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            Controls.Add(layout); Name = "ChatTableView"; Size = new System.Drawing.Size(500, 136); Margin = new System.Windows.Forms.Padding(0, 4, 0, 4);
            ((System.ComponentModel.ISupportInitialize)grid).EndInit(); layout.ResumeLayout(false); layout.PerformLayout(); ResumeLayout(false); PerformLayout();
        }
    }
}

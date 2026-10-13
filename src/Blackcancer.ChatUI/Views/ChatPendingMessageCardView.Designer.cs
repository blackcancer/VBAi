// SPDX-License-Identifier: MPL-2.0
namespace Blackcancer.ChatUI.Views
{
    partial class ChatPendingMessageCardView
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.TableLayoutPanel layout;
        private System.Windows.Forms.FlowLayoutPanel toolbar;
        private System.Windows.Forms.Label positionLabel, entryStatus;
        private System.Windows.Forms.TextBox preview;
        private System.Windows.Forms.Button edit, remove, sendNext;
        private System.Windows.Forms.ToolTip toolTips;
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChatPendingMessageCardView));
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.toolbar = new System.Windows.Forms.FlowLayoutPanel();
            this.positionLabel = new System.Windows.Forms.Label();
            this.entryStatus = new System.Windows.Forms.Label();
            this.preview = new System.Windows.Forms.TextBox();
            this.edit = new System.Windows.Forms.Button();
            this.remove = new System.Windows.Forms.Button();
            this.sendNext = new System.Windows.Forms.Button();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.layout.SuspendLayout();
            this.toolbar.SuspendLayout();
            this.SuspendLayout();
            this.layout.AutoSize = true;
            this.layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.layout.ColumnCount = 1;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.RowCount = 4;
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 48F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.layout.Controls.Add(this.positionLabel, 0, 0);
            this.layout.Controls.Add(this.preview, 0, 1);
            this.layout.Controls.Add(this.entryStatus, 0, 2);
            this.layout.Controls.Add(this.toolbar, 0, 3);
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.Padding = new System.Windows.Forms.Padding(4);
            this.layout.Name = "layout";
            this.positionLabel.AutoSize = true;
            this.positionLabel.Name = "positionLabel";
            this.positionLabel.Text = "1";
            this.entryStatus.AutoSize = true;
            this.entryStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.entryStatus.Name = "entryStatus";
            this.preview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.preview.Multiline = true;
            this.preview.ReadOnly = true;
            this.preview.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.preview.MaxLength = 2147483647;
            this.preview.Name = "preview";
            this.preview.TabIndex = 0;
            this.toolbar.AutoSize = true;
            this.toolbar.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.toolbar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.toolbar.Controls.Add(this.edit);
            this.toolbar.Controls.Add(this.remove);
            this.toolbar.Controls.Add(this.sendNext);
            this.toolbar.Name = "toolbar";
            this.toolbar.TabIndex = 1;
            this.edit.AutoSize = true;
            this.edit.Enabled = false;
            this.edit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.edit.Name = "edit";
            this.edit.TabIndex = 0;
            resources.ApplyResources(this.edit, "edit");
            this.remove.AutoSize = true;
            this.remove.Enabled = false;
            this.remove.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.remove.Name = "remove";
            this.remove.TabIndex = 1;
            resources.ApplyResources(this.remove, "remove");
            this.sendNext.AutoSize = true;
            this.sendNext.Enabled = false;
            this.sendNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.sendNext.Name = "sendNext";
            this.sendNext.TabIndex = 2;
            resources.ApplyResources(this.sendNext, "sendNext");
            this.toolTips.SetToolTip(this.sendNext, resources.GetString("sendNextTip"));
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.layout);
            this.Margin = new System.Windows.Forms.Padding(0, 0, 0, 6);
            this.Name = "ChatPendingMessageCardView";
            this.Size = new System.Drawing.Size(400, 124);
            this.layout.ResumeLayout(false);
            this.layout.PerformLayout();
            this.toolbar.ResumeLayout(false);
            this.toolbar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

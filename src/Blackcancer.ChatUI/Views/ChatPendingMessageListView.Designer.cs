// SPDX-License-Identifier: MPL-2.0
namespace Blackcancer.ChatUI.Views
{
    partial class ChatPendingMessageListView
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.TableLayoutPanel layout;
        private System.Windows.Forms.Label header, statusNote, emptyNote;
        private System.Windows.Forms.FlowLayoutPanel messageFlow;
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
            if (disposing) DisposeOwnedAppearance();
        }
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChatPendingMessageListView));
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.header = new System.Windows.Forms.Label();
            this.statusNote = new System.Windows.Forms.Label();
            this.emptyNote = new System.Windows.Forms.Label();
            this.messageFlow = new System.Windows.Forms.FlowLayoutPanel();
            this.layout.SuspendLayout();
            this.SuspendLayout();
            this.layout.ColumnCount = 1;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.RowCount = 4;
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.Controls.Add(this.header, 0, 0);
            this.layout.Controls.Add(this.statusNote, 0, 1);
            this.layout.Controls.Add(this.emptyNote, 0, 2);
            this.layout.Controls.Add(this.messageFlow, 0, 3);
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.Name = "layout";
            this.layout.Padding = new System.Windows.Forms.Padding(4);
            this.header.AutoSize = true;
            this.header.Dock = System.Windows.Forms.DockStyle.Fill;
            this.header.Name = "header";
            resources.ApplyResources(this.header, "header");
            this.statusNote.AutoSize = true;
            this.statusNote.Dock = System.Windows.Forms.DockStyle.Fill;
            this.statusNote.Name = "statusNote";
            this.statusNote.Margin = new System.Windows.Forms.Padding(3, 0, 3, 2);
            this.emptyNote.AutoSize = true;
            this.emptyNote.Dock = System.Windows.Forms.DockStyle.Fill;
            this.emptyNote.Name = "emptyNote";
            resources.ApplyResources(this.emptyNote, "emptyNote");
            this.messageFlow.AutoScroll = true;
            this.messageFlow.Dock = System.Windows.Forms.DockStyle.Fill;
            this.messageFlow.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.messageFlow.WrapContents = false;
            this.messageFlow.Margin = new System.Windows.Forms.Padding(0);
            this.messageFlow.Name = "messageFlow";
            this.messageFlow.TabIndex = 0;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.layout);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.MinimumSize = new System.Drawing.Size(180, 100);
            this.Name = "ChatPendingMessageListView";
            this.Size = new System.Drawing.Size(600, 200);
            this.layout.ResumeLayout(false);
            this.layout.PerformLayout();
            this.ResumeLayout(false);
        }
    }
}

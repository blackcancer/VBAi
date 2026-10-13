// SPDX-License-Identifier: MPL-2.0
namespace Blackcancer.ChatUI.Views
{
    partial class ChatReferenceComparisonForm
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.TableLayoutPanel layout, readers, oldReader, newReader;
        private System.Windows.Forms.FlowLayoutPanel toolbar;
        private System.Windows.Forms.Button copyOld, copyNew, accept, cancel;
        private System.Windows.Forms.Label statusNote, note, oldTitle, newTitle, oldMetadataLabel, newMetadataLabel,
            oldContentLabel, newContentLabel, oldCaptureNote, newCaptureNote;
        private System.Windows.Forms.TextBox oldIdentity, newIdentity, oldMetadata, newMetadata, oldContent, newContent;

        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
            if (disposing) DisposeOwnedAppearance();
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChatReferenceComparisonForm));
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.readers = new System.Windows.Forms.TableLayoutPanel();
            this.oldReader = new System.Windows.Forms.TableLayoutPanel();
            this.newReader = new System.Windows.Forms.TableLayoutPanel();
            this.toolbar = new System.Windows.Forms.FlowLayoutPanel();
            this.copyOld = new System.Windows.Forms.Button();
            this.copyNew = new System.Windows.Forms.Button();
            this.accept = new System.Windows.Forms.Button();
            this.cancel = new System.Windows.Forms.Button();
            this.statusNote = new System.Windows.Forms.Label();
            this.note = new System.Windows.Forms.Label();
            this.oldTitle = new System.Windows.Forms.Label();
            this.newTitle = new System.Windows.Forms.Label();
            this.oldMetadataLabel = new System.Windows.Forms.Label();
            this.newMetadataLabel = new System.Windows.Forms.Label();
            this.oldContentLabel = new System.Windows.Forms.Label();
            this.newContentLabel = new System.Windows.Forms.Label();
            this.oldCaptureNote = new System.Windows.Forms.Label();
            this.newCaptureNote = new System.Windows.Forms.Label();
            this.oldIdentity = new System.Windows.Forms.TextBox();
            this.newIdentity = new System.Windows.Forms.TextBox();
            this.oldMetadata = new System.Windows.Forms.TextBox();
            this.newMetadata = new System.Windows.Forms.TextBox();
            this.oldContent = new System.Windows.Forms.TextBox();
            this.newContent = new System.Windows.Forms.TextBox();
            this.layout.SuspendLayout();
            this.readers.SuspendLayout();
            this.oldReader.SuspendLayout();
            this.newReader.SuspendLayout();
            this.toolbar.SuspendLayout();
            this.SuspendLayout();
            // layout
            this.layout.ColumnCount = 1;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.RowCount = 4;
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.Controls.Add(this.toolbar, 0, 0);
            this.layout.Controls.Add(this.statusNote, 0, 1);
            this.layout.Controls.Add(this.note, 0, 2);
            this.layout.Controls.Add(this.readers, 0, 3);
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.Name = "layout";
            this.layout.Padding = new System.Windows.Forms.Padding(12);
            this.layout.TabIndex = 0;
            // toolbar
            this.toolbar.AutoSize = true;
            this.toolbar.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.toolbar.Controls.Add(this.copyOld);
            this.toolbar.Controls.Add(this.copyNew);
            this.toolbar.Controls.Add(this.accept);
            this.toolbar.Controls.Add(this.cancel);
            this.toolbar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.toolbar.Margin = new System.Windows.Forms.Padding(0, 0, 0, 8);
            this.toolbar.Name = "toolbar";
            this.toolbar.TabIndex = 0;
            // copyOld
            this.copyOld.AutoSize = true;
            this.copyOld.Enabled = false;
            this.copyOld.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.copyOld.Name = "copyOld";
            this.copyOld.TabIndex = 0;
            resources.ApplyResources(this.copyOld, "copyOld");
            // copyNew
            this.copyNew.AutoSize = true;
            this.copyNew.Enabled = false;
            this.copyNew.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.copyNew.Name = "copyNew";
            this.copyNew.TabIndex = 1;
            resources.ApplyResources(this.copyNew, "copyNew");
            // accept
            this.accept.AutoSize = true;
            this.accept.Enabled = false;
            this.accept.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.accept.Name = "accept";
            this.accept.TabIndex = 2;
            resources.ApplyResources(this.accept, "accept");
            // cancel
            this.cancel.AutoSize = true;
            this.cancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cancel.Name = "cancel";
            this.cancel.TabIndex = 3;
            resources.ApplyResources(this.cancel, "cancel");
            // statusNote
            this.statusNote.AutoSize = true;
            this.statusNote.Dock = System.Windows.Forms.DockStyle.Fill;
            this.statusNote.Name = "statusNote";
            this.statusNote.Margin = new System.Windows.Forms.Padding(0, 0, 0, 6);
            resources.ApplyResources(this.statusNote, "statusNote");
            // note
            this.note.AutoSize = true;
            this.note.Dock = System.Windows.Forms.DockStyle.Fill;
            this.note.Name = "note";
            this.note.Margin = new System.Windows.Forms.Padding(0, 0, 0, 10);
            resources.ApplyResources(this.note, "note");
            // readers
            this.readers.ColumnCount = 2;
            this.readers.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.readers.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.readers.RowCount = 1;
            this.readers.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.readers.Controls.Add(this.oldReader, 0, 0);
            this.readers.Controls.Add(this.newReader, 1, 0);
            this.readers.Dock = System.Windows.Forms.DockStyle.Fill;
            this.readers.Margin = new System.Windows.Forms.Padding(0);
            this.readers.Name = "readers";
            this.readers.TabIndex = 1;
            // oldReader
            this.oldReader.ColumnCount = 1;
            this.oldReader.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.oldReader.RowCount = 7;
            this.oldReader.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.oldReader.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 64F));
            this.oldReader.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.oldReader.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.oldReader.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.oldReader.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.oldReader.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.oldReader.Controls.Add(this.oldTitle, 0, 0);
            this.oldReader.Controls.Add(this.oldIdentity, 0, 1);
            this.oldReader.Controls.Add(this.oldMetadataLabel, 0, 2);
            this.oldReader.Controls.Add(this.oldMetadata, 0, 3);
            this.oldReader.Controls.Add(this.oldCaptureNote, 0, 4);
            this.oldReader.Controls.Add(this.oldContentLabel, 0, 5);
            this.oldReader.Controls.Add(this.oldContent, 0, 6);
            this.oldReader.Dock = System.Windows.Forms.DockStyle.Fill;
            this.oldReader.Margin = new System.Windows.Forms.Padding(0, 0, 6, 0);
            this.oldReader.Name = "oldReader";
            this.oldReader.TabIndex = 0;
            // newReader
            this.newReader.ColumnCount = 1;
            this.newReader.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.newReader.RowCount = 7;
            this.newReader.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.newReader.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 64F));
            this.newReader.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.newReader.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.newReader.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.newReader.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.newReader.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.newReader.Controls.Add(this.newTitle, 0, 0);
            this.newReader.Controls.Add(this.newIdentity, 0, 1);
            this.newReader.Controls.Add(this.newMetadataLabel, 0, 2);
            this.newReader.Controls.Add(this.newMetadata, 0, 3);
            this.newReader.Controls.Add(this.newCaptureNote, 0, 4);
            this.newReader.Controls.Add(this.newContentLabel, 0, 5);
            this.newReader.Controls.Add(this.newContent, 0, 6);
            this.newReader.Dock = System.Windows.Forms.DockStyle.Fill;
            this.newReader.Margin = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.newReader.Name = "newReader";
            this.newReader.TabIndex = 1;
            // oldTitle
            this.oldTitle.AutoSize = true;
            this.oldTitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.oldTitle.Name = "oldTitle";
            resources.ApplyResources(this.oldTitle, "oldTitle");
            // newTitle
            this.newTitle.AutoSize = true;
            this.newTitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.newTitle.Name = "newTitle";
            resources.ApplyResources(this.newTitle, "newTitle");
            // oldIdentity
            this.oldIdentity.Dock = System.Windows.Forms.DockStyle.Fill;
            this.oldIdentity.Multiline = true;
            this.oldIdentity.ReadOnly = true;
            this.oldIdentity.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.oldIdentity.WordWrap = false;
            this.oldIdentity.MaxLength = 2147483647;
            this.oldIdentity.Name = "oldIdentity";
            this.oldIdentity.TabIndex = 0;
            // newIdentity
            this.newIdentity.Dock = System.Windows.Forms.DockStyle.Fill;
            this.newIdentity.Multiline = true;
            this.newIdentity.ReadOnly = true;
            this.newIdentity.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.newIdentity.WordWrap = false;
            this.newIdentity.MaxLength = 2147483647;
            this.newIdentity.Name = "newIdentity";
            this.newIdentity.TabIndex = 0;
            // oldMetadataLabel
            this.oldMetadataLabel.AutoSize = true;
            this.oldMetadataLabel.Name = "oldMetadataLabel";
            resources.ApplyResources(this.oldMetadataLabel, "oldMetadataLabel");
            // newMetadataLabel
            this.newMetadataLabel.AutoSize = true;
            this.newMetadataLabel.Name = "newMetadataLabel";
            resources.ApplyResources(this.newMetadataLabel, "newMetadataLabel");
            // oldMetadata
            this.oldMetadata.Dock = System.Windows.Forms.DockStyle.Fill;
            this.oldMetadata.Multiline = true;
            this.oldMetadata.ReadOnly = true;
            this.oldMetadata.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.oldMetadata.WordWrap = false;
            this.oldMetadata.MaxLength = 2147483647;
            this.oldMetadata.Name = "oldMetadata";
            this.oldMetadata.TabIndex = 1;
            // newMetadata
            this.newMetadata.Dock = System.Windows.Forms.DockStyle.Fill;
            this.newMetadata.Multiline = true;
            this.newMetadata.ReadOnly = true;
            this.newMetadata.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.newMetadata.WordWrap = false;
            this.newMetadata.MaxLength = 2147483647;
            this.newMetadata.Name = "newMetadata";
            this.newMetadata.TabIndex = 1;
            // oldCaptureNote
            this.oldCaptureNote.AutoSize = true;
            this.oldCaptureNote.Dock = System.Windows.Forms.DockStyle.Fill;
            this.oldCaptureNote.Name = "oldCaptureNote";
            resources.ApplyResources(this.oldCaptureNote, "oldCaptureNote");
            // newCaptureNote
            this.newCaptureNote.AutoSize = true;
            this.newCaptureNote.Dock = System.Windows.Forms.DockStyle.Fill;
            this.newCaptureNote.Name = "newCaptureNote";
            resources.ApplyResources(this.newCaptureNote, "newCaptureNote");
            // oldContentLabel
            this.oldContentLabel.AutoSize = true;
            this.oldContentLabel.Name = "oldContentLabel";
            resources.ApplyResources(this.oldContentLabel, "oldContentLabel");
            // newContentLabel
            this.newContentLabel.AutoSize = true;
            this.newContentLabel.Name = "newContentLabel";
            resources.ApplyResources(this.newContentLabel, "newContentLabel");
            // oldContent
            this.oldContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.oldContent.Multiline = true;
            this.oldContent.ReadOnly = true;
            this.oldContent.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.oldContent.WordWrap = false;
            this.oldContent.MaxLength = 2147483647;
            this.oldContent.Name = "oldContent";
            this.oldContent.TabIndex = 2;
            // newContent
            this.newContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.newContent.Multiline = true;
            this.newContent.ReadOnly = true;
            this.newContent.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.newContent.WordWrap = false;
            this.newContent.MaxLength = 2147483647;
            this.newContent.Name = "newContent";
            this.newContent.TabIndex = 2;
            // ChatReferenceComparisonForm
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1080, 700);
            this.Controls.Add(this.layout);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.MinimumSize = new System.Drawing.Size(700, 480);
            this.Name = "ChatReferenceComparisonForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            resources.ApplyResources(this, "$this");
            this.layout.ResumeLayout(false);
            this.layout.PerformLayout();
            this.readers.ResumeLayout(false);
            this.oldReader.ResumeLayout(false);
            this.oldReader.PerformLayout();
            this.newReader.ResumeLayout(false);
            this.newReader.PerformLayout();
            this.toolbar.ResumeLayout(false);
            this.toolbar.PerformLayout();
            this.ResumeLayout(false);
        }
    }
}

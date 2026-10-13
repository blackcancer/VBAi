// SPDX-License-Identifier: MPL-2.0
namespace Blackcancer.ChatUI.Views
{
    partial class ChatReadingOptionsForm
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.TableLayoutPanel layout;
        private System.Windows.Forms.Label modeLabel, filterLabel, queryLabel, explanation;
        private System.Windows.Forms.ComboBox modePicker, filterPicker;
        private System.Windows.Forms.TextBox contextQuery;
        private System.Windows.Forms.FlowLayoutPanel actions;
        private System.Windows.Forms.Button saveButton, cancelButton;
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); if (disposing) DisposeOwnedAppearance(); }
        private void InitializeComponent()
        {
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.modeLabel = new System.Windows.Forms.Label();
            this.modePicker = new System.Windows.Forms.ComboBox();
            this.filterLabel = new System.Windows.Forms.Label();
            this.filterPicker = new System.Windows.Forms.ComboBox();
            this.queryLabel = new System.Windows.Forms.Label();
            this.contextQuery = new System.Windows.Forms.TextBox();
            this.explanation = new System.Windows.Forms.Label();
            this.actions = new System.Windows.Forms.FlowLayoutPanel();
            this.cancelButton = new System.Windows.Forms.Button();
            this.saveButton = new System.Windows.Forms.Button();
            this.layout.SuspendLayout();
            this.actions.SuspendLayout();
            this.SuspendLayout();
            // 
            // layout
            // 
            this.layout.ColumnCount = 2;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.Controls.Add(this.modeLabel, 0, 0);
            this.layout.Controls.Add(this.modePicker, 1, 0);
            this.layout.Controls.Add(this.filterLabel, 0, 1);
            this.layout.Controls.Add(this.filterPicker, 1, 1);
            this.layout.Controls.Add(this.queryLabel, 0, 2);
            this.layout.Controls.Add(this.contextQuery, 1, 2);
            this.layout.Controls.Add(this.explanation, 0, 3);
            this.layout.Controls.Add(this.actions, 0, 4);
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.Location = new System.Drawing.Point(0, 0);
            this.layout.Name = "layout";
            this.layout.Padding = new System.Windows.Forms.Padding(12);
            this.layout.RowCount = 5;
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.layout.Size = new System.Drawing.Size(500, 260);
            this.layout.TabIndex = 0;
            // 
            // modeLabel
            // 
            this.modeLabel.AutoSize = true;
            this.modeLabel.Location = new System.Drawing.Point(15, 12);
            this.modeLabel.Name = "modeLabel";
            this.modeLabel.Size = new System.Drawing.Size(84, 15);
            this.modeLabel.TabIndex = 0;
            this.modeLabel.Text = "Reading mode";
            // 
            // modePicker
            // 
            this.modePicker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.modePicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.modePicker.FormattingEnabled = true;
            this.modePicker.Location = new System.Drawing.Point(135, 15);
            this.modePicker.Name = "modePicker";
            this.modePicker.Size = new System.Drawing.Size(350, 23);
            this.modePicker.TabIndex = 1;
            // 
            // filterLabel
            // 
            this.filterLabel.AutoSize = true;
            this.filterLabel.Location = new System.Drawing.Point(15, 39);
            this.filterLabel.Name = "filterLabel";
            this.filterLabel.Size = new System.Drawing.Size(73, 15);
            this.filterLabel.TabIndex = 2;
            this.filterLabel.Text = "Activity type";
            // 
            // filterPicker
            // 
            this.filterPicker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.filterPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.filterPicker.FormattingEnabled = true;
            this.filterPicker.Location = new System.Drawing.Point(135, 42);
            this.filterPicker.Name = "filterPicker";
            this.filterPicker.Size = new System.Drawing.Size(350, 23);
            this.filterPicker.TabIndex = 3;
            // 
            // queryLabel
            // 
            this.queryLabel.AutoSize = true;
            this.queryLabel.Location = new System.Drawing.Point(15, 66);
            this.queryLabel.Name = "queryLabel";
            this.queryLabel.Size = new System.Drawing.Size(114, 15);
            this.queryLabel.TabIndex = 4;
            this.queryLabel.Text = "Search captured text";
            // 
            // contextQuery
            // 
            this.contextQuery.Dock = System.Windows.Forms.DockStyle.Fill;
            this.contextQuery.Location = new System.Drawing.Point(135, 69);
            this.contextQuery.MaxLength = 256;
            this.contextQuery.Name = "contextQuery";
            this.contextQuery.Size = new System.Drawing.Size(350, 23);
            this.contextQuery.TabIndex = 5;
            // 
            // explanation
            // 
            this.explanation.AutoSize = true;
            this.layout.SetColumnSpan(this.explanation, 2);
            this.explanation.Dock = System.Windows.Forms.DockStyle.Fill;
            this.explanation.Location = new System.Drawing.Point(15, 107);
            this.explanation.Margin = new System.Windows.Forms.Padding(3, 12, 3, 12);
            this.explanation.Name = "explanation";
            this.explanation.Size = new System.Drawing.Size(470, 92);
            this.explanation.TabIndex = 6;
            this.explanation.Text = "These choices change how captured activities are displayed. Their meaning and sto" +
    "rage are supplied by the host.";
            // 
            // actions
            // 
            this.actions.AutoSize = true;
            this.layout.SetColumnSpan(this.actions, 2);
            this.actions.Controls.Add(this.cancelButton);
            this.actions.Controls.Add(this.saveButton);
            this.actions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.actions.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.actions.Location = new System.Drawing.Point(15, 214);
            this.actions.Name = "actions";
            this.actions.Size = new System.Drawing.Size(470, 31);
            this.actions.TabIndex = 7;
            // 
            // cancelButton
            // 
            this.cancelButton.AutoSize = true;
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(392, 3);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 25);
            this.cancelButton.TabIndex = 0;
            this.cancelButton.Text = "Cancel";
            // 
            // saveButton
            // 
            this.saveButton.AutoSize = true;
            this.saveButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.saveButton.Location = new System.Drawing.Point(311, 3);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(75, 25);
            this.saveButton.TabIndex = 1;
            this.saveButton.Text = "Save";
            // 
            // ChatReadingOptionsForm
            // 
            this.AcceptButton = this.saveButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(500, 260);
            this.Controls.Add(this.layout);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(440, 290);
            this.Name = "ChatReadingOptionsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Reading options";
            this.layout.ResumeLayout(false);
            this.layout.PerformLayout();
            this.actions.ResumeLayout(false);
            this.actions.PerformLayout();
            this.ResumeLayout(false);

        }
    }
}

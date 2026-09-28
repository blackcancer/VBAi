namespace CodexVBE
{
    public sealed partial class GitConnectionView
    {
        internal System.Windows.Forms.Label remoteLabel;
        internal System.Windows.Forms.Label branchLabel;
        internal System.Windows.Forms.Label help;
        internal System.Windows.Forms.TextBox remote;
        internal System.Windows.Forms.TextBox branch;
        internal System.Windows.Forms.Button connect;
        internal System.Windows.Forms.TableLayoutPanel connectionLayout;
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.ToolTip toolTips;
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.connectionLayout = new System.Windows.Forms.TableLayoutPanel();
            this.remoteLabel = new System.Windows.Forms.Label();
            this.branchLabel = new System.Windows.Forms.Label();
            this.help = new System.Windows.Forms.Label();
            this.remote = new System.Windows.Forms.TextBox();
            this.branch = new System.Windows.Forms.TextBox();
            this.connect = new CodexVBE.ThemedButton();
            this.SuspendLayout();
            this.help.Text = "VBA changes become source files in the repository. No folder next to the macro.\r\nThe first pull replaces VBA with a restorable backup.\r\nSign in through Git Credential Manager; commit identity is configured in Git.";
            this.help.Dock = System.Windows.Forms.DockStyle.Fill;
            this.help.AutoSize = true;
            this.help.Name = "help";
            this.remoteLabel.Text = "GitHub repository";
            this.remoteLabel.AutoSize = true;
            this.remoteLabel.Name = "remoteLabel";
            this.remote.Dock = System.Windows.Forms.DockStyle.Fill;
            this.remote.Name = "remote";
            this.remote.TabIndex = 0;
            this.toolTips.SetToolTip(this.remote, "HTTPS URL, for example https://github.com/organization/macros.git. Do not paste a token.");
            this.branchLabel.Text = "Branch";
            this.branchLabel.AutoSize = true;
            this.branchLabel.Name = "branchLabel";
            this.branch.Text = "main";
            this.branch.Dock = System.Windows.Forms.DockStyle.Fill;
            this.branch.Name = "branch";
            this.branch.TabIndex = 1;
            this.toolTips.SetToolTip(this.branch, "Branch to synchronize. Pushes are never forced.");
            this.connect.Text = "Link repository";
            this.connect.AutoSize = true;
            this.connect.Name = "connect";
            this.connect.TabIndex = 2;
            this.toolTips.SetToolTip(this.connect, "Save the link for this document on this computer. Does not publish code.");
            this.connectionLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.connectionLayout.ColumnCount = 2;
            this.connectionLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 140F));
            this.connectionLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.connectionLayout.RowCount = 4;
            this.connectionLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.connectionLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.connectionLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.connectionLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.connectionLayout.Controls.Add(this.help, 0, 0);
            this.connectionLayout.SetColumnSpan(this.help, 2);
            this.connectionLayout.Controls.Add(this.remoteLabel, 0, 1);
            this.connectionLayout.Controls.Add(this.remote, 1, 1);
            this.connectionLayout.Controls.Add(this.branchLabel, 0, 2);
            this.connectionLayout.Controls.Add(this.branch, 1, 2);
            this.connectionLayout.Controls.Add(this.connect, 1, 3);
            this.Controls.Add(this.connectionLayout);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitConnectionView";
            this.Size = new System.Drawing.Size(860, 500);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

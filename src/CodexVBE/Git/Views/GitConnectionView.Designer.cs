namespace CodexVBE
{
    /// <summary>Designer-generated repository remote and branch connection controls.</summary>
    public sealed partial class GitConnectionView
    {
        /// <summary>Caption for the remote URL input.</summary>
        internal System.Windows.Forms.Label remoteLabel;
        /// <summary>Caption for the current branch input.</summary>
        internal System.Windows.Forms.Label branchLabel;
        /// <summary>Help text for connecting a repository.</summary>
        internal System.Windows.Forms.Label help;
        /// <summary>Remote repository URL.</summary>
        internal System.Windows.Forms.TextBox remote;
        /// <summary>Current branch name.</summary>
        internal System.Windows.Forms.TextBox branch;
        /// <summary>Connects to the configured repository.</summary>
        internal CodexVBE.ThemedButton connect;
        /// <summary>Layout for remote, branch, help, and connection action.</summary>
        internal System.Windows.Forms.TableLayoutPanel connectionLayout;
        /// <summary>Container that owns Designer components.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Tooltips associated with connection controls.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Creates and arranges repository connection controls.</summary>
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
            this.connectionLayout.SuspendLayout();
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
            this.remoteLabel.Location = new System.Drawing.Point(3, 53);
            this.remoteLabel.Size = new System.Drawing.Size(102, 21);
            this.remoteLabel.TabIndex = 1;
            this.branchLabel.Location = new System.Drawing.Point(3, 89);
            this.branchLabel.Size = new System.Drawing.Size(42, 21);
            this.branchLabel.TabIndex = 2;
            this.help.Location = new System.Drawing.Point(3, 0);
            this.help.Size = new System.Drawing.Size(854, 53);
            this.help.TabIndex = 0;
            this.remote.Location = new System.Drawing.Point(143, 56);
            this.remote.Size = new System.Drawing.Size(714, 23);
            this.branch.Location = new System.Drawing.Point(143, 92);
            this.branch.Size = new System.Drawing.Size(714, 23);
            this.connect.Location = new System.Drawing.Point(143, 128);
            this.connect.Size = new System.Drawing.Size(95, 28);
            this.connectionLayout.Name = "connectionLayout";
            this.connectionLayout.Location = new System.Drawing.Point(0, 0);
            this.connectionLayout.Size = new System.Drawing.Size(860, 500);
            this.connectionLayout.TabIndex = 0;
            this.connectionLayout.ResumeLayout(false);
            this.connectionLayout.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

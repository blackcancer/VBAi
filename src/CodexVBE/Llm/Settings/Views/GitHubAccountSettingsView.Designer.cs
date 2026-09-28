namespace CodexVBE
{
    public sealed partial class GitHubAccountSettingsView
    {
        internal System.Windows.Forms.Label githubLabel;
        internal System.Windows.Forms.Label githubStatus;
        internal System.Windows.Forms.Label githubAccountLabel;
        internal System.Windows.Forms.Label githubNote;
        internal System.Windows.Forms.ComboBox githubAccount;
        internal System.Windows.Forms.FlowLayoutPanel githubActions;
        internal System.Windows.Forms.Button githubLogin;
        internal System.Windows.Forms.Button githubRefresh;
        internal System.Windows.Forms.ToolTip githubToolTips;
        internal System.Windows.Forms.TableLayoutPanel accountLayout;
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
            this.accountLayout = new System.Windows.Forms.TableLayoutPanel();
            this.githubLabel = new System.Windows.Forms.Label();
            this.githubStatus = new System.Windows.Forms.Label();
            this.githubAccountLabel = new System.Windows.Forms.Label();
            this.githubNote = new System.Windows.Forms.Label();
            this.githubAccount = new CodexVBE.ThemedComboBox();
            this.githubActions = new System.Windows.Forms.FlowLayoutPanel();
            this.githubLogin = new CodexVBE.ThemedButton();
            this.githubRefresh = new CodexVBE.ThemedButton();
            this.githubToolTips = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            this.githubLabel.Name = "githubLabel";
            this.githubLabel.Text = "GitHub · repositories";
            this.githubLabel.AutoSize = true;
            this.githubLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.githubStatus.Name = "githubStatus";
            this.githubStatus.Text = "Connect an account to synchronize your VBA sources.";
            this.githubStatus.AutoSize = true;
            this.githubStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.githubAccountLabel.Name = "githubAccountLabel";
            this.githubAccountLabel.Text = "GitHub account";
            this.githubAccountLabel.AutoSize = true;
            this.githubAccount.Name = "githubAccount";
            this.githubAccount.AccessibleName = "GitHub account for VBA repositories";
            this.githubAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.githubAccount.Dock = System.Windows.Forms.DockStyle.Fill;
            this.githubAccount.TabIndex = 10;
            this.githubActions.Name = "githubActions";
            this.githubActions.AutoSize = true;
            this.githubActions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.githubActions.Controls.Add(this.githubLogin);
            this.githubActions.Controls.Add(this.githubRefresh);
            this.githubLogin.Name = "githubLogin";
            this.githubLogin.Text = "Sign in to GitHub";
            this.githubLogin.AutoSize = true;
            this.githubLogin.MinimumSize = new System.Drawing.Size(150, 30);
            this.githubLogin.TabIndex = 11;
            this.githubRefresh.Name = "githubRefresh";
            this.githubRefresh.Text = "Refresh accounts";
            this.githubRefresh.AutoSize = true;
            this.githubRefresh.MinimumSize = new System.Drawing.Size(140, 30);
            this.githubRefresh.TabIndex = 12;
            this.githubNote.Name = "githubNote";
            this.githubNote.Text = "Credentials stay in Git Credential Manager. Sign-in takes effect immediately; Save keeps the chosen account for VBAi. Copilot authentication is configured separately.";
            this.githubNote.AutoSize = true;
            this.githubNote.Dock = System.Windows.Forms.DockStyle.Fill;
            this.githubToolTips.SetToolTip(this.githubLogin, "Opens GitHub authentication in your browser. No token to copy into VBAi.");
            this.githubToolTips.SetToolTip(this.githubRefresh, "Reads locally saved accounts; does not check repository permissions yet.");
            this.githubToolTips.SetToolTip(this.githubAccount, "Account used for subsequent macro fetch, pull and push operations. Does not change global Git configuration.");
            this.githubLabel.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.githubStatus.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.githubAccountLabel.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.githubAccount.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.githubActions.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.githubNote.Margin = new System.Windows.Forms.Padding(0, 4, 12, 8);
            this.githubLabel.MinimumSize = new System.Drawing.Size(156, 0);
            this.githubAccountLabel.MinimumSize = new System.Drawing.Size(156, 0);
            this.accountLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.accountLayout.AutoSize = true;
            this.accountLayout.ColumnCount = 2;
            this.accountLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            this.accountLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.accountLayout.RowCount = 4;
            this.accountLayout.Padding = new System.Windows.Forms.Padding(12);
            this.accountLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.accountLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.accountLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.accountLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.accountLayout.Controls.Add(this.githubLabel, 0, 0);
            this.accountLayout.Controls.Add(this.githubStatus, 1, 0);
            this.accountLayout.Controls.Add(this.githubAccountLabel, 0, 1);
            this.accountLayout.Controls.Add(this.githubAccount, 1, 1);
            this.accountLayout.Controls.Add(this.githubActions, 1, 2);
            this.accountLayout.Controls.Add(this.githubNote, 1, 3);
            this.AutoScroll = true;
            this.Controls.Add(this.accountLayout);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitHubAccountSettingsView";
            this.Size = new System.Drawing.Size(860, 500);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

namespace CodexVBE
{
    /// <summary>Designer-generated controls and layout for update settings and release actions.</summary>
    partial class UpdateWindow
    {
        /// <summary>Container that owns the Designer components.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Root layout for release metadata, preferences, and actions.</summary>
        private System.Windows.Forms.TableLayoutPanel layout;
        /// <summary>Window heading.</summary>
        private System.Windows.Forms.Label heading;
        /// <summary>Current and available version text.</summary>
        private System.Windows.Forms.Label version;
        /// <summary>Release description and notes.</summary>
        private System.Windows.Forms.Label description;
        /// <summary>Automatic release-check preference.</summary>
        private System.Windows.Forms.CheckBox automaticCheck;
        /// <summary>Automatic installer-download preference.</summary>
        private System.Windows.Forms.CheckBox automaticDownload;
        /// <summary>Automatic installer-launch preference.</summary>
        private System.Windows.Forms.CheckBox automaticInstall;
        /// <summary>Prerelease eligibility preference.</summary>
        private System.Windows.Forms.CheckBox previews;
        /// <summary>Release notes section caption.</summary>
        private System.Windows.Forms.Label notesLabel;
        /// <summary>Secondary release action buttons.</summary>
        private System.Windows.Forms.FlowLayoutPanel secondaryButtons;
        /// <summary>Read-only release notes pane.</summary>
        private CodexVBE.UiTextBox notes;
        /// <summary>Status and download feedback.</summary>
        private System.Windows.Forms.Label status;
        /// <summary>Check and download progress indicator.</summary>
        private System.Windows.Forms.ProgressBar progress;
        /// <summary>Primary and secondary update actions.</summary>
        private System.Windows.Forms.FlowLayoutPanel buttons;
        /// <summary>Closes the update window.</summary>
        private CodexVBE.UiActionButton close;
        /// <summary>Checks for a newer release.</summary>
        private CodexVBE.UiActionButton check;
        /// <summary>Downloads the selected installer.</summary>
        private CodexVBE.UiActionButton download;
        /// <summary>Schedules installation of the staged package.</summary>
        private CodexVBE.UiActionButton install;
        /// <summary>Skips the selected release.</summary>
        private CodexVBE.UiActionButton skip;
        /// <summary>Saves the selected preferences.</summary>
        private CodexVBE.UiActionButton save;
        /// <summary>Cancels a scheduled installation that has not started.</summary>
        private CodexVBE.UiActionButton cancelPending;
        /// <summary>Tooltips associated with update controls.</summary>
        private System.Windows.Forms.ToolTip tips;
        /// <summary>Releases runtime subscriptions and Designer-owned components.</summary><param name="disposing">Whether managed components should be disposed.</param>
        protected override void Dispose(bool disposing) { if (disposing) { DisposeRuntime(); if (components != null) components.Dispose(); } base.Dispose(disposing); }
        /// <summary>Creates and arranges the update window controls.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.heading = new System.Windows.Forms.Label();
            this.version = new System.Windows.Forms.Label();
            this.description = new System.Windows.Forms.Label();
            this.automaticCheck = new System.Windows.Forms.CheckBox();
            this.automaticDownload = new System.Windows.Forms.CheckBox();
            this.automaticInstall = new System.Windows.Forms.CheckBox();
            this.previews = new System.Windows.Forms.CheckBox();
            this.notesLabel = new System.Windows.Forms.Label();
            this.secondaryButtons = new System.Windows.Forms.FlowLayoutPanel();
            this.notes = new CodexVBE.UiTextBox();
            this.status = new System.Windows.Forms.Label();
            this.progress = new System.Windows.Forms.ProgressBar();
            this.buttons = new System.Windows.Forms.FlowLayoutPanel();
            this.close = new CodexVBE.UiActionButton();
            this.check = new CodexVBE.UiActionButton();
            this.download = new CodexVBE.UiActionButton();
            this.install = new CodexVBE.UiActionButton();
            this.skip = new CodexVBE.UiActionButton();
            this.save = new CodexVBE.UiActionButton();
            this.cancelPending = new CodexVBE.UiActionButton();
            this.tips = new System.Windows.Forms.ToolTip(this.components);
            this.layout.SuspendLayout();
            this.buttons.SuspendLayout();
            this.SuspendLayout();
            this.layout.Name = "layout";
            this.heading.Name = "heading";
            this.version.Name = "version";
            this.description.Name = "description";
            this.automaticCheck.Name = "automaticCheck";
            this.automaticDownload.Name = "automaticDownload";
            this.automaticInstall.Paint += new System.Windows.Forms.PaintEventHandler(this.PaintDisabled);
            this.automaticInstall.Name = "automaticInstall";
            this.previews.Name = "previews";
            this.notes.Name = "notes";
            this.status.Name = "status";
            this.progress.Name = "progress";
            this.buttons.Name = "buttons";
            this.close.Paint += new System.Windows.Forms.PaintEventHandler(this.PaintDisabled);
            this.close.Name = "close";
            this.check.Paint += new System.Windows.Forms.PaintEventHandler(this.PaintDisabled);
            this.check.Name = "check";
            this.download.Paint += new System.Windows.Forms.PaintEventHandler(this.PaintDisabled);
            this.download.Name = "download";
            this.install.Paint += new System.Windows.Forms.PaintEventHandler(this.PaintDisabled);
            this.install.Name = "install";
            this.skip.Paint += new System.Windows.Forms.PaintEventHandler(this.PaintDisabled);
            this.skip.Name = "skip";
            this.save.Paint += new System.Windows.Forms.PaintEventHandler(this.PaintDisabled);
            this.save.Name = "save";
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.Padding = new System.Windows.Forms.Padding(16);
            this.layout.ColumnCount = 1;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.RowCount = 13;
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.heading, 0, 0);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.version, 0, 1);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.description, 0, 2);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.automaticCheck, 0, 3);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.automaticDownload, 0, 4);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.automaticInstall, 0, 5);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.previews, 0, 6);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.notesLabel, 0, 7);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.Controls.Add(this.notes, 0, 8);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.status, 0, 9);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.layout.Controls.Add(this.progress, 0, 10);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.buttons, 0, 11);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.secondaryButtons, 0, 12);
            this.heading.Text = "VBAi updates";
            this.description.Text = "GitHub · Windows x64 · SHA-256 · Authenticode";
            this.automaticCheck.Text = "Check automatically once a day";
            this.automaticDownload.Text = "Download updates automatically";
            this.automaticInstall.Text = "Install automatically after closing VBA hosts";
            this.previews.Text = "Include prereleases";
            this.heading.AutoSize = true;
            this.heading.Dock = System.Windows.Forms.DockStyle.Fill;
            this.heading.Margin = new System.Windows.Forms.Padding(0, 0, 0, 10);
            this.version.AutoSize = true;
            this.version.Dock = System.Windows.Forms.DockStyle.Fill;
            this.version.Margin = new System.Windows.Forms.Padding(0, 0, 0, 10);
            this.description.AutoSize = true;
            this.description.Dock = System.Windows.Forms.DockStyle.Fill;
            this.description.Margin = new System.Windows.Forms.Padding(0, 0, 0, 10);
            this.automaticCheck.AutoSize = true;
            this.automaticCheck.Dock = System.Windows.Forms.DockStyle.Fill;
            this.automaticCheck.Margin = new System.Windows.Forms.Padding(0, 0, 0, 10);
            this.automaticDownload.AutoSize = true;
            this.automaticDownload.Dock = System.Windows.Forms.DockStyle.Fill;
            this.automaticDownload.Margin = new System.Windows.Forms.Padding(0, 0, 0, 10);
            this.automaticInstall.AutoSize = true;
            this.automaticInstall.Dock = System.Windows.Forms.DockStyle.Fill;
            this.automaticInstall.Margin = new System.Windows.Forms.Padding(0, 0, 0, 10);
            this.previews.AutoSize = true;
            this.previews.Dock = System.Windows.Forms.DockStyle.Fill;
            this.previews.Margin = new System.Windows.Forms.Padding(0, 0, 0, 10);
            this.status.AutoSize = true;
            this.status.Dock = System.Windows.Forms.DockStyle.Fill;
            this.status.Margin = new System.Windows.Forms.Padding(0, 0, 0, 10);
            this.heading.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold);
            this.automaticCheck.Checked = true;
            this.automaticDownload.Checked = true;
            this.automaticDownload.CheckedChanged += new System.EventHandler(this.DownloadPreferenceChanged);
            this.notesLabel.Name = "notesLabel";
            this.notesLabel.Text = "Release notes";
            this.notesLabel.AutoSize = true;
            this.notesLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.notesLabel.Margin = new System.Windows.Forms.Padding(0, 6, 0, 8);
            this.secondaryButtons.Name = "secondaryButtons";
            this.secondaryButtons.AutoSize = true;
            this.secondaryButtons.Dock = System.Windows.Forms.DockStyle.Fill;
            this.secondaryButtons.FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight;
            this.secondaryButtons.WrapContents = true;
            this.notes.Multiline = true;
            this.notes.ReadOnly = true;
            this.notes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.notes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.notes.MinimumSize = new System.Drawing.Size(0, 150);
            this.status.MinimumSize = new System.Drawing.Size(0, 38);
            this.status.AccessibleRole = System.Windows.Forms.AccessibleRole.StatusBar;
            this.progress.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progress.Margin = System.Windows.Forms.Padding.Empty;
            this.progress.Visible = false;
            this.buttons.Dock = System.Windows.Forms.DockStyle.Fill;
            this.buttons.AutoSize = true;
            this.buttons.WrapContents = true;
            this.buttons.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.buttons.Controls.Add(this.close);
            this.close.Text = "Close";
            this.close.AutoSize = true;
            this.close.MinimumSize = new System.Drawing.Size(85, 36);
            this.close.Padding = new System.Windows.Forms.Padding(8, 4, 8, 4);
            this.close.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.tips.SetToolTip(this.close, "Close");
            this.buttons.Controls.Add(this.install);
            this.install.Text = "Install update";
            this.install.AutoSize = true;
            this.install.MinimumSize = new System.Drawing.Size(85, 36);
            this.install.Padding = new System.Windows.Forms.Padding(8, 4, 8, 4);
            this.install.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.tips.SetToolTip(this.install, "Install update");
            this.install.Click += new System.EventHandler(this.Install_Click);
            this.buttons.Controls.Add(this.download);
            this.download.Text = "Download";
            this.download.AutoSize = true;
            this.download.MinimumSize = new System.Drawing.Size(85, 36);
            this.download.Padding = new System.Windows.Forms.Padding(8, 4, 8, 4);
            this.download.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.tips.SetToolTip(this.download, "Download");
            this.download.Click += new System.EventHandler(this.Download_Click);
            this.buttons.Controls.Add(this.check);
            this.check.Text = "Check for updates";
            this.check.AutoSize = true;
            this.check.MinimumSize = new System.Drawing.Size(85, 36);
            this.check.Padding = new System.Windows.Forms.Padding(8, 4, 8, 4);
            this.check.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.tips.SetToolTip(this.check, "Check for updates");
            this.check.Click += new System.EventHandler(this.Check_Click);

            this.skip.Text = "Skip version";
            this.skip.AutoSize = true;
            this.skip.MinimumSize = new System.Drawing.Size(85, 36);
            this.skip.Padding = new System.Windows.Forms.Padding(8, 4, 8, 4);
            this.skip.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.tips.SetToolTip(this.skip, "Skip version");
            this.skip.Click += new System.EventHandler(this.Skip_Click);
            this.secondaryButtons.Controls.Add(this.save);
            this.secondaryButtons.Controls.Add(this.skip);
            this.save.Text = "Save";
            this.save.AutoSize = true;
            this.save.MinimumSize = new System.Drawing.Size(85, 36);
            this.save.Padding = new System.Windows.Forms.Padding(8, 4, 8, 4);
            this.save.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.tips.SetToolTip(this.save, "Save");
            this.save.Click += new System.EventHandler(this.Save_Click);
            this.secondaryButtons.Controls.Add(this.cancelPending);
            this.cancelPending.Paint += new System.Windows.Forms.PaintEventHandler(this.PaintDisabled);
            this.cancelPending.Name = "cancelPending";
            this.cancelPending.Text = "Cancel scheduled update";
            this.cancelPending.Enabled = false;
            this.cancelPending.AutoSize = true;
            this.cancelPending.MinimumSize = new System.Drawing.Size(85, 36);
            this.cancelPending.Padding = new System.Windows.Forms.Padding(8, 4, 8, 4);
            this.cancelPending.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cancelPending.Click += new System.EventHandler(this.CancelPending_Click);
            this.tips.SetToolTip(this.cancelPending, "Cancel scheduled update");
            this.download.Enabled = false;
            this.install.Enabled = false;
            this.skip.Enabled = false;
            this.close.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.tips.SetToolTip(this.automaticCheck, "Check automatically once a day");
            this.tips.SetToolTip(this.automaticDownload, "Download updates automatically");
            this.tips.SetToolTip(this.automaticInstall, "Install automatically after closing VBA hosts");
            this.tips.SetToolTip(this.previews, "Include prereleases");
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(850, 630);
            this.MinimumSize = new System.Drawing.Size(780, 570);
            this.Controls.Add(this.layout);
            this.CancelButton = this.close;
            this.Name = "UpdateWindow";
            this.Text = "VBAi updates";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.layout.ResumeLayout(false);
            this.layout.PerformLayout();
            this.buttons.ResumeLayout(false);
            this.buttons.PerformLayout();
            this.close.Symbol = CodexVBE.UiSymbol.Close;
            this.close.IconOnly = true;
            this.close.AutoSize = false;
            this.close.MinimumSize = System.Drawing.Size.Empty;
            this.close.Size = new System.Drawing.Size(32, 30);
            this.check.Symbol = CodexVBE.UiSymbol.Refresh;
            this.check.IconOnly = true;
            this.check.AutoSize = false;
            this.check.MinimumSize = System.Drawing.Size.Empty;
            this.check.Size = new System.Drawing.Size(32, 30);
            this.download.Symbol = CodexVBE.UiSymbol.Download;
            this.download.IconOnly = true;
            this.download.AutoSize = false;
            this.download.MinimumSize = System.Drawing.Size.Empty;
            this.download.Size = new System.Drawing.Size(32, 30);
            this.install.Symbol = CodexVBE.UiSymbol.Download;
            this.install.Primary = true;
            this.skip.Symbol = CodexVBE.UiSymbol.Next;
            this.skip.IconOnly = true;
            this.skip.AutoSize = false;
            this.skip.MinimumSize = System.Drawing.Size.Empty;
            this.skip.Size = new System.Drawing.Size(32, 30);
            this.save.Symbol = CodexVBE.UiSymbol.Save;
            this.save.IconOnly = true;
            this.save.AutoSize = false;
            this.save.MinimumSize = System.Drawing.Size.Empty;
            this.save.Size = new System.Drawing.Size(32, 30);
            this.cancelPending.Symbol = CodexVBE.UiSymbol.Stop;
            this.ResumeLayout(false);
        }
    }
}

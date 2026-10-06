namespace VBAi
{

    /// <summary>Designer-generated controls and layout for the crash-report dialog.</summary>
    partial class CrashReportWindow
    {

        /// <summary>Container that owns nonvisual form components.</summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>Root layout for the report identity, inputs, preview, and actions.</summary>
        private System.Windows.Forms.TableLayoutPanel layout;

        /// <summary>Dialog heading.</summary>
        private System.Windows.Forms.Label heading;

        /// <summary>Displayed report identifier.</summary>
        private System.Windows.Forms.Label reportIdentity;

        /// <summary>Privacy notice describing the report contents.</summary>
        private System.Windows.Forms.Label privacy;

        /// <summary>Displayed report destination.</summary>
        private System.Windows.Forms.Label destination;

        /// <summary>Title input caption.</summary>
        private System.Windows.Forms.Label titleLabel;

        /// <summary>Editable issue title.</summary>
        private VBAi.UiTextBox titleInput;

        /// <summary>Description input caption.</summary>
        private System.Windows.Forms.Label descriptionLabel;

        /// <summary>Editable issue description.</summary>
        private VBAi.UiTextBox descriptionInput;

        /// <summary>Preview caption.</summary>
        private System.Windows.Forms.Label previewLabel;

        /// <summary>Formatted report preview.</summary>
        private VBAi.UiTextBox preview;

        /// <summary>Delivery and save feedback.</summary>
        private System.Windows.Forms.Label status;

        /// <summary>Link to the created GitHub issue.</summary>
        private System.Windows.Forms.LinkLabel issueLink;

        /// <summary>Progress indicator shown during delivery.</summary>
        private System.Windows.Forms.ProgressBar progress;

        /// <summary>Flow layout containing report actions.</summary>
        private System.Windows.Forms.FlowLayoutPanel buttons;

        /// <summary>Closes the dialog after delivery is no longer active.</summary>
        private VBAi.UiActionButton closeButton;

        /// <summary>Publishes the report through the configured delivery service.</summary>
        private VBAi.UiActionButton sendButton;

        /// <summary>Hands the report to Outlook or a local email draft.</summary>
        private VBAi.UiActionButton emailButton;

        /// <summary>Copies the formatted report body.</summary>
        private VBAi.UiActionButton copyButton;

        /// <summary>Saves a local Markdown report.</summary>
        private VBAi.UiActionButton saveButton;

        /// <summary>Tooltips associated with report controls.</summary>
        private System.Windows.Forms.ToolTip tips;

        /// <summary>Releases runtime subscriptions and Designer-owned components.</summary>
        /// <param name="disposing">Whether managed components should be disposed.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) { DisposeRuntime(); if (components != null) components.Dispose(); }
            base.Dispose(disposing);
        }

        /// <summary>Creates and arranges the crash-report dialog controls.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.heading = new System.Windows.Forms.Label();
            this.reportIdentity = new System.Windows.Forms.Label();
            this.privacy = new System.Windows.Forms.Label();
            this.destination = new System.Windows.Forms.Label();
            this.titleLabel = new System.Windows.Forms.Label();
            this.titleInput = new VBAi.UiTextBox();
            this.descriptionLabel = new System.Windows.Forms.Label();
            this.descriptionInput = new VBAi.UiTextBox();
            this.previewLabel = new System.Windows.Forms.Label();
            this.preview = new VBAi.UiTextBox();
            this.status = new System.Windows.Forms.Label();
            this.issueLink = new System.Windows.Forms.LinkLabel();
            this.progress = new System.Windows.Forms.ProgressBar();
            this.buttons = new System.Windows.Forms.FlowLayoutPanel();
            this.closeButton = new VBAi.UiActionButton();
            this.sendButton = new VBAi.UiActionButton();
            this.emailButton = new VBAi.UiActionButton();
            this.copyButton = new VBAi.UiActionButton();
            this.saveButton = new VBAi.UiActionButton();
            this.tips = new System.Windows.Forms.ToolTip(this.components);
            this.layout.SuspendLayout();
            this.buttons.SuspendLayout();
            this.SuspendLayout();
            this.layout.Name = "layout";
            this.heading.Name = "heading";
            this.reportIdentity.Name = "reportIdentity";
            this.privacy.Name = "privacy";
            this.titleLabel.Name = "titleLabel";
            this.titleInput.Name = "titleInput";
            this.descriptionLabel.Name = "descriptionLabel";
            this.descriptionInput.Name = "descriptionInput";
            this.previewLabel.Name = "previewLabel";
            this.preview.Name = "preview";
            this.status.Name = "status";
            this.issueLink.Name = "issueLink";
            this.progress.Name = "progress";
            this.buttons.Name = "buttons";
            this.closeButton.Name = "closeButton";
            this.sendButton.Name = "sendButton";
            this.emailButton.Name = "emailButton";
            this.copyButton.Name = "copyButton";
            this.saveButton.Name = "saveButton";
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.Padding = new System.Windows.Forms.Padding(16);
            this.layout.ColumnCount = 1;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.RowCount = 14;
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 88F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.heading, 0, 0);
            this.layout.Controls.Add(this.reportIdentity, 0, 1);
            this.layout.Controls.Add(this.privacy, 0, 2);
            this.layout.Controls.Add(this.destination, 0, 3);
            this.destination.Name = "destination";
            this.destination.Text = "GitHub · blackcancer/VBAi → Outlook · init-sys-rev@hotmail.com";
            this.destination.AutoSize = true;
            this.destination.UseCompatibleTextRendering = true;
            this.destination.Dock = System.Windows.Forms.DockStyle.Fill;
            this.destination.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.destination.Margin = new System.Windows.Forms.Padding(0, 0, 0, 12);
            this.layout.Controls.Add(this.titleLabel, 0, 4);
            this.layout.Controls.Add(this.titleInput, 0, 5);
            this.layout.Controls.Add(this.descriptionLabel, 0, 6);
            this.layout.Controls.Add(this.descriptionInput, 0, 7);
            this.layout.Controls.Add(this.previewLabel, 0, 8);
            this.layout.Controls.Add(this.preview, 0, 9);
            this.layout.Controls.Add(this.status, 0, 10);
            this.layout.Controls.Add(this.issueLink, 0, 11);
            this.layout.Controls.Add(this.progress, 0, 12);
            this.layout.Controls.Add(this.buttons, 0, 13);
            this.heading.Text = "Report an issue";
            this.privacy.Text = "Review before sending. No VBA, chat, credentials or file paths are collected automatically.";
            this.titleLabel.Text = "Title";
            this.descriptionLabel.Text = "What happened?";
            this.previewLabel.Text = "Report preview";
            this.issueLink.Text = "Open GitHub issue";
            this.heading.AutoSize = true;
            this.heading.Dock = System.Windows.Forms.DockStyle.Fill;
            this.heading.Margin = new System.Windows.Forms.Padding(0, 0, 0, 8);
            this.reportIdentity.AutoSize = true;
            this.reportIdentity.Dock = System.Windows.Forms.DockStyle.Fill;
            this.reportIdentity.Margin = new System.Windows.Forms.Padding(0, 0, 0, 8);
            this.privacy.AutoSize = true;
            this.privacy.Dock = System.Windows.Forms.DockStyle.Fill;
            this.privacy.Margin = new System.Windows.Forms.Padding(0, 0, 0, 8);
            this.titleLabel.AutoSize = true;
            this.titleLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.titleLabel.Margin = new System.Windows.Forms.Padding(0, 0, 0, 8);
            this.descriptionLabel.AutoSize = true;
            this.descriptionLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.descriptionLabel.Margin = new System.Windows.Forms.Padding(0, 0, 0, 8);
            this.previewLabel.AutoSize = true;
            this.previewLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.previewLabel.Margin = new System.Windows.Forms.Padding(0, 0, 0, 8);
            this.status.AutoSize = true;
            this.status.Dock = System.Windows.Forms.DockStyle.Fill;
            this.status.Margin = new System.Windows.Forms.Padding(0, 0, 0, 8);
            this.issueLink.AutoSize = true;
            this.issueLink.Dock = System.Windows.Forms.DockStyle.Fill;
            this.issueLink.Margin = new System.Windows.Forms.Padding(0, 0, 0, 8);
            this.heading.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold);
            this.privacy.MaximumSize = new System.Drawing.Size(1000, 0);
            this.status.AccessibleRole = System.Windows.Forms.AccessibleRole.StatusBar;
            this.status.MinimumSize = new System.Drawing.Size(0, 28);
            this.issueLink.Visible = false;
            this.issueLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.Issue_Click);
            this.titleInput.Dock = System.Windows.Forms.DockStyle.Fill;
            this.titleInput.MaxLength = 180;
            this.titleInput.TabIndex = 0;
            this.titleInput.TextChanged += new System.EventHandler(this.ContentChanged);
            this.descriptionInput.Dock = System.Windows.Forms.DockStyle.Fill;
            this.descriptionInput.MaxLength = 8000;
            this.descriptionInput.TabIndex = 1;
            this.descriptionInput.Multiline = true;
            this.descriptionInput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.descriptionInput.TextChanged += new System.EventHandler(this.ContentChanged);
            this.preview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.preview.MaxLength = 60000;
            this.preview.TabIndex = 2;
            this.preview.Multiline = true;
            this.preview.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.preview.ReadOnly = true;
            this.preview.Font = new System.Drawing.Font("Consolas", 9F);
            this.preview.MinimumSize = new System.Drawing.Size(0, 140);
            this.preview.TabStop = true;
            this.progress.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progress.Height = 4;
            this.progress.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progress.Visible = false;
            this.buttons.Dock = System.Windows.Forms.DockStyle.Fill;
            this.buttons.AutoSize = true;
            this.buttons.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.buttons.WrapContents = true;
            this.buttons.Controls.Add(this.closeButton);
            this.closeButton.Text = "Close";
            this.closeButton.AutoSize = true;
            this.closeButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.closeButton.MinimumSize = new System.Drawing.Size(80, 36);
            this.closeButton.Padding = new System.Windows.Forms.Padding(10, 4, 10, 4);
            this.closeButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.closeButton.TabIndex = 3;
            this.tips.SetToolTip(this.closeButton, "Close");
            this.buttons.Controls.Add(this.sendButton);
            this.sendButton.Text = "Send report";
            this.sendButton.AutoSize = true;
            this.sendButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.sendButton.MinimumSize = new System.Drawing.Size(80, 36);
            this.sendButton.Padding = new System.Windows.Forms.Padding(10, 4, 10, 4);
            this.sendButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.sendButton.TabIndex = 4;
            this.tips.SetToolTip(this.sendButton, "GitHub first; Outlook if unavailable.");
            this.sendButton.Click += new System.EventHandler(this.Send_Click);
            this.buttons.Controls.Add(this.emailButton);
            this.emailButton.Text = "Email";
            this.emailButton.AutoSize = true;
            this.emailButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.emailButton.MinimumSize = new System.Drawing.Size(80, 36);
            this.emailButton.Padding = new System.Windows.Forms.Padding(10, 4, 10, 4);
            this.emailButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.emailButton.TabIndex = 5;
            this.tips.SetToolTip(this.emailButton, "Send with Outlook, or open an email draft.");
            this.emailButton.Click += new System.EventHandler(this.Email_Click);
            this.buttons.Controls.Add(this.saveButton);
            this.saveButton.Text = "Save report";
            this.saveButton.AutoSize = true;
            this.saveButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.saveButton.MinimumSize = new System.Drawing.Size(80, 36);
            this.saveButton.Padding = new System.Windows.Forms.Padding(10, 4, 10, 4);
            this.saveButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.saveButton.TabIndex = 6;
            this.tips.SetToolTip(this.saveButton, "Save the full report and copy its path.");
            this.saveButton.Click += new System.EventHandler(this.Save_Click);
            this.buttons.Controls.Add(this.copyButton);
            this.copyButton.Text = "Copy";
            this.copyButton.AutoSize = true;
            this.copyButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.copyButton.MinimumSize = new System.Drawing.Size(80, 36);
            this.copyButton.Padding = new System.Windows.Forms.Padding(10, 4, 10, 4);
            this.copyButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.copyButton.TabIndex = 7;
            this.tips.SetToolTip(this.copyButton, "Copy technical details");
            this.copyButton.Click += new System.EventHandler(this.Copy_Click);
            this.closeButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.tips.SetToolTip(this.titleInput, "Title");
            this.tips.SetToolTip(this.descriptionInput, "Do not include confidential information in your description.");
            this.tips.SetToolTip(this.preview, "Report preview");
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.ClientSize = new System.Drawing.Size(780, 730);
            this.MinimumSize = new System.Drawing.Size(660, 660);
            this.Controls.Add(this.layout);
            this.CancelButton = this.closeButton;
            this.Text = "VBAi · Report an issue";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.WindowClosing);
            this.layout.ResumeLayout(false);
            this.layout.PerformLayout();
            this.buttons.ResumeLayout(false);
            this.buttons.PerformLayout();
            this.closeButton.Symbol = VBAi.UiSymbol.Close;
            this.closeButton.IconOnly = true;
            this.closeButton.AutoSize = false;
            this.closeButton.MinimumSize = System.Drawing.Size.Empty;
            this.closeButton.Size = new System.Drawing.Size(32, 30);
            this.sendButton.Symbol = VBAi.UiSymbol.Upload;
            this.sendButton.Primary = true;
            this.emailButton.Symbol = VBAi.UiSymbol.Mail;
            this.emailButton.IconOnly = true;
            this.emailButton.AutoSize = false;
            this.emailButton.MinimumSize = System.Drawing.Size.Empty;
            this.emailButton.Size = new System.Drawing.Size(32, 30);
            this.copyButton.Symbol = VBAi.UiSymbol.Copy;
            this.copyButton.IconOnly = true;
            this.copyButton.AutoSize = false;
            this.copyButton.MinimumSize = System.Drawing.Size.Empty;
            this.copyButton.Size = new System.Drawing.Size(32, 30);
            this.saveButton.Symbol = VBAi.UiSymbol.Save;
            this.saveButton.Primary = true;
            this.ResumeLayout(false);
        }
    }
}

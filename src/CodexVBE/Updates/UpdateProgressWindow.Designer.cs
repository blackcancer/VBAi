namespace CodexVBE
{
    /// <summary>Designer-generated controls and layout for updater progress.</summary>
partial class UpdateProgressWindow
    {
        /// <summary>Container that owns the progress timer.</summary>
private System.ComponentModel.IContainer components;
        /// <summary>Root table layout for progress text and actions.</summary>
private System.Windows.Forms.TableLayoutPanel layout;
        /// <summary>Window heading.</summary>
private System.Windows.Forms.Label heading;
        /// <summary>Window heading.</summary>
private System.Windows.Forms.Label version;
        /// <summary>Window heading.</summary>
private System.Windows.Forms.Label status;
        /// <summary>Installer progress indicator.</summary>
private System.Windows.Forms.ProgressBar progress;
        /// <summary>Close or cancel action.</summary>
private System.Windows.Forms.Button cancel;
        /// <summary>Timer that polls the updater job.</summary>
private System.Windows.Forms.Timer timer;
        /// <summary>Releases Designer-owned components.</summary><param name="disposing">Whether managed components should be disposed.</param>
protected override void Dispose(bool disposing) { if (disposing && components != null) components.Dispose(); base.Dispose(disposing); }
        /// <summary>Creates and arranges progress controls and the polling timer.</summary>
private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.heading = new System.Windows.Forms.Label();
            this.version = new System.Windows.Forms.Label();
            this.status = new System.Windows.Forms.Label();
            this.progress = new System.Windows.Forms.ProgressBar();
            this.cancel = new System.Windows.Forms.Button();
            this.timer = new System.Windows.Forms.Timer(this.components);
            this.layout.SuspendLayout();
            this.SuspendLayout();
            this.layout.Name = "layout";
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.Padding = new System.Windows.Forms.Padding(24);
            this.layout.ColumnCount = 1;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.RowCount = 5;
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.heading, 0, 0);
            this.layout.Controls.Add(this.version, 0, 1);
            this.layout.Controls.Add(this.status, 0, 2);
            this.layout.Controls.Add(this.progress, 0, 3);
            this.layout.Controls.Add(this.cancel, 0, 4);
            this.heading.Name = "heading";
            this.heading.Text = "VBAi updates";
            this.heading.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold);
            this.heading.Dock = System.Windows.Forms.DockStyle.Fill;
            this.heading.AutoSize = true;
            this.version.Name = "version";
            this.version.AutoSize = true;
            this.version.Dock = System.Windows.Forms.DockStyle.Fill;
            this.status.Name = "status";
            this.status.Text = "Waiting for VBA hosts to close.";
            this.status.Dock = System.Windows.Forms.DockStyle.Fill;
            this.status.AutoSize = true;
            this.status.Margin = new System.Windows.Forms.Padding(3, 18, 3, 18);
            this.progress.Name = "progress";
            this.progress.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progress.Margin = System.Windows.Forms.Padding.Empty;
            this.progress.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.cancel.Name = "cancel";
            this.cancel.Text = "Cancel";
            this.cancel.AutoSize = true;
            this.cancel.MinimumSize = new System.Drawing.Size(100, 36);
            this.cancel.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.cancel.Margin = new System.Windows.Forms.Padding(0, 18, 0, 0);
            this.cancel.Click += new System.EventHandler(this.Close_Click);
            this.timer.Interval = 2000;
            this.timer.Tick += new System.EventHandler(this.Poll);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(580, 280);
            this.MinimumSize = new System.Drawing.Size(500, 270);
            this.Controls.Add(this.layout);
            this.Name = "UpdateProgressWindow";
            this.Text = "VBAi updates";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.WindowClosing);
            this.layout.ResumeLayout(false);
            this.layout.PerformLayout();
            this.ResumeLayout(false);
        }
    }
}

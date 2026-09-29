namespace VBAi
{
    /// <summary>Designer-generated review comment list and selected comment text.</summary>
    public sealed partial class GitHubPullCommentsView
    {
        /// <summary>Layout for the comment list and body.</summary>
        internal System.Windows.Forms.TableLayoutPanel commentLayout;
        /// <summary>Review comments on the selected pull request.</summary>
        internal VBAi.UiListBox comments;
        /// <summary>Text of the selected review comment.</summary>
        internal VBAi.UiTextBox commentBody;
        /// <summary>Container that owns Designer components.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Tooltips associated with comments.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Creates and arranges pull-request comment controls.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.commentLayout = new System.Windows.Forms.TableLayoutPanel();
            this.comments = new VBAi.UiListBox();
            this.commentBody = new VBAi.UiTextBox();
            this.commentLayout.SuspendLayout();
            this.SuspendLayout();
            this.commentLayout.Name = "commentLayout";
            this.comments.Name = "comments";
            this.commentBody.Name = "commentBody";
            this.commentBody.Multiline = true;
            this.commentBody.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.commentBody.Dock = System.Windows.Forms.DockStyle.Fill;
            this.commentBody.ReadOnly = true;
            this.commentLayout.ColumnCount = 1;
            this.commentLayout.RowCount = 2;
            this.commentLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.commentLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 45F));
            this.commentLayout.Controls.Add(this.comments, 0, 0);
            this.comments.Dock = System.Windows.Forms.DockStyle.Fill;
            this.commentLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 55F));
            this.commentLayout.Controls.Add(this.commentBody, 0, 1);
            this.commentBody.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comments.HorizontalScrollbar = true;
            this.Controls.Add(this.commentLayout);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitHubPullCommentsView";
            this.Size = new System.Drawing.Size(860, 500);
            this.commentLayout.Location = new System.Drawing.Point(0, 0);
            this.commentLayout.Size = new System.Drawing.Size(860, 500);
            this.commentLayout.TabIndex = 0;
            this.comments.Location = new System.Drawing.Point(3, 3);
            this.comments.Size = new System.Drawing.Size(854, 219);
            this.comments.TabIndex = 0;
            this.commentBody.Location = new System.Drawing.Point(3, 228);
            this.commentBody.Size = new System.Drawing.Size(854, 269);
            this.commentBody.TabIndex = 1;
            this.commentLayout.ResumeLayout(false);
            this.commentLayout.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

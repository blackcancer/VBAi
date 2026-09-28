namespace CodexVBE
{
    public sealed partial class GitHubPullCommentsView
    {
        internal System.Windows.Forms.TableLayoutPanel commentLayout;
        internal System.Windows.Forms.ListBox comments;
        internal System.Windows.Forms.TextBox commentBody;
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
            this.commentLayout = new System.Windows.Forms.TableLayoutPanel();
            this.comments = new System.Windows.Forms.ListBox();
            this.commentBody = new System.Windows.Forms.TextBox();
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
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

namespace CodexVBE
{
    public sealed partial class ChatWelcomeView
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.ToolTip toolTips;
        private ChatComposerPanel layout;
        internal System.Windows.Forms.Label title;
        internal System.Windows.Forms.Label hint;
        internal ChatActionButton explain;
        internal ChatActionButton fix;
        internal ChatActionButton improve;
        /// <summary>Releases the Designer components.</summary>
        /// <param name="disposing">Whether to release managed resources.</param>
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.layout = new ChatComposerPanel();
            this.title = new System.Windows.Forms.Label();
            this.title.Name = "title";
            this.title.AutoSize = true;
            this.title.Dock = System.Windows.Forms.DockStyle.Fill;
            this.title.Text = "What would you like to build?";
            this.title.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold);
            this.hint = new System.Windows.Forms.Label();
            this.hint.Name = "hint";
            this.hint.AutoSize = true;
            this.hint.Dock = System.Windows.Forms.DockStyle.Fill;
            this.hint.Text = "Add #aModule or @aFunction to work on your code.";
            this.explain = new ChatActionButton();
            this.explain.Name = "explain";
            this.explain.AutoSize = true;
            this.explain.Text = "Explain a procedure";
            this.fix = new ChatActionButton();
            this.fix.Name = "fix";
            this.fix.AutoSize = true;
            this.fix.Text = "Fix an error";
            this.improve = new ChatActionButton();
            this.improve.Name = "improve";
            this.improve.AutoSize = true;
            this.improve.Text = "Improve the code";
            this.SuspendLayout(); this.layout.SuspendLayout();
            this.layout.AutoSize = true;
            this.layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.layout.Dock = System.Windows.Forms.DockStyle.Top;
            this.layout.ColumnCount = 1;
            this.layout.RowCount = 5;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.Controls.Add(this.title, 0, 0);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.hint, 0, 1);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.explain, 0, 2);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.fix, 0, 3);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.improve, 0, 4);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Name = "layout";
            this.layout.Padding = new System.Windows.Forms.Padding(10);
            this.toolTips.SetToolTip(this.fix, "Prepare a fix");
            this.Controls.Add(this.layout);
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "ChatWelcomeView";
            this.Size = new System.Drawing.Size(500, 240);
            this.layout.ResumeLayout(false); this.layout.PerformLayout(); this.ResumeLayout(false); this.PerformLayout();
        }
    }
}

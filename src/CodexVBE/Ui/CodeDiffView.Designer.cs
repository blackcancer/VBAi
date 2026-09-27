namespace CodexVBE
{
    public sealed partial class CodeDiffView
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.FlowLayoutPanel toolbar;
        private System.Windows.Forms.CheckBox unified;
        private System.Windows.Forms.CheckBox collapse;
        private System.Windows.Forms.Button previous;
        private System.Windows.Forms.Button next;
        private System.Windows.Forms.TextBox search;
        private System.Windows.Forms.Button find;
        private System.Windows.Forms.DataGridView grid;
        private System.Windows.Forms.ToolTip tips;
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolbar = new System.Windows.Forms.FlowLayoutPanel();
            this.unified = new System.Windows.Forms.CheckBox();
            this.collapse = new System.Windows.Forms.CheckBox();
            this.previous = new CodexVBE.ThemedButton();
            this.next = new CodexVBE.ThemedButton();
            this.search = new System.Windows.Forms.TextBox();
            this.find = new CodexVBE.ThemedButton();
            this.grid = new System.Windows.Forms.DataGridView();
            this.tips = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            this.toolbar.Dock = System.Windows.Forms.DockStyle.Top;
            this.toolbar.AutoSize = true;
            this.toolbar.Controls.AddRange(new System.Windows.Forms.Control[] { this.unified, this.collapse, this.previous, this.next, this.search, this.find });
            this.unified.Text = "Unified";
            this.unified.AutoSize = true;
            this.unified.CheckedChanged += new System.EventHandler(this.OptionsChanged);
            this.collapse.Text = "Fold context";
            this.collapse.AutoSize = true;
            this.collapse.Checked = true;
            this.collapse.CheckedChanged += new System.EventHandler(this.OptionsChanged);
            this.previous.Text = "↑";
            this.previous.Width = 32;
            this.previous.Click += new System.EventHandler(this.Previous_Click);
            this.next.Text = "↓";
            this.next.Width = 32;
            this.next.Click += new System.EventHandler(this.Next_Click);
            this.search.Width = 120;
            this.search.AccessibleName = "Search code";
            this.search.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Search_KeyDown);
            this.find.Text = "Find";
            this.find.AutoSize = true;
            this.find.Click += new System.EventHandler(this.Find_Click);
            this.tips.SetToolTip(this.previous, "Previous change");
            this.tips.SetToolTip(this.next, "Next change");
            this.tips.SetToolTip(this.search, "Search code");
            this.tips.SetToolTip(this.collapse, "Show three context lines around each change.");
            this.grid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grid.VirtualMode = true;
            this.grid.ReadOnly = true;
            this.grid.AllowUserToAddRows = false;
            this.grid.AllowUserToDeleteRows = false;
            this.grid.RowHeadersVisible = false;
            this.grid.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.None;
            this.grid.Font = new System.Drawing.Font("Consolas", 10F);
            this.grid.CellValueNeeded += new System.Windows.Forms.DataGridViewCellValueEventHandler(this.ValueNeeded);
            this.grid.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.FormatCell);
            this.grid.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.PaintCell);
            this.grid.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.ExpandContext);
            this.Controls.Add(this.grid);
            this.Controls.Add(this.toolbar);
            this.Name = "CodeDiffView";
            this.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Size = new System.Drawing.Size(700, 350);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

namespace CodexVBE
{
    /// <summary>Déclare les contrôles WinForms générés pour la barre d’outils et la grille de différences.</summary>
    public sealed partial class CodeDiffView
    {
        /// <summary>Conteneur des composants WinForms du concepteur.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Barre qui regroupe les options, navigation et recherche du diff.</summary>
        private System.Windows.Forms.FlowLayoutPanel toolbar;
        /// <summary>Option qui choisit l’affichage unifié des changements.</summary>
        private System.Windows.Forms.CheckBox unified;
        /// <summary>Option qui replie le contexte éloigné des changements.</summary>
        private System.Windows.Forms.CheckBox collapse;
        /// <summary>Bouton de navigation vers le changement précédent.</summary>
        private CodexVBE.ThemedButton previous;
        /// <summary>Bouton de navigation vers le changement suivant.</summary>
        private CodexVBE.ThemedButton next;
        /// <summary>Champ de recherche dans le contenu du diff.</summary>
        private System.Windows.Forms.TextBox search;
        /// <summary>Bouton qui lance la recherche courante.</summary>
        private CodexVBE.ThemedButton find;
        /// <summary>Grille virtuelle en lecture seule qui affiche les lignes du diff.</summary>
        private System.Windows.Forms.DataGridView grid;
        /// <summary>Composant qui affiche les explications des commandes de la barre d’outils.</summary>
        private System.Windows.Forms.ToolTip tips;
        /// <summary>Colonnes des positions et du texte avant et après le changement.</summary>
        private System.Windows.Forms.DataGridViewTextBoxColumn oldLineColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn beforeColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn newLineColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn afterColumn;
        /// <summary>Crée et configure la barre d’outils, la grille et leurs gestionnaires d’événements.</summary>
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
            ((System.ComponentModel.ISupportInitialize)(this.grid)).BeginInit();
            this.toolbar.SuspendLayout();
            this.grid.SuspendLayout();
            this.SuspendLayout();
            this.tips = new System.Windows.Forms.ToolTip(this.components);
            this.oldLineColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.beforeColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.newLineColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.afterColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();

            this.oldLineColumn.Name = "oldLineColumn";
            this.oldLineColumn.HeaderText = "−";
            this.oldLineColumn.Width = 52;
            this.oldLineColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.beforeColumn.Name = "beforeColumn";
            this.beforeColumn.HeaderText = "Before";
            this.beforeColumn.MinimumWidth = 120;
            this.beforeColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.beforeColumn.FillWeight = 100F;
            this.beforeColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.newLineColumn.Name = "newLineColumn";
            this.newLineColumn.HeaderText = "+";
            this.newLineColumn.Width = 52;
            this.newLineColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.afterColumn.Name = "afterColumn";
            this.afterColumn.HeaderText = "After";
            this.afterColumn.MinimumWidth = 120;
            this.afterColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.afterColumn.FillWeight = 100F;
            this.afterColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.grid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { this.oldLineColumn, this.beforeColumn, this.newLineColumn, this.afterColumn });
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
            this.toolbar.Name = "toolbar";
            this.toolbar.Location = new System.Drawing.Point(0, 0);
            this.toolbar.Size = new System.Drawing.Size(700, 30);
            this.toolbar.TabIndex = 1;
            this.unified.Name = "unified";
            this.unified.Location = new System.Drawing.Point(3, 3);
            this.unified.Size = new System.Drawing.Size(59, 18);
            this.unified.TabIndex = 0;
            this.collapse.Name = "collapse";
            this.collapse.Location = new System.Drawing.Point(68, 3);
            this.collapse.Size = new System.Drawing.Size(85, 18);
            this.collapse.TabIndex = 1;
            this.previous.Name = "previous";
            this.previous.Location = new System.Drawing.Point(159, 3);
            this.previous.Size = new System.Drawing.Size(32, 23);
            this.previous.TabIndex = 2;
            this.next.Name = "next";
            this.next.Location = new System.Drawing.Point(197, 3);
            this.next.Size = new System.Drawing.Size(32, 23);
            this.next.TabIndex = 3;
            this.search.Name = "search";
            this.search.Location = new System.Drawing.Point(235, 3);
            this.search.Size = new System.Drawing.Size(120, 20);
            this.search.TabIndex = 4;
            this.find.Name = "find";
            this.find.Location = new System.Drawing.Point(361, 3);
            this.find.Size = new System.Drawing.Size(75, 24);
            this.find.TabIndex = 5;
            this.grid.Name = "grid";
            this.grid.Location = new System.Drawing.Point(0, 30);
            this.grid.Size = new System.Drawing.Size(700, 320);
            this.grid.TabIndex = 0;
            ((System.ComponentModel.ISupportInitialize)(this.grid)).EndInit();
            this.toolbar.ResumeLayout(false);
            this.toolbar.PerformLayout();
            this.grid.ResumeLayout(false);
            this.grid.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

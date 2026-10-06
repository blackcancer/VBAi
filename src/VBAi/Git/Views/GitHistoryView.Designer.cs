namespace VBAi
{

    /// <summary>Designer-generated commit history list and details controls.</summary>
    public sealed partial class GitHistoryView
    {

        /// <summary>Commit history entries.</summary>
        internal VBAi.UiListBox history;

        /// <summary>Details for the selected commit.</summary>
        internal VBAi.UiTextBox historyDetails;

        /// <summary>Compares the selected commit with the current project.</summary>
        internal VBAi.ThemedButton historyCompare;

        /// <summary>Container that owns Designer components.</summary>
        private System.ComponentModel.IContainer components;

        /// <summary>Tooltips associated with history actions.</summary>
        private System.Windows.Forms.ToolTip toolTips;

        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Creates and arranges commit history controls.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.history = new VBAi.UiListBox();
            this.historyDetails = new VBAi.UiTextBox();
            this.historyCompare = new VBAi.ThemedButton();
            this.SuspendLayout();
            this.history.Dock = System.Windows.Forms.DockStyle.Fill;
            this.history.HorizontalScrollbar = true;
            this.history.Name = "history";
            this.historyDetails.Name = "historyDetails";
            this.history.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.historyDetails.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.historyDetails.Height = 100;
            this.historyDetails.Multiline = true;
            this.historyDetails.ReadOnly = true;
            this.historyDetails.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.toolTips.SetToolTip(this.history, "Select one commit for details, or two commits to compare their VBA sources.");
            this.historyCompare.Name = "historyCompare";
            this.historyCompare.Text = "Compare revisions";
            this.historyCompare.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.historyCompare.AutoSize = true;
            this.Controls.Add(this.history);
            this.Controls.Add(this.historyDetails);
            this.Controls.Add(this.historyCompare);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitHistoryView";
            this.Size = new System.Drawing.Size(860, 500);
            this.history.Location = new System.Drawing.Point(0, 0);
            this.history.Size = new System.Drawing.Size(860, 372);
            this.history.TabIndex = 0;
            this.historyDetails.Location = new System.Drawing.Point(0, 372);
            this.historyDetails.Size = new System.Drawing.Size(860, 100);
            this.historyDetails.TabIndex = 1;
            this.historyCompare.Location = new System.Drawing.Point(0, 472);
            this.historyCompare.Size = new System.Drawing.Size(860, 28);
            this.historyCompare.TabIndex = 2;
            this.historyCompare.Symbol = VBAi.UiSymbol.Inspect;
            this.historyCompare.IconOnly = true;
            this.historyCompare.AutoSize = false;
            this.historyCompare.MinimumSize = System.Drawing.Size.Empty;
            this.historyCompare.Size = new System.Drawing.Size(32, 30);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

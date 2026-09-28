namespace CodexVBE
{
    /// <summary>Déclare les contrôles WinForms qui structurent la fenêtre Monaco.</summary>
    partial class ModernEditorWindow
    {
        /// <summary>Conteneur des composants non visuels créés par le concepteur.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Grille principale qui place la barre d’outils, les onglets, la surface Web et le statut.</summary>
        private System.Windows.Forms.TableLayoutPanel layout;
        /// <summary>Barre de commandes de l’éditeur, avec défilement horizontal.</summary>
        private System.Windows.Forms.FlowLayoutPanel toolbar;
        /// <summary>Commande qui applique la version éditée après comparaison.</summary>
        private CodexVBE.ThemedButton resolve;
        /// <summary>Commande qui compare le brouillon au code VBA natif.</summary>
        private CodexVBE.ThemedButton compare;
        /// <summary>Commande qui revient à l’édition après la comparaison.</summary>
        private CodexVBE.ThemedButton edit;
        /// <summary>Commande qui recharge la version native en préservant le brouillon.</summary>
        private CodexVBE.ThemedButton reload;
        /// <summary>Commande qui restaure un brouillon récupéré.</summary>
        private CodexVBE.ThemedButton restore;
        /// <summary>Onglets des modules ouverts.</summary>
        private CodexVBE.ThemedTabControl tabs;
        /// <summary>Surface qui héberge le contrôle WebView2.</summary>
        private System.Windows.Forms.Panel surface;
        /// <summary>Message d’état et de synchronisation du document actif.</summary>
        private System.Windows.Forms.Label status;
        /// <summary>Minuterie de capture des révisions et de synchronisation périodique.</summary>
        private System.Windows.Forms.Timer timer;
        /// <summary>Info-bulles descriptives des commandes de la barre d’outils.</summary>
        private System.Windows.Forms.ToolTip tips;
        /// <summary>Libère les ressources de WebView2, du worker et des fenêtres CodePane détenues.</summary>
        /// <param name="disposing">Indique si la libération concerne aussi les ressources managées.</param>
        protected override void Dispose(bool disposing) { if (disposing) { DisposeRuntime(); components?.Dispose(); } base.Dispose(disposing); }
        /// <summary>Crée et dispose les contrôles de la fenêtre d’édition.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.toolbar = new System.Windows.Forms.FlowLayoutPanel();
            this.resolve = new CodexVBE.ThemedButton();
            this.compare = new CodexVBE.ThemedButton();
            this.edit = new CodexVBE.ThemedButton();
            this.reload = new CodexVBE.ThemedButton();
            this.restore = new CodexVBE.ThemedButton();
            this.tabs = new CodexVBE.ThemedTabControl();
            this.surface = new System.Windows.Forms.Panel();
            this.status = new System.Windows.Forms.Label();
            this.timer = new System.Windows.Forms.Timer(this.components);
            this.tips = new System.Windows.Forms.ToolTip(this.components);
            this.layout.SuspendLayout();
            this.toolbar.SuspendLayout();
            this.SuspendLayout();
            this.layout.ColumnCount = 1;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.RowCount = 4;
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 44F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.Controls.Add(this.toolbar, 0, 0);
            this.layout.Controls.Add(this.tabs, 0, 1);
            this.layout.Controls.Add(this.surface, 0, 2);
            this.layout.Controls.Add(this.status, 0, 3);
            this.layout.Name = "layout";
            this.toolbar.AutoSize = false;
            this.toolbar.WrapContents = false;
            this.toolbar.AutoScroll = true;
            this.toolbar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.toolbar.Padding = new System.Windows.Forms.Padding(5);
            this.toolbar.Controls.Add(this.compare);
            this.toolbar.Controls.Add(this.edit);
            this.toolbar.Controls.Add(this.reload);
            this.toolbar.Controls.Add(this.restore);
            this.toolbar.Controls.Add(this.resolve);
            this.toolbar.Name = "toolbar";
            this.compare.AutoSize = true;
            this.compare.Visible = false;
            this.compare.Text = "Compare with VBA";
            this.compare.Name = "compare";
            this.compare.Click += new System.EventHandler(this.DiffClick);
            this.edit.AutoSize = true;
            this.edit.Visible = false;
            this.edit.Text = "Code";
            this.edit.Name = "edit";
            this.edit.Click += new System.EventHandler(this.EditClick);
            this.reload.AutoSize = true;
            this.reload.Visible = false;
            this.reload.Text = "Reload VBA version";
            this.reload.Name = "reload";
            this.reload.Click += new System.EventHandler(this.ReloadClick);
            this.restore.AutoSize = true;
            this.restore.Visible = false;
            this.restore.Text = "Restore draft";
            this.restore.Name = "restore";
            this.restore.Enabled = false;
            this.restore.Click += new System.EventHandler(this.RestoreClick);
            this.resolve.AutoSize = true;
            this.resolve.Visible = false;
            this.resolve.Name = "resolve";
            this.resolve.Text = "Use edited version";
            this.resolve.Click += new System.EventHandler(this.ResolveClick);
            this.tabs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabs.Name = "tabs";
            this.tabs.ShowCloseButtons = true;
            this.tabs.Padding = new System.Drawing.Point(18, 3);
            this.tabs.CloseRequested += new System.EventHandler<System.Windows.Forms.TabControlEventArgs>(this.CloseTabRequested);
            this.tabs.Margin = new System.Windows.Forms.Padding(0);
            this.tabs.SelectedIndexChanged += new System.EventHandler(this.TabChanged);
            this.surface.Dock = System.Windows.Forms.DockStyle.Fill;
            this.surface.Name = "surface";
            this.surface.Margin = new System.Windows.Forms.Padding(0);
            this.status.AutoSize = true;
            this.status.Dock = System.Windows.Forms.DockStyle.Fill;
            this.status.Padding = new System.Windows.Forms.Padding(9);
            this.status.Name = "status";
            this.timer.Interval = 900;
            this.timer.Tick += new System.EventHandler(this.TimerTick);
            this.tips.SetToolTip(this.tabs, "Close this tab and preserve unsynchronized changes as a recovery draft.");
            this.tips.SetToolTip(this.resolve, "Apply your edited version after comparing; refuse if VBA changed again.");
            this.tips.SetToolTip(this.reload, "Keep a recovery draft, then load the current VBA source.");
            this.tips.SetToolTip(this.restore, "Restore a previous draft without overwriting changed VBA code.");
            this.tips.SetToolTip(this.edit, "Return from the comparison to editing.");
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1100, 740);
            this.MinimumSize = new System.Drawing.Size(720, 460);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Controls.Add(this.layout);
            this.ControlBox = false;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.ShowInTaskbar = false;
            this.Text = "VBAi editor";
            this.Name = "ModernEditorWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ClosingWindow);
            this.toolbar.ResumeLayout(false);
            this.toolbar.PerformLayout();
            this.layout.ResumeLayout(false);
            this.layout.PerformLayout();
            this.ResumeLayout(false);
        }
    }
}

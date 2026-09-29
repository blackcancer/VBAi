namespace VBAi
{
    /// <summary>Contrôles générés de la vignette de contexte avec actions d’ouverture et de retrait.</summary>
    public sealed partial class ChatContextChipView
    {
        /// <summary>Container that owns the disposable components created by the WinForms Designer.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Flow layout panel that contains this transcript view&apos;s child controls.</summary>
        private System.Windows.Forms.FlowLayoutPanel layout;
        /// <summary>Bouton portant l’action principale sur l’élément.</summary>
        private VBAi.ChatActionButton open;
        /// <summary>Bouton de retrait de l’élément du contexte.</summary>
        private VBAi.ChatActionButton remove;
        /// <summary>ToolTip component used to show full text for transcript controls.</summary>
        private System.Windows.Forms.ToolTip toolTips;
                /// <summary>Libère les composants du modèle Designer.</summary>
                /// <param name="disposing">Indique si les composants gérés doivent également être libérés.</param>
        protected override void Dispose(bool disposing) { if (disposing && components != null) components.Dispose(); base.Dispose(disposing); }
        /// <summary>Creates and configures the chat context chip view controls serialized by the WinForms Designer.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.layout = new System.Windows.Forms.FlowLayoutPanel();
            this.open = new VBAi.ChatActionButton();
            this.remove = new VBAi.ChatActionButton();
            this.layout.SuspendLayout();
            this.SuspendLayout();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            // open
            this.open.AutoSize = true;
            this.open.Cursor = System.Windows.Forms.Cursors.Hand;
            this.open.Location = new System.Drawing.Point(2, 2);
            this.open.Margin = new System.Windows.Forms.Padding(2);
            this.open.Name = "open";
            this.open.Padding = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.open.Size = new System.Drawing.Size(112, 26);
            this.open.TabIndex = 0;
            this.open.Text = "#Reference";
            this.open.Click += new System.EventHandler(this.Open_Click);
            // remove
            this.remove.AutoSize = true;
            this.remove.Cursor = System.Windows.Forms.Cursors.Hand;
            this.remove.Location = new System.Drawing.Point(118, 2);
            this.remove.Margin = new System.Windows.Forms.Padding(2);
            this.remove.Name = "remove";
            this.remove.Padding = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.remove.Size = new System.Drawing.Size(28, 26);
            this.remove.TabIndex = 1;
            this.remove.Text = "×";
            this.remove.Click += new System.EventHandler(this.Remove_Click);
            // layout
            this.layout.AutoSize = true;
            this.layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.layout.Controls.Add(this.open);
            this.layout.Controls.Add(this.remove);
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.Location = new System.Drawing.Point(0, 0);
            this.layout.Margin = new System.Windows.Forms.Padding(0);
            this.layout.Name = "layout";
            this.layout.Size = new System.Drawing.Size(150, 30);
            this.layout.TabIndex = 0;
            this.layout.WrapContents = false;
            // ChatContextChipView
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.layout);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "ChatContextChipView";
            this.Size = new System.Drawing.Size(150, 30);
            this.layout.ResumeLayout(false);
            this.layout.PerformLayout();
            this.open.Symbol = VBAi.UiSymbol.Code;
            this.remove.Symbol = VBAi.UiSymbol.Close;
            this.remove.IconOnly = true;
            this.remove.AutoSize = false;
            this.remove.MinimumSize = System.Drawing.Size.Empty;
            this.remove.Size = new System.Drawing.Size(32, 30);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

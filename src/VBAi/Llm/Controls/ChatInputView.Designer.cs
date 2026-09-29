namespace VBAi
{
    /// <summary>Contrôles générés pour l’hôte WPF et l’aperçu WinForms de la saisie.</summary>
    public sealed partial class ChatInputView
    {
        /// <summary>Container that owns the disposable components created by the WinForms Designer.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Hôte WinForms destiné au moteur de saisie WPF.</summary>
        private VBAi.ChatContentHost host;
        /// <summary>Panneau d’aperçu visible dans le concepteur Visual Studio.</summary>
        private System.Windows.Forms.Panel previewPanel;
        /// <summary>Représentation WinForms statique destinée au concepteur.</summary>
        private VBAi.UiTextBox previewEditor;
                /// <summary>Libère les composants du modèle Designer.</summary>
                /// <param name="disposing">Indique si les composants gérés doivent également être libérés.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) { if (components != null) components.Dispose(); }
            base.Dispose(disposing);
        }
        /// <summary>Creates and configures the chat input view controls serialized by the WinForms Designer.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.host = new VBAi.ChatContentHost();
            this.previewPanel = new System.Windows.Forms.Panel();
            this.previewEditor = new VBAi.UiTextBox();
            this.previewPanel.SuspendLayout();
            this.SuspendLayout();
            // host
            this.host.Dock = System.Windows.Forms.DockStyle.Fill;
            this.host.Location = new System.Drawing.Point(0, 0);
            this.host.Margin = new System.Windows.Forms.Padding(0);
            this.host.Name = "host";
            this.host.Size = new System.Drawing.Size(552, 88);
            this.host.TabIndex = 0;
            this.host.Visible = false;
            // previewEditor: fixed WinForms representation for the Visual Studio Designer
            this.previewEditor.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.previewEditor.Dock = System.Windows.Forms.DockStyle.Fill;
            this.previewEditor.Location = new System.Drawing.Point(12, 12);
            this.previewEditor.Multiline = true;
            this.previewEditor.Name = "previewEditor";
            this.previewEditor.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.previewEditor.Size = new System.Drawing.Size(528, 64);
            this.previewEditor.TabIndex = 0;
            // previewPanel
            this.previewPanel.Controls.Add(this.previewEditor);
            this.previewPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.previewPanel.Location = new System.Drawing.Point(0, 0);
            this.previewPanel.Name = "previewPanel";
            this.previewPanel.Padding = new System.Windows.Forms.Padding(12);
            this.previewPanel.Size = new System.Drawing.Size(552, 88);
            this.previewPanel.TabIndex = 1;
            // ChatInputView
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.host);
            this.Controls.Add(this.previewPanel);
            this.Font = new System.Drawing.Font("Segoe UI", 10.5F);
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "ChatInputView";
            this.Size = new System.Drawing.Size(552, 88);
            this.previewPanel.ResumeLayout(false);
            this.previewPanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

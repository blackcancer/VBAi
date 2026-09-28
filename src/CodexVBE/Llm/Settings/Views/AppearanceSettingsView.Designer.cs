namespace CodexVBE
{
    /// <summary>Contrôles générés pour le thème et l’habillage natif du VBE.</summary>
public sealed partial class AppearanceSettingsView
    {
        /// <summary>Disposition des préférences d’apparence.</summary>
internal System.Windows.Forms.FlowLayoutPanel themePanel;
        /// <summary>Libellé du sélecteur de thème.</summary>
internal System.Windows.Forms.Label themeLabel;
        /// <summary>Sélecteur du thème de l’interface.</summary>
internal CodexVBE.ThemedComboBox themePicker;
        /// <summary>Option d’habillage sombre des fenêtres natives du VBE.</summary>
internal System.Windows.Forms.CheckBox nativeVbeDark;
        /// <summary>Note décrivant la portée de l’habillage natif.</summary>
internal System.Windows.Forms.Label nativeVbeDarkNote;
        /// <summary>Conteneur des composants WinForms non visuels.</summary>
private System.ComponentModel.IContainer components;
        /// <summary>Info-bulles de la vue.</summary>
private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Crée les contrôles de la vue et applique leurs propriétés Designer.</summary>
private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.themePanel = new System.Windows.Forms.FlowLayoutPanel();
            this.themeLabel = new System.Windows.Forms.Label();
            this.themePicker = new CodexVBE.ThemedComboBox();
            this.themePanel.SuspendLayout();
            this.nativeVbeDark = new System.Windows.Forms.CheckBox();
            this.nativeVbeDarkNote = new System.Windows.Forms.Label();
            this.SuspendLayout();
            this.themePanel.Name = "themePanel";
            this.themePanel.AutoSize = true;
            this.themePanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.themePanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.themePanel.WrapContents = false;
            this.themeLabel.Name = "themeLabel";
            this.themeLabel.Text = "Appearance";
            this.themeLabel.AutoSize = true;
            this.themePicker.Name = "themePicker";
            this.themePicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.themePicker.Items.AddRange(new object[] { "System", "Light", "Dark" });
            this.nativeVbeDark.Name = "nativeVbeDark";
            this.nativeVbeDark.Text = "Native dark VBE (experimental)";
            this.nativeVbeDark.AutoSize = true;
            this.nativeVbeDark.Margin = new System.Windows.Forms.Padding(3, 16, 3, 3);
            this.nativeVbeDarkNote.Name = "nativeVbeDarkNote";
            this.nativeVbeDarkNote.Text = "Applies dark native surfaces and editor colors after Settings closes. Original editor colors are saved and restored when disabled. Some legacy controls remain experimental.";
            this.nativeVbeDarkNote.AutoSize = true;
            this.nativeVbeDarkNote.MaximumSize = new System.Drawing.Size(760, 0);
            this.themePanel.Controls.Add(this.themeLabel);
            this.themePanel.Controls.Add(this.themePicker);
            this.themePanel.Controls.Add(this.nativeVbeDark);
            this.themePanel.Controls.Add(this.nativeVbeDarkNote);
            this.themePanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.Padding = new System.Windows.Forms.Padding(12);
            this.Controls.Add(this.themePanel);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "AppearanceSettingsView";
            this.Size = new System.Drawing.Size(860, 500);
            this.themePanel.Location = new System.Drawing.Point(12, 12);
            this.themePanel.Size = new System.Drawing.Size(836, 27);
            this.themePanel.TabIndex = 0;
            this.themeLabel.Location = new System.Drawing.Point(3, 0);
            this.themeLabel.Size = new System.Drawing.Size(70, 21);
            this.themeLabel.TabIndex = 0;
            this.themePicker.Location = new System.Drawing.Point(79, 3);
            this.themePicker.Size = new System.Drawing.Size(121, 23);
            this.themePicker.TabIndex = 1;
            this.themePanel.ResumeLayout(false);
            this.themePanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

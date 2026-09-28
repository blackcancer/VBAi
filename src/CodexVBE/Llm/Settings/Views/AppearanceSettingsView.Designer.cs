namespace CodexVBE
{
    public sealed partial class AppearanceSettingsView
    {
        internal System.Windows.Forms.FlowLayoutPanel themePanel;
        internal System.Windows.Forms.Label themeLabel;
        internal System.Windows.Forms.ComboBox themePicker;
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.themePanel = new System.Windows.Forms.FlowLayoutPanel();
            this.themeLabel = new System.Windows.Forms.Label();
            this.themePicker = new CodexVBE.ThemedComboBox();
            this.SuspendLayout();
            this.themePanel.Name = "themePanel";
            this.themePanel.AutoSize = true;
            this.themePanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.themeLabel.Name = "themeLabel";
            this.themeLabel.Text = "Appearance";
            this.themeLabel.AutoSize = true;
            this.themePicker.Name = "themePicker";
            this.themePicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.themePicker.Items.AddRange(new object[] { "System", "Light", "Dark" });
            this.themePanel.Controls.Add(this.themeLabel);
            this.themePanel.Controls.Add(this.themePicker);
            this.themePanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.Padding = new System.Windows.Forms.Padding(12);
            this.Controls.Add(this.themePanel);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "AppearanceSettingsView";
            this.Size = new System.Drawing.Size(860, 500);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

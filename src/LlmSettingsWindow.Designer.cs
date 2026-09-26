using System.Drawing;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class LlmSettingsWindow
    {
        private TableLayoutPanel grid;
        private FlowLayoutPanel codexActions;
        private FlowLayoutPanel buttons;
        private ComboBox provider;
        private Label codexStatus;
        private Button codexLogin;
        private Button codexRefresh;
        private TextBox openAiEndpoint;
        private TextBox ollamaEndpoint;
        private TextBox openAiKey;
        private CheckBox clearKey;
        private Button saveButton;
        private Button cancelButton;
        private Label providerLabel;
        private Label accountLabel;
        private Label authenticationLabel;
        private Label openAiEndpointLabel;
        private Label ollamaEndpointLabel;
        private Label keyLabel;
        private Label keyNote;

        private void InitializeComponent()
        {
            this.grid = new TableLayoutPanel();
            this.codexActions = new FlowLayoutPanel();
            this.buttons = new FlowLayoutPanel();
            this.provider = new ComboBox();
            this.codexStatus = new Label();
            this.codexLogin = new Button();
            this.codexRefresh = new Button();
            this.openAiEndpoint = new TextBox();
            this.ollamaEndpoint = new TextBox();
            this.openAiKey = new TextBox();
            this.clearKey = new CheckBox();
            this.saveButton = new Button();
            this.cancelButton = new Button();
            this.providerLabel = new Label();
            this.accountLabel = new Label();
            this.authenticationLabel = new Label();
            this.openAiEndpointLabel = new Label();
            this.ollamaEndpointLabel = new Label();
            this.keyLabel = new Label();
            this.keyNote = new Label();
            this.grid.SuspendLayout();
            this.codexActions.SuspendLayout();
            this.buttons.SuspendLayout();
            this.SuspendLayout();

            this.grid.Dock = DockStyle.Fill;
            this.grid.ColumnCount = 2;
            this.grid.RowCount = 8;
            this.grid.Padding = new Padding(12);
            this.grid.AutoScroll = true;
            this.grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220F));
            this.grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            this.grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
            this.grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
            this.grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
            this.grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
            this.grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
            this.grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
            this.grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
            this.grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 55F));

            this.providerLabel.Text = "Fournisseur";
            this.accountLabel.Text = "Compte ChatGPT";
            this.authenticationLabel.Text = "Authentification";
            this.openAiEndpointLabel.Text = "URL OpenAI (facultatif)";
            this.ollamaEndpointLabel.Text = "URL Ollama (facultatif)";
            this.keyLabel.Text = "Nouvelle clé OpenAI API";
            this.providerLabel.Dock = DockStyle.Fill;
            this.providerLabel.TextAlign = ContentAlignment.MiddleLeft;
            this.accountLabel.Dock = DockStyle.Fill;
            this.accountLabel.TextAlign = ContentAlignment.MiddleLeft;
            this.authenticationLabel.Dock = DockStyle.Fill;
            this.authenticationLabel.TextAlign = ContentAlignment.MiddleLeft;
            this.openAiEndpointLabel.Dock = DockStyle.Fill;
            this.openAiEndpointLabel.TextAlign = ContentAlignment.MiddleLeft;
            this.ollamaEndpointLabel.Dock = DockStyle.Fill;
            this.ollamaEndpointLabel.TextAlign = ContentAlignment.MiddleLeft;
            this.keyLabel.Dock = DockStyle.Fill;
            this.keyLabel.TextAlign = ContentAlignment.MiddleLeft;
            this.provider.DropDownStyle = ComboBoxStyle.DropDownList;
            this.provider.Dock = DockStyle.Fill;
            this.codexStatus.Text = "Vérification de ChatGPT…";
            this.codexStatus.Dock = DockStyle.Fill;
            this.codexStatus.AutoEllipsis = true;
            this.codexStatus.TextAlign = ContentAlignment.MiddleLeft;
            this.codexLogin.Text = "Se connecter à ChatGPT";
            this.codexLogin.Width = 185;
            this.codexRefresh.Text = "Actualiser l'état";
            this.codexRefresh.Width = 130;
            this.codexActions.Dock = DockStyle.Fill;
            this.codexActions.WrapContents = false;
            this.codexActions.Margin = new Padding(0);
            this.codexActions.Controls.Add(this.codexLogin);
            this.codexActions.Controls.Add(this.codexRefresh);
            this.openAiEndpoint.Dock = DockStyle.Fill;
            this.ollamaEndpoint.Dock = DockStyle.Fill;
            this.openAiKey.Dock = DockStyle.Fill;
            this.openAiKey.UseSystemPasswordChar = true;
            this.clearKey.Text = "Supprimer la clé API enregistrée";
            this.clearKey.Dock = DockStyle.Fill;
            this.keyNote.Text = "Clé vide : conserver la clé actuelle. Les secrets sont chiffrés pour ce compte Windows.";
            this.keyNote.AutoSize = true;
            this.keyNote.Dock = DockStyle.Fill;
            this.grid.Controls.Add(this.providerLabel, 0, 0);
            this.grid.Controls.Add(this.provider, 1, 0);
            this.grid.Controls.Add(this.accountLabel, 0, 1);
            this.grid.Controls.Add(this.codexStatus, 1, 1);
            this.grid.Controls.Add(this.authenticationLabel, 0, 2);
            this.grid.Controls.Add(this.codexActions, 1, 2);
            this.grid.Controls.Add(this.openAiEndpointLabel, 0, 3);
            this.grid.Controls.Add(this.openAiEndpoint, 1, 3);
            this.grid.Controls.Add(this.ollamaEndpointLabel, 0, 4);
            this.grid.Controls.Add(this.ollamaEndpoint, 1, 4);
            this.grid.Controls.Add(this.keyLabel, 0, 5);
            this.grid.Controls.Add(this.openAiKey, 1, 5);
            this.grid.Controls.Add(this.clearKey, 1, 6);
            this.grid.Controls.Add(this.keyNote, 1, 7);

            this.buttons.Dock = DockStyle.Bottom;
            this.buttons.Height = 48;
            this.buttons.FlowDirection = FlowDirection.RightToLeft;
            this.saveButton.Text = "Enregistrer";
            this.saveButton.Width = 105;
            this.cancelButton.Text = "Annuler";
            this.cancelButton.Width = 105;
            this.cancelButton.DialogResult = DialogResult.Cancel;
            this.buttons.Controls.Add(this.saveButton);
            this.buttons.Controls.Add(this.cancelButton);

            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Font = new Font("Segoe UI", 9F);
            this.ClientSize = new Size(624, 330);
            this.MinimumSize = new Size(560, 235);
            this.StartPosition = FormStartPosition.CenterParent;
            this.ShowInTaskbar = false;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Text = "CodexVBE — Configuration LLM";
            this.Controls.Add(this.grid);
            this.Controls.Add(this.buttons);
            this.CancelButton = this.cancelButton;
            this.buttons.ResumeLayout(false);
            this.codexActions.ResumeLayout(false);
            this.grid.ResumeLayout(false);
            this.grid.PerformLayout();
            this.ResumeLayout(false);
        }
    }
}

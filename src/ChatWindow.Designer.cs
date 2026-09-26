using System.Drawing;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class ChatWindow
    {
        private TableLayoutPanel rootLayout;
        private TableLayoutPanel headerLayout;
        private TableLayoutPanel actionLayout;
        private Label providerLabel;
        private Label modelLabel;
        private Label effortLabel;
        private TextBox transcript;
        private TextBox prompt;
        private Button send;
        private Label status;
        private ComboBox providerPicker;
        private ComboBox modelPicker;
        private ComboBox effortPicker;
        private Button refreshModels;
        private Button configure;

        private void InitializeComponent()
        {
            this.rootLayout = new TableLayoutPanel();
            this.headerLayout = new TableLayoutPanel();
            this.actionLayout = new TableLayoutPanel();
            this.providerLabel = new Label();
            this.modelLabel = new Label();
            this.effortLabel = new Label();
            this.transcript = new TextBox();
            this.prompt = new TextBox();
            this.send = new Button();
            this.status = new Label();
            this.providerPicker = new ComboBox();
            this.modelPicker = new ComboBox();
            this.effortPicker = new ComboBox();
            this.refreshModels = new Button();
            this.configure = new Button();
            this.rootLayout.SuspendLayout();
            this.headerLayout.SuspendLayout();
            this.actionLayout.SuspendLayout();
            this.SuspendLayout();

            this.rootLayout.Dock = DockStyle.Fill;
            this.rootLayout.Padding = new Padding(12);
            this.rootLayout.ColumnCount = 1;
            this.rootLayout.RowCount = 5;
            this.rootLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            this.rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44F));
            this.rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            this.rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 96F));
            this.rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46F));
            this.rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 27F));

            this.headerLayout.Dock = DockStyle.Fill;
            this.headerLayout.ColumnCount = 3;
            this.headerLayout.RowCount = 1;
            this.headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85F));
            this.headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            this.headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 132F));
            this.providerLabel.Text = "Fournisseur";
            this.providerLabel.Dock = DockStyle.Fill;
            this.providerLabel.TextAlign = ContentAlignment.MiddleLeft;
            this.providerPicker.Dock = DockStyle.Fill;
            this.providerPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            this.providerPicker.Margin = new Padding(4, 8, 10, 4);
            this.configure.Text = "Configuration…";
            this.configure.Dock = DockStyle.Fill;
            this.configure.Margin = new Padding(0, 5, 0, 5);
            this.headerLayout.Controls.Add(this.providerLabel, 0, 0);
            this.headerLayout.Controls.Add(this.providerPicker, 1, 0);
            this.headerLayout.Controls.Add(this.configure, 2, 0);

            this.transcript.Multiline = true;
            this.transcript.ReadOnly = true;
            this.transcript.ScrollBars = ScrollBars.Vertical;
            this.transcript.Dock = DockStyle.Fill;
            this.transcript.BackColor = Color.White;
            this.transcript.Font = new Font("Segoe UI", 9.5F);
            this.transcript.BorderStyle = BorderStyle.FixedSingle;
            this.transcript.Margin = new Padding(4, 3, 4, 8);
            this.prompt.Multiline = true;
            this.prompt.ScrollBars = ScrollBars.Vertical;
            this.prompt.Dock = DockStyle.Fill;
            this.prompt.Margin = new Padding(4, 3, 4, 4);

            this.actionLayout.Dock = DockStyle.Fill;
            this.actionLayout.ColumnCount = 6;
            this.actionLayout.RowCount = 1;
            this.actionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 52F));
            this.actionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            this.actionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 95F));
            this.actionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110F));
            this.actionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 87F));
            this.actionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 86F));
            this.modelLabel.Text = "Modèle";
            this.modelLabel.Dock = DockStyle.Fill;
            this.modelLabel.TextAlign = ContentAlignment.MiddleLeft;
            this.modelPicker.Dock = DockStyle.Fill;
            this.modelPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            this.modelPicker.Margin = new Padding(4, 7, 8, 4);
            this.effortLabel.Text = "Raisonnement";
            this.effortLabel.Dock = DockStyle.Fill;
            this.effortLabel.TextAlign = ContentAlignment.MiddleLeft;
            this.effortPicker.Dock = DockStyle.Fill;
            this.effortPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            this.effortPicker.Margin = new Padding(4, 7, 8, 4);
            this.effortPicker.Enabled = false;
            this.refreshModels.Text = "Actualiser";
            this.refreshModels.Dock = DockStyle.Fill;
            this.refreshModels.Margin = new Padding(3, 4, 3, 4);
            this.send.Text = "Envoyer";
            this.send.Dock = DockStyle.Fill;
            this.send.Margin = new Padding(5, 4, 0, 4);
            this.actionLayout.Controls.Add(this.modelLabel, 0, 0);
            this.actionLayout.Controls.Add(this.modelPicker, 1, 0);
            this.actionLayout.Controls.Add(this.effortLabel, 2, 0);
            this.actionLayout.Controls.Add(this.effortPicker, 3, 0);
            this.actionLayout.Controls.Add(this.refreshModels, 4, 0);
            this.actionLayout.Controls.Add(this.send, 5, 0);

            this.status.Text = "Chargement…";
            this.status.Dock = DockStyle.Fill;
            this.status.AutoEllipsis = true;
            this.status.TextAlign = ContentAlignment.MiddleLeft;
            this.status.Margin = new Padding(4, 0, 0, 0);
            this.rootLayout.Controls.Add(this.headerLayout, 0, 0);
            this.rootLayout.Controls.Add(this.transcript, 0, 1);
            this.rootLayout.Controls.Add(this.prompt, 0, 2);
            this.rootLayout.Controls.Add(this.actionLayout, 0, 3);
            this.rootLayout.Controls.Add(this.status, 0, 4);

            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Font = new Font("Segoe UI", 9F);
            this.MinimumSize = new Size(650, 470);
            this.ClientSize = new Size(744, 661);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowInTaskbar = true;
            this.Text = "CodexVBE — Assistant";
            this.Controls.Add(this.rootLayout);
            this.actionLayout.ResumeLayout(false);
            this.headerLayout.ResumeLayout(false);
            this.rootLayout.ResumeLayout(false);
            this.rootLayout.PerformLayout();
            this.ResumeLayout(false);
        }
    }
}

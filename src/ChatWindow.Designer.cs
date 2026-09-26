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
            SuspendLayout();
            Font = new Font("Segoe UI", 9F);
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            rootLayout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 6 };
            rootLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 104F));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            headerLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, RowCount = 2, Margin = new Padding(0, 0, 0, 12) };
            headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            headerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            headerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            providerLabel = Caption("Fournisseur");
            providerPicker = Picker("Fournisseur");
            configure = ActionButton("Configuration…");
            headerLayout.Controls.Add(providerLabel, 0, 0);
            headerLayout.Controls.Add(providerPicker, 0, 1);
            headerLayout.Controls.Add(configure, 1, 1);

            transcript = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill,
                BackColor = SystemColors.Window, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 12),
                AccessibleName = "Conversation", TabStop = true };
            actionLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, RowCount = 2, Margin = new Padding(0, 0, 0, 8) };
            actionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 62F));
            actionLayout.Padding = new Padding(0, 0, 0, 4);
            actionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38F));
            actionLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            actionLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            modelLabel = Caption("Modèle");
            effortLabel = Caption("Raisonnement");
            modelPicker = Picker("Modèle");
            effortPicker = Picker("Niveau de raisonnement");
            modelPicker.Margin = new Padding(0, 0, 8, 0);
            effortPicker.Enabled = false;
            actionLayout.Controls.Add(modelLabel, 0, 0);
            actionLayout.Controls.Add(effortLabel, 1, 0);
            actionLayout.Controls.Add(modelPicker, 0, 1);
            actionLayout.Controls.Add(effortPicker, 1, 1);
            prompt = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 8), AccessibleName = "Votre demande", AcceptsReturn = true };
            var footer = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, RowCount = 1, Margin = new Padding(0) };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            refreshModels = ActionButton("Actualiser les modèles");
            refreshModels.Anchor = AnchorStyles.Left;
            send = ActionButton("Envoyer");
            send.MinimumSize = new Size(96, 30);
            footer.Controls.Add(refreshModels, 0, 0);
            footer.Controls.Add(send, 1, 0);
            status = new Label { Text = "Chargement…", Dock = DockStyle.Fill, AutoSize = true,
                MinimumSize = new Size(0, 36), MaximumSize = new Size(0, 56), AutoEllipsis = true,
                Margin = new Padding(0, 8, 0, 0), AccessibleName = "État de la conversation" };
            rootLayout.Controls.Add(headerLayout, 0, 0);
            rootLayout.Controls.Add(transcript, 0, 1);
            rootLayout.Controls.Add(actionLayout, 0, 2);
            rootLayout.Controls.Add(prompt, 0, 3);
            rootLayout.Controls.Add(footer, 0, 4);
            rootLayout.Controls.Add(status, 0, 5);
            MinimumSize = new Size(420, 540);
            ClientSize = new Size(460, 700);
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = true;
            Text = "CodexVBE — Assistant";
            Controls.Add(rootLayout);
            ResumeLayout(true);
        }

        private static Label Caption(string text)
        {
            return new Label { Text = text, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 4) };
        }

        private static ComboBox Picker(string name)
        {
            return new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList,
                IntegralHeight = false, DropDownHeight = 300, MinimumSize = new Size(0, 25),
                Margin = new Padding(0), AccessibleName = name };
        }

        private static Button ActionButton(string text)
        {
            return new Button { Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(0, 30), Padding = new Padding(8, 2, 8, 2), Margin = new Padding(8, 0, 0, 0) };
        }
    }
}

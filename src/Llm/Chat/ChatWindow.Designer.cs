using System.Drawing;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class ChatWindow
    {
        private TableLayoutPanel rootLayout;
        private TableLayoutPanel headerLayout;
        private TableLayoutPanel actionLayout;
        private TableLayoutPanel footer;
        private Label providerLabel;
        private Label modelLabel;
        private Label effortLabel;
        private Label promptLabel;
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
            this.rootLayout = new System.Windows.Forms.TableLayoutPanel();
            this.headerLayout = new System.Windows.Forms.TableLayoutPanel();
            this.providerLabel = new System.Windows.Forms.Label();
            this.providerPicker = new System.Windows.Forms.ComboBox();
            this.configure = new System.Windows.Forms.Button();
            this.transcript = new System.Windows.Forms.TextBox();
            this.actionLayout = new System.Windows.Forms.TableLayoutPanel();
            this.modelLabel = new System.Windows.Forms.Label();
            this.effortLabel = new System.Windows.Forms.Label();
            this.modelPicker = new System.Windows.Forms.ComboBox();
            this.effortPicker = new System.Windows.Forms.ComboBox();
            this.promptLabel = new System.Windows.Forms.Label();
            this.prompt = new System.Windows.Forms.TextBox();
            this.footer = new System.Windows.Forms.TableLayoutPanel();
            this.refreshModels = new System.Windows.Forms.Button();
            this.send = new System.Windows.Forms.Button();
            this.status = new System.Windows.Forms.Label();
            this.rootLayout.SuspendLayout();
            this.headerLayout.SuspendLayout();
            this.actionLayout.SuspendLayout();
            this.footer.SuspendLayout();
            this.SuspendLayout();
            //
            // rootLayout
            //
            this.rootLayout.ColumnCount = 1;
            this.rootLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.rootLayout.Controls.Add(this.headerLayout, 0, 0);
            this.rootLayout.Controls.Add(this.transcript, 0, 1);
            this.rootLayout.Controls.Add(this.actionLayout, 0, 2);
            this.rootLayout.Controls.Add(this.promptLabel, 0, 3);
            this.rootLayout.Controls.Add(this.prompt, 0, 4);
            this.rootLayout.Controls.Add(this.footer, 0, 5);
            this.rootLayout.Controls.Add(this.status, 0, 6);
            this.rootLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rootLayout.Location = new System.Drawing.Point(0, 0);
            this.rootLayout.Name = "rootLayout";
            this.rootLayout.Padding = new System.Windows.Forms.Padding(12);
            this.rootLayout.RowCount = 7;
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 104F));
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.rootLayout.Size = new System.Drawing.Size(460, 700);
            this.rootLayout.TabIndex = 0;
            //
            // headerLayout
            //
            this.headerLayout.AutoSize = true;
            this.headerLayout.ColumnCount = 2;
            this.headerLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.headerLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.headerLayout.Controls.Add(this.providerLabel, 0, 0);
            this.headerLayout.Controls.Add(this.providerPicker, 0, 1);
            this.headerLayout.Controls.Add(this.configure, 1, 1);
            this.headerLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.headerLayout.Location = new System.Drawing.Point(12, 12);
            this.headerLayout.Margin = new System.Windows.Forms.Padding(0, 0, 0, 12);
            this.headerLayout.Name = "headerLayout";
            this.headerLayout.RowCount = 2;
            this.headerLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.headerLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.headerLayout.Size = new System.Drawing.Size(436, 49);
            this.headerLayout.TabIndex = 0;
            //
            // providerLabel
            //
            this.providerLabel.AutoSize = true;
            this.providerLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.providerLabel.Location = new System.Drawing.Point(0, 0);
            this.providerLabel.Margin = new System.Windows.Forms.Padding(0, 0, 0, 4);
            this.providerLabel.Name = "providerLabel";
            this.providerLabel.Size = new System.Drawing.Size(312, 15);
            this.providerLabel.TabIndex = 0;
            this.providerLabel.Text = "Fournisseur";
            //
            // providerPicker
            //
            this.providerPicker.AccessibleName = "Fournisseur";
            this.providerPicker.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.providerPicker.DropDownHeight = 300;
            this.providerPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.providerPicker.IntegralHeight = false;
            this.providerPicker.Location = new System.Drawing.Point(0, 23);
            this.providerPicker.Margin = new System.Windows.Forms.Padding(0);
            this.providerPicker.Name = "providerPicker";
            this.providerPicker.Size = new System.Drawing.Size(312, 23);
            this.providerPicker.TabIndex = 0;
            //
            // configure
            //
            this.configure.AutoSize = true;
            this.configure.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.configure.Location = new System.Drawing.Point(320, 19);
            this.configure.Margin = new System.Windows.Forms.Padding(8, 0, 0, 0);
            this.configure.MinimumSize = new System.Drawing.Size(0, 30);
            this.configure.Name = "configure";
            this.configure.Padding = new System.Windows.Forms.Padding(8, 2, 8, 2);
            this.configure.Size = new System.Drawing.Size(116, 30);
            this.configure.TabIndex = 1;
            this.configure.Text = "Configuration…";
            //
            // transcript
            //
            this.transcript.AccessibleName = "Conversation";
            this.transcript.BackColor = System.Drawing.SystemColors.Window;
            this.transcript.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.transcript.Dock = System.Windows.Forms.DockStyle.Fill;
            this.transcript.Location = new System.Drawing.Point(12, 73);
            this.transcript.Margin = new System.Windows.Forms.Padding(0, 0, 0, 12);
            this.transcript.Multiline = true;
            this.transcript.Name = "transcript";
            this.transcript.ReadOnly = true;
            this.transcript.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.transcript.Size = new System.Drawing.Size(436, 354);
            this.transcript.TabIndex = 1;
            //
            // actionLayout
            //
            this.actionLayout.AutoSize = true;
            this.actionLayout.ColumnCount = 2;
            this.actionLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 62F));
            this.actionLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 38F));
            this.actionLayout.Controls.Add(this.modelLabel, 0, 0);
            this.actionLayout.Controls.Add(this.effortLabel, 1, 0);
            this.actionLayout.Controls.Add(this.modelPicker, 0, 1);
            this.actionLayout.Controls.Add(this.effortPicker, 1, 1);
            this.actionLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.actionLayout.Location = new System.Drawing.Point(12, 439);
            this.actionLayout.Margin = new System.Windows.Forms.Padding(0, 0, 0, 8);
            this.actionLayout.Name = "actionLayout";
            this.actionLayout.Padding = new System.Windows.Forms.Padding(0, 0, 0, 4);
            this.actionLayout.RowCount = 2;
            this.actionLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.actionLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.actionLayout.Size = new System.Drawing.Size(436, 44);
            this.actionLayout.TabIndex = 2;
            //
            // modelLabel
            //
            this.modelLabel.AutoSize = true;
            this.modelLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.modelLabel.Location = new System.Drawing.Point(0, 0);
            this.modelLabel.Margin = new System.Windows.Forms.Padding(0, 0, 0, 4);
            this.modelLabel.Name = "modelLabel";
            this.modelLabel.Size = new System.Drawing.Size(270, 15);
            this.modelLabel.TabIndex = 0;
            this.modelLabel.Text = "Modèle";
            //
            // effortLabel
            //
            this.effortLabel.AutoSize = true;
            this.effortLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.effortLabel.Location = new System.Drawing.Point(270, 0);
            this.effortLabel.Margin = new System.Windows.Forms.Padding(0, 0, 0, 4);
            this.effortLabel.Name = "effortLabel";
            this.effortLabel.Size = new System.Drawing.Size(166, 15);
            this.effortLabel.TabIndex = 1;
            this.effortLabel.Text = "Raisonnement";
            //
            // modelPicker
            //
            this.modelPicker.AccessibleName = "Modèle";
            this.modelPicker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.modelPicker.DropDownHeight = 300;
            this.modelPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.modelPicker.DropDownWidth = 360;
            this.modelPicker.IntegralHeight = false;
            this.modelPicker.Location = new System.Drawing.Point(0, 19);
            this.modelPicker.Margin = new System.Windows.Forms.Padding(0, 0, 8, 0);
            this.modelPicker.Name = "modelPicker";
            this.modelPicker.Size = new System.Drawing.Size(262, 23);
            this.modelPicker.TabIndex = 0;
            //
            // effortPicker
            //
            this.effortPicker.AccessibleName = "Niveau de raisonnement";
            this.effortPicker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.effortPicker.DropDownHeight = 300;
            this.effortPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.effortPicker.Enabled = false;
            this.effortPicker.IntegralHeight = false;
            this.effortPicker.Location = new System.Drawing.Point(270, 19);
            this.effortPicker.Margin = new System.Windows.Forms.Padding(0);
            this.effortPicker.Name = "effortPicker";
            this.effortPicker.Size = new System.Drawing.Size(166, 23);
            this.effortPicker.TabIndex = 1;
            //
            // promptLabel
            //
            this.promptLabel.AutoSize = true;
            this.promptLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.promptLabel.Location = new System.Drawing.Point(12, 491);
            this.promptLabel.Margin = new System.Windows.Forms.Padding(0, 0, 0, 4);
            this.promptLabel.Name = "promptLabel";
            this.promptLabel.Size = new System.Drawing.Size(436, 15);
            this.promptLabel.TabIndex = 3;
            this.promptLabel.Text = "Votre demande";
            //
            // prompt
            //
            this.prompt.AcceptsReturn = true;
            this.prompt.AccessibleName = "Votre demande";
            this.prompt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.prompt.Dock = System.Windows.Forms.DockStyle.Fill;
            this.prompt.Location = new System.Drawing.Point(12, 510);
            this.prompt.Margin = new System.Windows.Forms.Padding(0, 0, 0, 8);
            this.prompt.Multiline = true;
            this.prompt.Name = "prompt";
            this.prompt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.prompt.Size = new System.Drawing.Size(436, 96);
            this.prompt.TabIndex = 3;
            //
            // footer
            //
            this.footer.AutoSize = true;
            this.footer.ColumnCount = 2;
            this.footer.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.footer.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.footer.Controls.Add(this.refreshModels, 0, 0);
            this.footer.Controls.Add(this.send, 1, 0);
            this.footer.Dock = System.Windows.Forms.DockStyle.Top;
            this.footer.Location = new System.Drawing.Point(12, 614);
            this.footer.Margin = new System.Windows.Forms.Padding(0);
            this.footer.Name = "footer";
            this.footer.RowCount = 1;
            this.footer.Size = new System.Drawing.Size(436, 30);
            this.footer.TabIndex = 4;
            //
            // refreshModels
            //
            this.refreshModels.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.refreshModels.AutoSize = true;
            this.refreshModels.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.refreshModels.Location = new System.Drawing.Point(0, 0);
            this.refreshModels.Margin = new System.Windows.Forms.Padding(0);
            this.refreshModels.MinimumSize = new System.Drawing.Size(0, 30);
            this.refreshModels.Name = "refreshModels";
            this.refreshModels.Padding = new System.Windows.Forms.Padding(8, 2, 8, 2);
            this.refreshModels.Size = new System.Drawing.Size(150, 30);
            this.refreshModels.TabIndex = 1;
            this.refreshModels.Text = "Actualiser les modèles";
            //
            // send
            //
            this.send.AutoSize = true;
            this.send.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.send.Location = new System.Drawing.Point(340, 0);
            this.send.Margin = new System.Windows.Forms.Padding(8, 0, 0, 0);
            this.send.MinimumSize = new System.Drawing.Size(96, 30);
            this.send.Name = "send";
            this.send.Padding = new System.Windows.Forms.Padding(8, 2, 8, 2);
            this.send.Size = new System.Drawing.Size(96, 30);
            this.send.TabIndex = 0;
            this.send.Text = "Envoyer";
            //
            // status
            //
            this.status.AccessibleName = "État de la conversation";
            this.status.AutoEllipsis = true;
            this.status.AutoSize = true;
            this.status.Dock = System.Windows.Forms.DockStyle.Fill;
            this.status.Location = new System.Drawing.Point(12, 652);
            this.status.Margin = new System.Windows.Forms.Padding(0, 8, 0, 0);
            this.status.MaximumSize = new System.Drawing.Size(0, 56);
            this.status.MinimumSize = new System.Drawing.Size(0, 36);
            this.status.Name = "status";
            this.status.Size = new System.Drawing.Size(436, 36);
            this.status.TabIndex = 5;
            this.status.Text = "Chargement…";
            //
            // ChatWindow
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(460, 700);
            this.Controls.Add(this.rootLayout);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.MinimumSize = new System.Drawing.Size(420, 540);
            this.Name = "ChatWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CodexVBE — Assistant";
            this.rootLayout.ResumeLayout(false);
            this.rootLayout.PerformLayout();
            this.headerLayout.ResumeLayout(false);
            this.headerLayout.PerformLayout();
            this.actionLayout.ResumeLayout(false);
            this.actionLayout.PerformLayout();
            this.footer.ResumeLayout(false);
            this.footer.PerformLayout();
            this.ResumeLayout(false);

        }
    }
}

using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace CodexVBE
{
    internal sealed partial class ChatWindow
    {
        private ElementHost shellHost;

        private void InitializeComponent()
        {
            this.shellHost = new ElementHost();
            this.SuspendLayout();
            this.shellHost.Dock = DockStyle.Fill;
            this.shellHost.Location = new Point(0, 0);
            this.shellHost.Name = "shellHost";
            this.shellHost.Size = new Size(600, 820);
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.ClientSize = new Size(600, 820);
            this.Controls.Add(this.shellHost);
            this.Font = new Font("Segoe UI", 9F);
            this.MinimumSize = new Size(440, 560);
            this.Name = "ChatWindow";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "CodexVBE — Assistant";
            this.ResumeLayout(false);
        }
    }
}

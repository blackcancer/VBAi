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
            SuspendLayout();
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(600, 820);
            MinimumSize = new Size(440, 560);
            Font = new Font("Segoe UI", 9F);
            Text = "CodexVBE — Assistant";
            Name = "ChatWindow";
            StartPosition = FormStartPosition.CenterScreen;
            shellHost = new ElementHost { Dock = DockStyle.Fill };
            Controls.Add(shellHost);
            InitializeShell();
            ResumeLayout(false);
        }
    }
}

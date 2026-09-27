namespace CodexVBE
{
    internal sealed partial class VbeApprovalDialog
    {
        private System.Windows.Forms.TextBox details;
        private System.Windows.Forms.FlowLayoutPanel actions;
        private System.Windows.Forms.Button approve;
        private System.Windows.Forms.Button reject;

        private void InitializeComponent()
        {
            this.details = new System.Windows.Forms.TextBox();
            this.actions = new System.Windows.Forms.FlowLayoutPanel();
            this.approve = new System.Windows.Forms.Button();
            this.reject = new System.Windows.Forms.Button();
            this.actions.SuspendLayout();
            this.SuspendLayout();
            //
            // details
            //
            this.details.Dock = System.Windows.Forms.DockStyle.Fill;
            this.details.Multiline = true;
            this.details.Name = "details";
            this.details.ReadOnly = true;
            this.details.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.details.TabIndex = 0;
            this.details.WordWrap = false;
            //
            // actions
            //
            this.actions.Controls.Add(this.approve);
            this.actions.Controls.Add(this.reject);
            this.actions.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.actions.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.actions.Height = 44;
            this.actions.Name = "actions";
            this.actions.TabIndex = 1;
            //
            // approve
            //
            this.approve.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.approve.Name = "approve";
            this.approve.Size = new System.Drawing.Size(100, 30);
            this.approve.TabIndex = 0;
            this.approve.Text = "Autoriser";
            //
            // reject
            //
            this.reject.DialogResult = System.Windows.Forms.DialogResult.No;
            this.reject.Name = "reject";
            this.reject.Size = new System.Drawing.Size(100, 30);
            this.reject.TabIndex = 1;
            this.reject.Text = "Refuser";
            //
            // VbeApprovalDialog
            //
            this.CancelButton = this.reject;
            this.ClientSize = new System.Drawing.Size(740, 530);
            this.Controls.Add(this.details);
            this.Controls.Add(this.actions);
            this.MaximizeBox = true;
            this.MinimizeBox = false;
            this.Name = "VbeApprovalDialog";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "CodexVBE — valider la modification";
            this.actions.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

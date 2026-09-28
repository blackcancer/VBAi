namespace CodexVBE
{
    public sealed partial class ChatToolWindow
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.Timer siteResizeTimer;
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.siteResizeTimer = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            this.siteResizeTimer.Interval = 300;
            this.siteResizeTimer.Tick += new System.EventHandler(this.SiteResizeTimer_Tick);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "ChatToolWindow";
            this.Size = new System.Drawing.Size(520, 760);
            this.ResumeLayout(false);
        }
    }
}

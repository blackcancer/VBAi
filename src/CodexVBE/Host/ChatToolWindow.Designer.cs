namespace CodexVBE
{
    /// <summary>Contrôles générés du site WinForms hébergé dans le volet natif VBE.</summary>
public sealed partial class ChatToolWindow
    {
        /// <summary>Container that owns the disposable components created by the WinForms Designer.</summary>
private System.ComponentModel.IContainer components;
        /// <summary>Minuteur qui suit les dimensions de la zone native du volet.</summary>
private System.Windows.Forms.Timer siteResizeTimer;
        /// <summary>Creates and configures the chat tool window controls serialized by the WinForms Designer.</summary>
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

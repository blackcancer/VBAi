namespace CodexVBE
{
    public sealed partial class ChatContextPreviewView
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.GroupBox section;
        private System.Windows.Forms.TextBox content;
        /// <summary>Libère les composants du modèle Designer.</summary>
        protected override void Dispose(bool disposing) { if (disposing && components != null) components.Dispose(); base.Dispose(disposing); }
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.section = new System.Windows.Forms.GroupBox();
            this.content = new System.Windows.Forms.TextBox();
            this.section.SuspendLayout();
            this.SuspendLayout();
            // content
            this.content.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.content.Dock = System.Windows.Forms.DockStyle.Fill;
            this.content.Location = new System.Drawing.Point(6, 22);
            this.content.Multiline = true;
            this.content.Name = "content";
            this.content.ReadOnly = true;
            this.content.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.content.Size = new System.Drawing.Size(488, 82);
            this.content.TabIndex = 0;
            // section
            this.section.Controls.Add(this.content);
            this.section.Dock = System.Windows.Forms.DockStyle.Fill;
            this.section.Location = new System.Drawing.Point(0, 0);
            this.section.Name = "section";
            this.section.Padding = new System.Windows.Forms.Padding(6);
            this.section.Size = new System.Drawing.Size(500, 110);
            this.section.TabIndex = 0;
            this.section.TabStop = false;
            this.section.Text = "Context";
            // ChatContextPreviewView
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.section);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "ChatContextPreviewView";
            this.Size = new System.Drawing.Size(500, 110);
            this.section.ResumeLayout(false);
            this.section.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

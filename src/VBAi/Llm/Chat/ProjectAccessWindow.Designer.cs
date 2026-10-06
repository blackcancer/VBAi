namespace VBAi
{

    /// <summary>Collects the projects and shared context that the chat assistant may read.</summary>
    partial class ProjectAccessWindow
    {

        /// <summary>Container that owns tooltip and control components created by this window.</summary>
        private System.ComponentModel.IContainer components;

        /// <summary>Static explanation of bound-project read access and additional-project consent.</summary>
        private System.Windows.Forms.Label explanation;

        /// <summary>Checklist of additional projects that may be granted read access to the new chat.</summary>
        private VBAi.UiCheckedListBox projectList;

        /// <summary>Consent toggle for shared VBE, debugger-window, and clipboard context in this chat.</summary>
        private System.Windows.Forms.CheckBox sharedContext;

        /// <summary>Explains data transmitted by shared context and that consent grants no edit authority.</summary>
        private System.Windows.Forms.Label sharedExplanation;

        /// <summary>Accept button that applies selected access settings to a new conversation.</summary>
        private VBAi.UiActionButton applyButton;

        /// <summary>Tracks the cancel button state of project access window.</summary>
        private VBAi.UiActionButton cancelButton;

        /// <summary>ToolTip component containing project and shared-context consent explanations.</summary>
        private System.Windows.Forms.ToolTip toolTip;

        /// <summary>Unsubscribes from theme changes and disposes Designer-owned components.</summary>
        /// <param name="disposing">True when managed components should be disposed.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) { UiTheme.Changed -= ApplyAppearance; components?.Dispose(); }
            base.Dispose(disposing);
        }

        /// <summary>Creates and lays out the fixed access-consent controls and their tooltip text.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.explanation = new System.Windows.Forms.Label();
            this.projectList = new VBAi.UiCheckedListBox();
            this.sharedContext = new System.Windows.Forms.CheckBox();
            this.sharedExplanation = new System.Windows.Forms.Label();
            this.applyButton = new VBAi.UiActionButton();
            this.cancelButton = new VBAi.UiActionButton();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            this.explanation.Name = "explanation";
            this.projectList.Name = "projectList";
            this.sharedContext.Name = "sharedContext";
            this.sharedExplanation.Name = "sharedExplanation";
            this.applyButton.Name = "applyButton";
            this.cancelButton.Name = "cancelButton";
            this.projectList.TabIndex = 0;
            this.sharedContext.TabIndex = 1;
            this.applyButton.TabIndex = 2;
            this.cancelButton.TabIndex = 3;
            this.Name = "ProjectAccessWindow";
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.explanation.Location = new System.Drawing.Point(18, 18);
            this.explanation.Size = new System.Drawing.Size(524, 64);
            this.explanation.Text = "The bound project is always readable. Select additional projects for read access only. Changing access starts a new conversation; previously transmitted data cannot be recalled.";
            this.projectList.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.projectList.Location = new System.Drawing.Point(18, 90);
            this.projectList.Size = new System.Drawing.Size(524, 148);
            this.projectList.CheckOnClick = true;
            this.sharedContext.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left;
            this.sharedContext.Location = new System.Drawing.Point(18, 252);
            this.sharedContext.Size = new System.Drawing.Size(524, 30);
            this.sharedContext.Text = "Allow shared VBE and clipboard context";
            this.sharedExplanation.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.sharedExplanation.Location = new System.Drawing.Point(18, 288);
            this.sharedExplanation.Size = new System.Drawing.Size(524, 60);
            this.sharedExplanation.Text = "Shared tools may transmit data from any open project, debugger windows, native inventories and the system clipboard. This permission does not authorize edits in other projects.";
            this.applyButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
            this.applyButton.Location = new System.Drawing.Point(290, 362);
            this.applyButton.Size = new System.Drawing.Size(148, 32);
            this.applyButton.Text = "Apply to new chat";
            this.applyButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.cancelButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
            this.cancelButton.Location = new System.Drawing.Point(448, 362);
            this.cancelButton.Size = new System.Drawing.Size(94, 32);
            this.cancelButton.Text = "Cancel";
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.toolTip.SetToolTip(this.projectList, "Additional projects can be read, but only the bound project can be edited.");
            this.toolTip.SetToolTip(this.sharedContext, "Authorize unfiltered native debugger, window and clipboard data for this chat.");
            this.AcceptButton = this.applyButton;
            this.CancelButton = this.cancelButton;
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.ClientSize = new System.Drawing.Size(560, 412);
            this.MinimumSize = new System.Drawing.Size(576, 451);
            this.Controls.Add(this.explanation);
            this.Controls.Add(this.projectList);
            this.Controls.Add(this.sharedContext);
            this.Controls.Add(this.sharedExplanation);
            this.Controls.Add(this.applyButton);
            this.Controls.Add(this.cancelButton);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Project access";
            this.applyButton.Symbol = VBAi.UiSymbol.Check;
            this.applyButton.Primary = true;
            this.cancelButton.Symbol = VBAi.UiSymbol.Close;
            this.cancelButton.IconOnly = true;
            this.cancelButton.AutoSize = false;
            this.cancelButton.MinimumSize = System.Drawing.Size.Empty;
            this.cancelButton.Size = new System.Drawing.Size(32, 30);
            this.ResumeLayout(false);
        }
    }
}

namespace CodexVBE
{
    /// <summary>Collects the projects and shared context that the chat assistant may read.</summary>
    partial class ProjectAccessWindow
    {
        /// <summary>Stores the components used by ProjectAccessWindow.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Stores the explanation used by ProjectAccessWindow.</summary>
        private System.Windows.Forms.Label explanation;
        /// <summary>Stores the project list used by ProjectAccessWindow.</summary>
        private System.Windows.Forms.CheckedListBox projectList;
        /// <summary>Stores the shared context used by ProjectAccessWindow.</summary>
        private System.Windows.Forms.CheckBox sharedContext;
        /// <summary>Stores the shared explanation used by ProjectAccessWindow.</summary>
        private System.Windows.Forms.Label sharedExplanation;
        /// <summary>Stores the apply button used by ProjectAccessWindow.</summary>
        private System.Windows.Forms.Button applyButton;
        /// <summary>Stores the cancel button used by ProjectAccessWindow.</summary>
        private System.Windows.Forms.Button cancelButton;
        /// <summary>Stores the tool tip used by ProjectAccessWindow.</summary>
        private System.Windows.Forms.ToolTip toolTip;
        /// <summary>Performs the dispose operation for ProjectAccessWindow.</summary>
        /// <param name="disposing">Indicates whether disposing is enabled.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) { UiTheme.Changed -= ApplyAppearance; components?.Dispose(); }
            base.Dispose(disposing);
        }
        /// <summary>Performs the initialize component operation for ProjectAccessWindow.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.explanation = new System.Windows.Forms.Label();
            this.projectList = new System.Windows.Forms.CheckedListBox();
            this.sharedContext = new System.Windows.Forms.CheckBox();
            this.sharedExplanation = new System.Windows.Forms.Label();
            this.applyButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
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
            this.ResumeLayout(false);
        }
    }
}

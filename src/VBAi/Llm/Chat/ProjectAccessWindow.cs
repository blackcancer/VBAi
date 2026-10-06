using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Collects the projects and shared context that the chat assistant may read.</summary>
    internal sealed partial class ProjectAccessWindow : Form
    {

        /// <summary>Initializes a ProjectAccessWindow instance with the supplied state.</summary>
        public ProjectAccessWindow()
        {
            InitializeComponent();
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            UiText.Apply(this, components);
            Icon = VbeWindowIcons.Icon("assistant");
            ApplyAppearance();
            UiTheme.Changed += ApplyAppearance;
        }

        /// <summary>Handles populate for project access window.</summary>
        /// <param name="projects">i enumerable&lt;key value pair&lt;string, string&gt;&gt; that supplies the projects for this operation.</param>
        /// <param name="grants">i enumerable&lt;string&gt; that supplies the grants for this operation.</param>
        /// <param name="shared">Indicates whether shared is enabled.</param>
        internal void Populate(IEnumerable<KeyValuePair<string, string>> projects, IEnumerable<string> grants, bool shared)
        {
            var selected = new HashSet<string>(grants ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            foreach (var project in projects)
                projectList.Items.Add(new ProjectChoice { Selector = project.Key, Label = project.Value }, selected.Contains(project.Key));
            sharedContext.Checked = shared;
        }

        /// <summary>Gets the selected projects.</summary>
        /// <value>Current selected projects exposed by project access window.</value>
        internal string[] SelectedProjects => projectList.CheckedItems.Cast<ProjectChoice>().Select(item => item.Selector).ToArray();

        /// <summary>Gets the shared context.</summary>
        /// <value>Current shared context exposed by project access window.</value>
        internal bool SharedContext => sharedContext.Checked;

        /// <summary>Owns the project choice state and operations.</summary>
        private sealed class ProjectChoice
        {

            /// <summary>Maintains the selector and label state for project choice.</summary>
            internal string Selector, Label;

            /// <summary>Handles to string for project choice.</summary>
            /// <returns>Text produced by the operation for to string on project choice.</returns>
            public override string ToString() => Label;
        }

        /// <summary>Handles apply appearance for project access window.</summary>
        private void ApplyAppearance()
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(ApplyAppearance)); return; }
            UiTheme.Apply(this);
        }
    }
}

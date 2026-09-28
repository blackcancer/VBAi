using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;

namespace CodexVBE
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

        /// <summary>Performs the populate operation for ProjectAccessWindow.</summary>
        /// <param name="projects">The projects used by this operation.</param>
        /// <param name="grants">The grants used by this operation.</param>
        /// <param name="shared">Indicates whether shared is enabled.</param>
        internal void Populate(IEnumerable<KeyValuePair<string, string>> projects, IEnumerable<string> grants, bool shared)
        {
            var selected = new HashSet<string>(grants ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            foreach (var project in projects)
                projectList.Items.Add(new ProjectChoice { Selector = project.Key, Label = project.Value }, selected.Contains(project.Key));
            sharedContext.Checked = shared;
        }

        /// <summary>Gets the selected projects.</summary>
        /// <value>The current value represented by this member.</value>
        internal string[] SelectedProjects => projectList.CheckedItems.Cast<ProjectChoice>().Select(item => item.Selector).ToArray();
        /// <summary>Gets the shared context.</summary>
        /// <value>The current value represented by this member.</value>
        internal bool SharedContext => sharedContext.Checked;
        /// <summary>Provides the project choice implementation.</summary>
        private sealed class ProjectChoice
        {
            /// <summary>Stores the selector,label used by ProjectChoice.</summary>
            internal string Selector, Label;
            /// <summary>Performs the to string operation for ProjectChoice.</summary>
            /// <returns>The result produced by this operation.</returns>
            public override string ToString() => Label;
        }
        /// <summary>Performs the apply appearance operation for ProjectAccessWindow.</summary>
        private void ApplyAppearance()
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(ApplyAppearance)); return; }
            UiTheme.Apply(this);
        }
    }
}

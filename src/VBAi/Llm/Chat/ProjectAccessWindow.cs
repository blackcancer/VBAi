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

        /// <summary>Creates the project-consent dialog, applies theme resources, and subscribes to later theme changes.</summary>
        public ProjectAccessWindow()
        {
            InitializeComponent();
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            UiText.Apply(this, components);
            Icon = VbeWindowIcons.Icon("assistant");
            ApplyAppearance();
            UiTheme.Changed += ApplyAppearance;
        }

        /// <summary>Loads available additional projects and restores previously granted read/shared-context choices.</summary>
        /// <param name="projects">Selector and display-label pairs for projects the chat may request read access to.</param>
        /// <param name="grants">Previously selected project selectors; matching is case-insensitive.</param>
        /// <param name="shared">Whether the shared VBE/clipboard context consent should be checked.</param>
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

        /// <summary>Selector/display pair shown in the additional-project checklist.</summary>
        private sealed class ProjectChoice
        {

            /// <summary>Stable project selector submitted as a read grant and user-visible display label.</summary>
            internal string Selector, Label;

            /// <summary>Supplies the display text rendered by the checked-list control.</summary>
            /// <returns>Project label; the selector remains separate from displayed text.</returns>
            public override string ToString() => Label;
        }

        /// <summary>Applies the current theme on the window's UI thread, posting when invoked from another thread.</summary>
        private void ApplyAppearance()
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(ApplyAppearance)); return; }
            UiTheme.Apply(this);
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class ProjectAccessWindow : Form
    {
        internal ProjectAccessWindow()
        {
            InitializeComponent();
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            UiText.Apply(this, components);
            Icon = VbeWindowIcons.Icon("assistant");
            ApplyAppearance();
            UiTheme.Changed += ApplyAppearance;
        }

        internal void Populate(IEnumerable<KeyValuePair<string, string>> projects, IEnumerable<string> grants, bool shared)
        {
            var selected = new HashSet<string>(grants ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            foreach (var project in projects)
                projectList.Items.Add(new ProjectChoice { Selector = project.Key, Label = project.Value }, selected.Contains(project.Key));
            sharedContext.Checked = shared;
        }

        internal string[] SelectedProjects => projectList.CheckedItems.Cast<ProjectChoice>().Select(item => item.Selector).ToArray();
        internal bool SharedContext => sharedContext.Checked;
        private sealed class ProjectChoice
        {
            internal string Selector, Label;
            public override string ToString() => Label;
        }
        private void ApplyAppearance()
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(ApplyAppearance)); return; }
            UiTheme.Apply(this);
        }
    }
}

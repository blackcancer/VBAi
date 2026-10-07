using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks localized usage, Designer isolation, tooltip ownership and the complete fixed-interface inventory.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class UiHelpHintsTests
    {
        [STATestMethod]
        public void UnknownAndUnnamedControlsDoNotAcquireInventedHints()
        {
            using (var control = new Button())
            {
                Assert.IsNull(UiHelpHints.ForControl(null));
                Assert.IsNull(UiHelpHints.ForControl(control));
                control.Name = "NotInCatalogue";
                Assert.IsNull(UiHelpHints.ForControl(control));
                UiHelpHints.Apply(control);
                Assert.IsNull(control.AccessibleDescription);
            }
        }

        [STATestMethod]
        public void MenuHintsExplainTheCommandAndPreserveContextualOverrides()
        {
            using (var menu = new ToolStripMenuItem("Copy code"))
            {
                UiHelpHints.Apply(menu);
                Assert.IsFalse(string.IsNullOrWhiteSpace(menu.ToolTipText));
                Assert.AreEqual(menu.ToolTipText, menu.AccessibleDescription);
                menu.ToolTipText = "Current code block only";
                UiHelpHints.Apply(menu);
                Assert.AreEqual("Current code block only", menu.ToolTipText);
            }
        }

        [STATestMethod]
        public void EveryFixedInteractiveControlHasAUsageHintOrAnExistingContextualHint()
        {
            var context = LicenseManager.CurrentContext;
            var missing = new System.Collections.Generic.List<string>();
            var flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            try
            {
                LicenseManager.CurrentContext = new DesigntimeLicenseContext();
                var types = typeof(UiHelpHints).Assembly.GetTypes().Where(type => !type.IsAbstract
                    && type.Namespace == "VBAi" && typeof(Control).IsAssignableFrom(type)
                    && (typeof(Form).IsAssignableFrom(type) || typeof(UserControl).IsAssignableFrom(type))
                    && type.GetConstructor(flags, null, Type.EmptyTypes, null) != null).ToArray();
                Assert.IsTrue(types.Length >= 48, "The inventory must cover product views, not a hand-picked subset.");
                foreach (var type in types)
                    using (var root = (Control)Activator.CreateInstance(type, true))
                    {
                        foreach (var control in Descendants(root).Where(IsInteractive))
                        {
                            if (!string.IsNullOrEmpty(UiHelpHints.ForControl(control))) continue;
                            bool found = false;
                            for (var owner = control; owner != null; owner = owner.Parent)
                                foreach (var field in owner.GetType().GetFields(flags))
                                    if (typeof(ToolTip).IsAssignableFrom(field.FieldType)
                                        && field.GetValue(owner) is ToolTip tips && !string.IsNullOrWhiteSpace(tips.GetToolTip(control))) found = true;
                            if (!found) missing.Add(type.Name + "." + control.Name + " (" + control.GetType().Name + ")");
                        }
                    }
            }
            finally { LicenseManager.CurrentContext = context; }
            Assert.AreEqual(0, missing.Count, string.Join(Environment.NewLine, missing.Distinct()));
        }

        /// <summary>Includes inputs, actions and navigable data surfaces; decorative layouts are excluded.</summary>
        private static bool IsInteractive(Control value) => value is ButtonBase || value is TextBoxBase
            || value is ComboBox || value is ListBox || value is TreeView || value is DataGridView
            || value is NumericUpDown || value is TrackBar || value is LinkLabel || value is TabControl;

        [STATestMethod]
        public void FrenchUsagePreservesTechnicalNamesAndHostCulture()
        {
            using (var culture = new LocalizationScope())
            using (var control = new TextBox { Name = "repositoryName" })
            {
                var before = System.Globalization.CultureInfo.CurrentCulture;
                LocalizationScope.Set("fr-FR");
                StringAssert.Contains(UiHelpHints.ForControl(control), "dépôt GitHub");
                Assert.AreSame(before, System.Globalization.CultureInfo.CurrentCulture);
                LocalizationScope.Set("en-US");
                StringAssert.Contains(UiHelpHints.ForControl(control), "GitHub repository");
            }
        }

        [STATestMethod]
        public void ExistingContextualHintIsNotReplacedOrCopiedToAnotherPopup()
        {
            using (var button = new Button { Name = "push" })
            using (var components = new Container())
            {
                var first = new ToolTip(components);
                var second = new ToolTip(components);
                first.SetToolTip(button, "Current repository: Fixture only");
                UiHelpHints.Apply(button, components);
                Assert.AreEqual("Current repository: Fixture only", first.GetToolTip(button));
                Assert.AreEqual("", second.GetToolTip(button));
            }
        }

        [STATestMethod]
        public void DesignerApplyDoesNotAllocateOrModifyTooltips()
        {
            var context = LicenseManager.CurrentContext;
            try
            {
                LicenseManager.CurrentContext = new DesigntimeLicenseContext();
                using (var button = new Button { Name = "push" })
                using (var tips = new ToolTip())
                {
                    UiHelpHints.Apply(button, null, tips);
                    Assert.AreEqual("", tips.GetToolTip(button));
                    Assert.IsNull(button.AccessibleDescription);
                }
            }
            finally { LicenseManager.CurrentContext = context; }
        }

        [STATestMethod]
        public void SuppliedTooltipIsBorrowedAndApplyIsIdempotent()
        {
            using (var control = new Button { Name = "branchCreate" })
            using (var tips = new ToolTip())
            {
                UiHelpHints.Apply(control, null, tips);
                var text = tips.GetToolTip(control);
                Assert.IsFalse(string.IsNullOrEmpty(text));
                UiHelpHints.Apply(control, null, tips);
                Assert.AreEqual(text, tips.GetToolTip(control));
                Assert.AreEqual(text, control.AccessibleDescription);
                control.Dispose();
                using (var next = new Button()) { tips.SetToolTip(next, "Still owned by caller"); }
            }
        }

        [STATestMethod]
        public void MostSpecificViewRuleWinsOverSharedControlName()
        {
            using (var view = new GitChangesView())
            {
                var changes = view.Controls.Cast<Control>().SelectMany(Descendants).First(c => c.Name == "changes");
                StringAssert.Contains(UiHelpHints.ForControl(changes), UiText.Get("Select the VBA source changes to review and include in a commit."));
            }
        }

        [STATestMethod]
        public void ThemeTraversalIncludesDynamicallyAddedControls()
        {
            using (var root = new Panel())
            using (var button = new Button { Name = "branchCreate" })
            {
                root.Controls.Add(button);
                UiTheme.Apply(root);
                Assert.AreEqual(UiHelpHints.ForControl(button), button.AccessibleDescription);
                UiTheme.Apply(root);
                root.Dispose();
            }
        }

        /// <summary>Walks existing children without invoking callbacks or creating another window.</summary>
        private static System.Collections.Generic.IEnumerable<Control> Descendants(Control root)
        {
            yield return root;
            foreach (Control child in root.Controls)
                foreach (var nested in Descendants(child)) yield return nested;
        }
    }
}

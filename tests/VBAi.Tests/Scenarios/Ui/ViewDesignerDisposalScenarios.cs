using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.ComponentModel;
using System.Reflection;
using System.Windows.Forms;

namespace VBAi.Tests.Scenarios
{
    /// <summary>Vérifie le cycle de vie des vues fixes sans ouvrir les services Git ou LLM.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ViewDesignerDisposalScenarios
    {
        /// <summary>Vérifie la conservation hors disposal managé et la libération avec ou sans conteneur.</summary>
        [STATestMethod]
        [DataRow("GitBranchesView")]
        [DataRow("GitChangesView")]
        [DataRow("GitCheckpointsView")]
        [DataRow("GitConflictsView")]
        [DataRow("GitConnectionView")]
        [DataRow("GitHistoryView")]
        [DataRow("GitHubPullChecksView")]
        [DataRow("GitHubPullCommentsView")]
        [DataRow("GitHubPullComposeView")]
        [DataRow("GitHubPullDetailsView")]
        [DataRow("GitHubPullFilesView")]
        [DataRow("GitHubPullRequestsView")]
        [DataRow("GitHubRepositoriesView")]
        [DataRow("GitImportView")]
        [DataRow("AppearanceSettingsView")]
        [DataRow("GitHubAccountSettingsView")]
        [DataRow("ProviderSettingsView")]
        public void FixedViewsReleaseTheirOwnedComponentsAndTolerateAbsentContainers(string name)
        {
            var type = typeof(CodeDiffView).Assembly.GetType("VBAi." + name, true);
            var field = type.GetField("components", BindingFlags.Instance | BindingFlags.NonPublic);
            var dispose = type.GetMethod("Dispose", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.DeclaredOnly, null, new[] { typeof(bool) }, null);
            Assert.IsNotNull(field); Assert.IsNotNull(dispose);
            using (var view = (UserControl)Activator.CreateInstance(type))
            using (var tracked = new Component())
            {
                bool disposed = false;
                tracked.Disposed += (sender, args) => disposed = true;
                ((IContainer)field.GetValue(view)).Add(tracked);
                dispose.Invoke(view, new object[] { false }); Assert.IsFalse(disposed);
                view.Dispose(); Assert.IsTrue(disposed); Assert.IsTrue(view.IsDisposed);
            }
            using (var view = (UserControl)Activator.CreateInstance(type))
            {
                ((IContainer)field.GetValue(view)).Dispose(); field.SetValue(view, null);
                dispose.Invoke(view, new object[] { false });
                view.Dispose(); Assert.IsTrue(view.IsDisposed);
            }
        }
    }
}

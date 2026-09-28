using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class UpdateWindowTests
    {
        [STATestMethod]
        public void WindowKeepsSettingsAndReleaseNotesInDesignerControlsWithoutAnyNativeInstaller()
        {
            using (var scope = new UpdateScope())
            using (var window = new UpdateWindow())
            {
                var prefs = new UpdatePreferences(); window.ReadPreferences = () => prefs; window.StorePreferences = p => prefs = p;
                window.ManagedInstallation = () => false;
                window.CreateFeed = () => new UpdateFeed(new UpdateFeedTests.Handler("[{\"tag_name\":\"v1.2.3\",\"body\":\"Fixture release notes\"}]"), ct => throw new AssertFailedException());
                LlmBoundaryScope.Call(window, "Check_Click", null, EventArgs.Empty);
                Assert.AreEqual("Fixture release notes", UiInvoke.Field<TextBox>(window, "notes").Text);
                Assert.IsFalse(UiInvoke.Field<Button>(window, "install").Enabled); Assert.IsFalse(UiInvoke.Field<Button>(window, "download").Enabled);
                UiInvoke.Field<CheckBox>(window, "automaticDownload").Checked = false;
                LlmBoundaryScope.Call(window, "Save_Click", null, EventArgs.Empty); Assert.IsFalse(prefs.DownloadAutomatically); Assert.IsFalse(prefs.InstallAutomatically);
            }
        }
        [STATestMethod]
        public void WindowLocalizesAllCulturesAndCanCancelAWaitingJob()
        {
            using (var scope = new UpdateScope())
            {
                var job = UpdateInstallerRunnerTests.Job(scope);
                using (var window = new UpdateWindow())
                {
                    LlmBoundaryScope.Call(window, "CancelPending_Click", null, EventArgs.Empty);
                    Assert.IsTrue(UpdateInstallJob.Load(scope.Root).Completed);
                }
                foreach (var language in UiLanguages.All)
                using (var culture = new LocalizationScope(language.CultureName))
                using (var window = new UpdateWindow())
                { Assert.AreEqual(UiText.Get("VBAi updates"), window.Text); Assert.AreEqual(RightToLeft.No, UiInvoke.Field<TextBox>(window, "notes").RightToLeft); }
            }
        }
    }
}

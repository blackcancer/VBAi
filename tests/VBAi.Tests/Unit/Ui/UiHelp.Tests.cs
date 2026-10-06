using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks local-only chapter routing, help availability and inert Designer construction.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class UiHelpTests
    {
        [TestMethod]
        public void KnownFormsRouteToTaskChaptersAndUnknownFormsRouteToStart()
        {
            Assert.AreEqual("git-conversion", UiHelp.TopicFor(typeof(GitWindow)));
            Assert.AreEqual("settings", UiHelp.TopicFor(typeof(LlmSettingsWindow)));
            Assert.AreEqual("privacy", UiHelp.TopicFor(typeof(ProjectAccessWindow)));
            Assert.AreEqual("start", UiHelp.TopicFor(typeof(Form)));
            Assert.AreEqual("start", UiHelp.TopicFor(null));
            Assert.AreEqual("VBAi.fr-FR.chm", Path.GetFileName(UiHelp.FilePath));
            Assert.IsTrue(Path.IsPathRooted(UiHelp.FilePath));
        }

        [STATestMethod]
        public void ExplicitGuideActionOpensOnlyTheFixedLocalPathAndSanitizesTheTopic()
        {
            var exists = UiHelp.Exists; var launch = UiHelp.Launch; var notice = UiHelp.ShowUnavailable;
            try
            {
                string openedPath = null, openedTopic = null;
                int launches = 0, notices = 0;
                UiHelp.Exists = path => path == UiHelp.FilePath;
                UiHelp.Launch = (owner, path, topic) => { openedPath = path; openedTopic = topic; launches++; };
                UiHelp.ShowUnavailable = owner => notices++;
                using (var window = new AboutWindow())
                {
                    Assert.AreEqual(0, launches, "Construction must not open help.");
                    UiInvoke.Call(typeof(AboutWindow), "Help_Click", window, null, EventArgs.Empty);
                    Assert.AreEqual(UiHelp.FilePath, openedPath);
                    Assert.AreEqual("start", openedTopic);
                    UiHelp.Open(window, "../../external.html");
                    Assert.AreEqual("start", openedTopic);
                    UiHelp.Open(window, "git-conversion");
                    Assert.AreEqual("git-conversion", openedTopic);
                    Assert.AreEqual(3, launches);
                    Assert.AreEqual(0, notices);
                    UiHelp.Exists = path => false;
                    UiHelp.Open(window);
                    Assert.AreEqual(3, launches);
                    Assert.AreEqual(1, notices);
                }
            }
            finally { UiHelp.Exists = exists; UiHelp.Launch = launch; UiHelp.ShowUnavailable = notice; }
        }

        [STATestMethod]
        public void GuideControlCanBeConstructedInDesignerWithoutFileOrViewerAccess()
        {
            var context = LicenseManager.CurrentContext;
            var exists = UiHelp.Exists;
            try
            {
                LicenseManager.CurrentContext = new DesigntimeLicenseContext();
                UiHelp.Exists = path => throw new AssertFailedException("Designer consulted runtime help storage");
                using (var window = new AboutWindow())
                {
                    UiHelp.Attach(window);
                    Assert.AreEqual("helpButton", UiInvoke.Field<Button>(window, "helpButton").Name);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(UiInvoke.Field<Button>(window, "helpButton").Text));
                }
            }
            finally { LicenseManager.CurrentContext = context; UiHelp.Exists = exists; }
        }

        [STATestMethod]
        public void HelpRequestIsAttachedOnceAndUsesTheFormsOwnChapter()
        {
            var exists = UiHelp.Exists; var launch = UiHelp.Launch;
            try
            {
                int calls = 0;
                UiHelp.Exists = path => true;
                UiHelp.Launch = (owner, path, topic) => { calls++; Assert.AreEqual("privacy", topic); };
                using (var form = new ProjectAccessWindow())
                {
                    UiHelp.Attach(form); UiHelp.Attach(form);
                    var args = new HelpEventArgs(System.Drawing.Point.Empty);
                    UiInvoke.Call(typeof(Form), "OnHelpRequested", form, args);
                    Assert.IsTrue(args.Handled);
                    Assert.AreEqual(1, calls);
                }
            }
            finally { UiHelp.Exists = exists; UiHelp.Launch = launch; }
        }
    }
}

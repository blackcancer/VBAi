using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.IO;
using System.Globalization;
using System.Collections.Generic;
using System.Windows.Forms;
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
            Assert.AreEqual("VBAi." + UiLanguages.For(UiText.Culture).CultureName + ".chm", Path.GetFileName(UiHelp.FilePath));
            Assert.IsTrue(Path.IsPathRooted(UiHelp.FilePath));
        }

        /// <summary>Every UI language and regional variant resolves only inside the packaged Help folder.</summary>
        [TestMethod]
        public void SupportedCulturesResolveToTheirOwnPackagedArchives()
        {
            foreach (var language in UiLanguages.All)
            {
                string path = UiHelp.PathFor(CultureInfo.GetCultureInfo(language.CultureName));
                Assert.AreEqual("VBAi." + language.CultureName + ".chm", Path.GetFileName(path));
                Assert.AreEqual("Help", Path.GetFileName(Path.GetDirectoryName(path)));
            }
            Assert.AreEqual("VBAi.zh-TW.chm", Path.GetFileName(UiHelp.PathFor(CultureInfo.GetCultureInfo("zh-HK"))));
            Assert.AreEqual("VBAi.pt-BR.chm", Path.GetFileName(UiHelp.PathFor(CultureInfo.GetCultureInfo("pt-PT"))));
            Assert.AreEqual("VBAi.en-US.chm", Path.GetFileName(UiHelp.PathFor(CultureInfo.GetCultureInfo("nl-NL"))));
        }

        /// <summary>Missing localized files fall back offline, with no duplicate filesystem probes.</summary>
        [TestMethod]
        public void ResolutionChecksLocalizedEnglishFrenchAndReportsNoAvailableGuide()
        {
            var exists = UiHelp.Exists;
            try
            {
                var checkedPaths = new List<string>();
                var japanese = CultureInfo.GetCultureInfo("ja-JP");
                UiHelp.Exists = path => { checkedPaths.Add(Path.GetFileName(path)); return path.EndsWith("en-US.chm"); };
                Assert.AreEqual("VBAi.en-US.chm", Path.GetFileName(UiHelp.Resolve(japanese)));
                CollectionAssert.AreEqual(new[] { "VBAi.ja-JP.chm", "VBAi.en-US.chm" }, checkedPaths);
                checkedPaths.Clear();
                UiHelp.Exists = path => { checkedPaths.Add(Path.GetFileName(path)); return path.EndsWith("fr-FR.chm"); };
                Assert.AreEqual("VBAi.fr-FR.chm", Path.GetFileName(UiHelp.Resolve(japanese)));
                CollectionAssert.AreEqual(new[] { "VBAi.ja-JP.chm", "VBAi.en-US.chm", "VBAi.fr-FR.chm" }, checkedPaths);
                checkedPaths.Clear();
                UiHelp.Exists = path => { checkedPaths.Add(Path.GetFileName(path)); return false; };
                Assert.IsNull(UiHelp.Resolve(CultureInfo.GetCultureInfo("en-US")));
                CollectionAssert.AreEqual(new[] { "VBAi.en-US.chm", "VBAi.fr-FR.chm" }, checkedPaths);
                UiHelp.Exists = path => true;
                Assert.AreEqual("VBAi.ja-JP.chm", Path.GetFileName(UiHelp.Resolve(japanese)));
            }
            finally { UiHelp.Exists = exists; }
        }

        /// <summary>Language changes are evaluated for each explicit request rather than cached.</summary>
        [STATestMethod]
        public void GuideFollowsInterfaceLanguageAfterItChanges()
        {
            var exists = UiHelp.Exists; var launch = UiHelp.Launch;
            try
            {
                UiHelp.Exists = path => true;
                string opened = null;
                UiHelp.Launch = (owner, path, topic) => opened = Path.GetFileName(path);
                using (var language = new LocalizationScope("de-DE"))
                {
                    UiHelp.Open(null);
                    Assert.AreEqual("VBAi.de-DE.chm", opened);
                    LocalizationScope.Set("ar-SA");
                    UiHelp.Open(null);
                    Assert.AreEqual("VBAi.ar-SA.chm", opened);
                }
            }
            finally { UiHelp.Exists = exists; UiHelp.Launch = launch; }
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

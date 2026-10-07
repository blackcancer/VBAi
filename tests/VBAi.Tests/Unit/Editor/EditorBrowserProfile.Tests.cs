using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorBrowserProfileTests
    {
        [TestMethod]
        public void ProfilesWaitForBothRetirementAndMatchingBrowserExitAndKeepOtherState()
        {
            string root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "VBAi-profile-test-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            string legacy = System.IO.Path.Combine(root, "legacy-private.txt");
            File.WriteAllText(legacy, "Preserve unknown prior state");
            var first = new VBAi.EditorBrowserProfile(root);
            var second = new VBAi.EditorBrowserProfile(root);
            try
            {
                Assert.AreNotEqual(first.Path, second.Path);
                string nested = System.IO.Path.Combine(first.Path, "Default");
                Directory.CreateDirectory(nested);
                File.WriteAllText(System.IO.Path.Combine(nested, "cache"), "Synthetic browser data");
                first.BrowserRequested(); first.ObserveBrowser(101); first.Retire();
                Assert.IsTrue(Directory.Exists(first.Path), "Controller disposal alone does not prove browser exit.");
                first.BrowserExited(102); first.Cleanup.GetAwaiter().GetResult();
                Assert.IsTrue(Directory.Exists(first.Path), "An unrelated browser exit must not permit deletion.");
                first.BrowserExited(101); first.Cleanup.GetAwaiter().GetResult();
                Assert.IsFalse(Directory.Exists(first.Path));
                Assert.IsTrue(Directory.Exists(second.Path)); Assert.IsTrue(File.Exists(legacy));
                second.BrowserRequested(); second.ObserveBrowser(201); second.BrowserExited(201);
                Assert.IsTrue(Directory.Exists(second.Path), "A crash while the editor is live retains the profile.");
                second.ObserveBrowser(202); second.Retire();
                Assert.IsTrue(Directory.Exists(second.Path), "A new browser invalidates the earlier exit observation.");
                second.BrowserExited(202); second.Cleanup.GetAwaiter().GetResult();
                Assert.IsFalse(Directory.Exists(second.Path)); Assert.IsTrue(File.Exists(legacy));
            }
            finally
            {
                first.BrowserExited(101); first.Retire(); first.Cleanup.GetAwaiter().GetResult();
                second.BrowserExited(202); second.Retire(); second.Cleanup.GetAwaiter().GetResult();
                File.Delete(legacy); Directory.Delete(root, false);
            }
        }

        [TestMethod]
        public void CleanupDoesNotForceLockedFilesOrTreatAnUnobservedInitializationAsExit()
        {
            string root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "VBAi-profile-test-" + Guid.NewGuid().ToString("N"));
            var profile = new VBAi.EditorBrowserProfile(root);
            string file = System.IO.Path.Combine(profile.Path, "locked-cache");
            File.WriteAllText(file, "Synthetic retained file");
            profile.BrowserRequested(); profile.Retire();
            Assert.IsTrue(Directory.Exists(profile.Path));
            using (File.Open(file, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                profile.BrowserExited(303); profile.Cleanup.GetAwaiter().GetResult();
                Assert.IsTrue(File.Exists(file));
            }
            File.Delete(file); Directory.Delete(profile.Path, false); Directory.Delete(root, false);
        }
    }
}

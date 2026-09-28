using System;
using System.Diagnostics;
using System.IO;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit.Editor
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorDraftRetentionTests
    {
        [TestMethod]
        public void RetentionPreservesLatestRecoveryAndLiveProcessDrafts()
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-retention-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(root, "module"));
            try
            {
                string old = Path.Combine(root, "module", "2147483647-old.draft"), latest = Path.Combine(root, "module", "2147483647-latest.draft"), live = Path.Combine(root, "module", Process.GetCurrentProcess().Id + "-live.draft");
                foreach (string file in new[] { old, latest, live }) { File.WriteAllText(file, "fixture"); File.SetLastWriteTimeUtc(file, DateTime.UtcNow.AddDays(-60)); }
                File.SetLastWriteTimeUtc(latest, DateTime.UtcNow.AddDays(-40));
                new EditorDraftStore(root).Cleanup(DateTime.UtcNow);
                Assert.IsFalse(File.Exists(old)); Assert.IsTrue(File.Exists(latest)); Assert.IsTrue(File.Exists(live));
            }
            finally { Directory.Delete(root, true); }
        }
    }
}

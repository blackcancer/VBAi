using System.IO;
using System.Linq;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorDraftStoreTests
    {
        [TestMethod]
        public void EncryptedDraftSurvivesAnotherWindowAndCorruptionDoesNotPreventRecovery()
        {
            using (var host = new EditorFixture())
            {
                var doc = new EditorDocument(host); doc.Edit(doc.Text + "\n' private recovery text");
                var store = new EditorDraftStore(host.Root); store.Save(doc);
                string file = Directory.GetFiles(host.Root, "*.draft", SearchOption.AllDirectories).Single();
                Assert.IsFalse(File.ReadAllText(file).Contains("private recovery text"));
                var other = new EditorDraftStore(host.Root); Assert.AreEqual(doc.Text, other.Recover(host.Key).Text);
                other.ClearOwn(doc); Assert.IsTrue(File.Exists(file));
                File.WriteAllText(Path.Combine(Path.GetDirectoryName(file), "corrupt.draft"), "invalid");
                Assert.AreEqual(doc.Baseline, other.Recover(host.Key).Baseline);
                host.Closed = true; store.Save(doc); Assert.AreEqual(doc.Text, store.Recover(host.Root).Text);
                store.ClearOwn(doc); Assert.IsNull(store.Recover(host.Root));
            }
        }
    }
}

namespace CodexVBE.Tests.Unit.Editor
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
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

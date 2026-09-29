using System.IO;
using System.Linq;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
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

namespace VBAi.Tests.Unit.Editor
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using VBAi;
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
        [TestMethod]
        public void RecoveryRejectsOversizedForeignInvalidAndLockedReplaceWithoutLosingOwnedDraft()
        {
            using (var host = new VBAi.Tests.Infrastructure.EditorFixture())
            {
                var doc = new EditorDocument(host); var store = new EditorDraftStore(host.Root);
                Assert.IsTrue(new EditorDraftStore().Root.EndsWith("EditorDrafts")); Assert.IsNull(store.Recover(host.Key));
                store.Save(doc); Assert.IsFalse(Directory.Exists(host.Root));
                doc.Edit(doc.Text + "\n' owned original"); store.Save(doc);
                string directory = Directory.GetDirectories(host.Root).Single(), file = Directory.GetFiles(directory).Single();
                doc.Edit(doc.Text + "\n' update"); store.Save(doc); string preserved = doc.Text;
                using (var locked = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    doc.Edit(doc.Text + "\n' rejected replacement"); Assert.ThrowsException<IOException>(() => store.Save(doc));
                    Assert.AreEqual(0, Directory.GetFiles(directory, "*.tmp").Length);
                }
                Assert.AreEqual(preserved, store.Recover(host.Key).Text);
                string oversized = Path.Combine(directory, "oversized.draft"); using (var stream = File.Create(oversized)) stream.SetLength(16 * 1024 * 1024 + 1);
                foreach (var draft in new[] { new EditorDraft { Key = "foreign", Baseline = "a", Text = "b" }, new EditorDraft { Key = host.Key, Baseline = null, Text = "b" } })
                {
                    string id = Guid.NewGuid().ToString("N"); store.SaveSnapshot(id, draft);
                    string encrypted = Directory.GetFiles(Path.Combine(host.Root, EditorDocument.Hash(draft.Key)), "*" + id + ".draft").Single();
                    File.Move(encrypted, Path.Combine(directory, id + ".draft"));
                }
                Assert.AreEqual(preserved, store.Recover(host.Key).Text);
            }
        }

        [TestMethod]
        public void RetentionPreservesRecentMalformedLinkedLockedAndReadonlyOwnedFiles()
        {
            using (var host = new VBAi.Tests.Infrastructure.EditorFixture())
            {
                string directory = Path.Combine(host.Root, "module"); Directory.CreateDirectory(directory);
                string latest = Path.Combine(directory, "2147483647-latest.draft"), recent = Path.Combine(directory, "2147483647-recent.draft"), malformed = Path.Combine(directory, "not-a-pid.draft"), linked = Path.Combine(directory, "2147483647-linked.draft"), locked = Path.Combine(directory, "2147483647-locked.draft"), readOnly = Path.Combine(directory, "2147483647-readonly.draft");
                foreach (string file in new[] { latest, recent, malformed, linked, locked, readOnly }) { File.WriteAllText(file, "owned"); File.SetLastWriteTimeUtc(file, DateTime.UtcNow.AddDays(-60)); }
                File.SetLastWriteTimeUtc(latest, DateTime.UtcNow); File.SetLastWriteTimeUtc(recent, DateTime.UtcNow.AddDays(-1)); File.SetAttributes(readOnly, FileAttributes.ReadOnly);
                var store = new EditorDraftStore(host.Root); int linkReads = 0;
                store.ReadAttributes = item => { if (item.FullName == linked) { linkReads++; return FileAttributes.ReparsePoint; } return item.Attributes; };
                try
                {
                    using (var hold = new FileStream(locked, FileMode.Open, FileAccess.Read, FileShare.Read)) store.Cleanup(DateTime.UtcNow);
                    Assert.AreEqual(1, linkReads); foreach (string file in new[] { latest, recent, malformed, linked, locked, readOnly }) Assert.IsTrue(File.Exists(file), file);
                    store.ReadAttributes = item => item is DirectoryInfo ? FileAttributes.ReparsePoint : throw new AssertFailedException("Linked directories must not be enumerated"); store.Cleanup(DateTime.UtcNow);
                    Assert.IsTrue(File.Exists(readOnly));
                }
                finally { File.SetAttributes(readOnly, FileAttributes.Normal); }
            }
        }

        [TestMethod]
        public void RetentionDeletesAnOwnedProcessDraftAfterTheCapturedProcessExits()
        {
            using (var host = new VBAi.Tests.Infrastructure.EditorFixture())
            using (var process = Process.Start(new ProcessStartInfo { FileName = Environment.GetEnvironmentVariable("COMSPEC"), Arguments = "/d /c exit 0", UseShellExecute = false, CreateNoWindow = true }))
            {
                Assert.IsTrue(process.WaitForExit(5000)); string directory = Path.Combine(host.Root, "module"); Directory.CreateDirectory(directory);
                string old = Path.Combine(directory, process.Id + "-old.draft"), latest = Path.Combine(directory, "2147483647-latest.draft");
                File.WriteAllText(old, "owned"); File.SetLastWriteTimeUtc(old, DateTime.UtcNow.AddDays(-60)); File.WriteAllText(latest, "newest");
                var store = new EditorDraftStore(host.Root); store.ReadProcess = pid => { Assert.AreEqual(process.Id, pid); return process; }; store.Cleanup(DateTime.UtcNow);
                Assert.IsFalse(File.Exists(old)); Assert.IsTrue(File.Exists(latest));
            }
        }
        [TestMethod]
        public void ActualOwnedJunctionRootsAndChildrenAreNeverFollowedByRetention()
        {
            using (var host = new VBAi.Tests.Infrastructure.EditorFixture())
            {
                string target = Path.Combine(host.Root, "target"), junction = Path.Combine(host.Root, "linked"); Directory.CreateDirectory(target);
                string old = Path.Combine(target, "2147483647-old.draft"), latest = Path.Combine(target, "2147483647-latest.draft");
                File.WriteAllText(old, "owned"); File.SetLastWriteTimeUtc(old, DateTime.UtcNow.AddDays(-60)); File.WriteAllText(latest, "newest");
                VBAi.Tests.Infrastructure.OwnedDraftJunction.Create(junction, target);
                try
                {
                    Assert.IsTrue((File.GetAttributes(junction) & FileAttributes.ReparsePoint) != 0);
                    new EditorDraftStore(junction).Cleanup(DateTime.UtcNow); Assert.IsTrue(File.Exists(old));
                    var store = new EditorDraftStore(host.Root); store.ReadAttributes = item => item.FullName == target ? FileAttributes.ReparsePoint : item.Attributes;
                    store.Cleanup(DateTime.UtcNow); Assert.IsTrue(File.Exists(old)); Assert.IsTrue(File.Exists(latest));
                }
                finally { Directory.Delete(junction); }
            }
        }
    }
}

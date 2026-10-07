using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorDocumentTests
    {
        [TestMethod]
        public void ConcurrentNativeChangeNeverOverwritesEitherRevision()
        {
            using (var host = new EditorFixture())
            {
                var doc = new EditorDocument(host); string baseline = host.Code;
                doc.Edit(baseline.Replace("Print 1", "Print 2")); host.Code = baseline.Replace("Print 1", "Print 3");
                Assert.ThrowsException<InvalidOperationException>(() => doc.Synchronize());
                Assert.IsTrue(doc.Conflict); Assert.AreEqual(0, host.Writes);
                StringAssert.Contains(doc.Text, "Print 2"); StringAssert.Contains(host.Code, "Print 3");
                Assert.AreEqual(baseline, doc.Baseline);
            }
        }
        [TestMethod]
        public void RunModeAndFailedWritesPreserveDirtyDrafts()
        {
            using (var host = new EditorFixture())
            {
                var doc = new EditorDocument(host); string original = host.Code; doc.Edit(original + "\n' draft");
                host.CanWrite = false; Assert.ThrowsException<InvalidOperationException>(() => doc.Synchronize());
                host.CanWrite = true; host.Fail = true; Assert.ThrowsException<InvalidOperationException>(() => doc.Synchronize());
                Assert.AreEqual(original, host.Code); Assert.IsTrue(doc.Dirty);
                host.Fail = false; string result = doc.Synchronize(); doc.Acknowledge(result, doc.Text);
                Assert.IsFalse(doc.Dirty); Assert.AreEqual(1, host.Writes);
            }
        }
        [TestMethod]
        public void ExternalRefreshIsOnlyAcceptedAfterRendererAcknowledgement()
        {
            using (var host = new EditorFixture())
            {
                var doc = new EditorDocument(host); string original = host.Code; host.Code += "\n' external";
                Assert.AreEqual(host.Code, doc.Observe()); Assert.AreEqual(original, doc.Text);
                doc.Edit(original + "\n' typing concurrently"); Assert.IsNull(doc.Observe()); Assert.IsTrue(doc.Conflict);
            }
        }
        [TestMethod]
        public void LateAcknowledgementDoesNotDiscardNewerTyping()
        {
            using (var host = new EditorFixture())
            {
                var doc = new EditorDocument(host); string captured = doc.Text;
                doc.Edit(captured + "\n' newer"); doc.Acknowledge(captured, captured);
                StringAssert.EndsWith(doc.Text, "' newer"); Assert.IsTrue(doc.Dirty);
            }
        }
        [TestMethod]
        public void ConflictResolutionRequiresTheExactReviewedNativeRevision()
        {
            using (var host = new EditorFixture())
            {
                var doc = new EditorDocument(host); doc.Edit(doc.Text + "\n' local");
                host.Code += "\n' external"; doc.Observe(); string reviewed = doc.Native;
                host.Code += "\n' newer";
                Assert.ThrowsException<InvalidOperationException>(() => doc.ResolveWithDraft(reviewed));
                Assert.AreEqual(0, host.Writes);
                doc.Observe(); string captured = doc.Text;
                string result = doc.ResolveWithDraft(doc.Native); doc.Acknowledge(result, captured);
                Assert.AreEqual(captured, host.Code); Assert.IsFalse(doc.Conflict); Assert.IsFalse(doc.Dirty);
            }
        }
        [DataTestMethod]
        [DataRow("a\nb\nc", "a\nx\nc", 2, 1, "x")]
        [DataRow("a\nb", "a\nb\nc", 3, 0, "c")]
        [DataRow("a\nb\nc", "a\nc", 2, 1, "")]
        [DataRow("", "a", 1, 0, "a")]
        [DataRow("a", "", 1, 1, "")]
        public void PatchesKeepUnchangedProcedureLines(string before, string after, int start, int count, string replacement)
        { var patch = EditorDocument.Difference(before, after); Assert.AreEqual(start, patch.Item1); Assert.AreEqual(count, patch.Item2); Assert.AreEqual(replacement, patch.Item3); }
        [TestMethod]
        public void InvalidTextIsRejectedBeforeItCanReachVba()
        {
            Assert.ThrowsException<InvalidOperationException>(() => EditorDocument.Validate("a\0b"));
            Assert.ThrowsException<InvalidOperationException>(() => EditorDocument.Validate(new string('a', EditorDocument.MaxLength + 1)));
        }
        [TestMethod]
        public void CleanEquivalentAndStalePlannedDraftsKeepTheirExactNativeRevision()
        {
            using (var host = new EditorFixture())
            {
                var doc = new EditorDocument(host); Assert.AreEqual(doc.Text, doc.Synchronize()); Assert.AreEqual(0, host.Writes);
                doc.Edit(doc.Text + "\n' already native"); host.Code = doc.Text; Assert.AreEqual(doc.Text, doc.Synchronize()); Assert.IsFalse(doc.Dirty); Assert.AreEqual(0, host.Writes);
                string baseline = doc.Baseline; doc.Edit(doc.Text + "\n' new draft");
                Assert.ThrowsException<InvalidOperationException>(() => doc.Synchronize(new EditorSyncPlan("stale", doc.Text)));
                Assert.ThrowsException<InvalidOperationException>(() => doc.Synchronize(new EditorSyncPlan(baseline, "stale")));
                host.CanWrite = false; Assert.ThrowsException<InvalidOperationException>(() => doc.ResolveWithDraft(host.Code)); Assert.AreEqual(0, host.Writes);
            }
            using (var host = new EditorFixture())
            { var doc = new EditorDocument(host); doc.Restore(doc.Baseline, doc.Text + "\n\'recovery"); Assert.IsTrue(doc.Dirty); doc.AcceptRemote(host.Code); Assert.IsFalse(doc.Dirty); }
            using (var f = new ModernEditorDebugFixture())
            { f.Document.Edit(f.Document.Text + "\n\'native draft"); string expected = f.Document.Text; Assert.AreEqual(expected, f.Document.Synchronize(new EditorSyncPlan(f.Document.Baseline, expected))); Assert.AreEqual(expected, EditorDocument.Normalize(f.Native.Adapter.Read())); }
            Assert.AreEqual("", EditorDocument.Normalize(null)); Assert.ThrowsException<InvalidOperationException>(() => EditorDocument.Validate(null));
            var equal = EditorDocument.Difference("same", "same"); Assert.AreEqual(0, equal.Item2); Assert.AreEqual("", equal.Item3);
        }
    }
}

using System;
using System.Reflection;
using System.Threading.Tasks;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit.Editor
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorTextChangeTests
    {
        [TestMethod]
        public void Utf16MulticursorEditsAreAtomicAndMalformedBatchesCannotPartiallyChangeText()
        {
            const string text = "a😀b\nsecond";
            Assert.IsTrue(EditorTextChange.TryApply(text, new[] {
                new EditorTextChange { rangeOffset = 5, rangeLength = 6, text = "line" },
                new EditorTextChange { rangeOffset = 1, rangeLength = 2, text = "é" }
            }, out var result));
            Assert.AreEqual("aéb\nline", result);
            foreach (var bad in new[] {
                new[] { new EditorTextChange { rangeOffset = -1, text = "x" } },
                new[] { new EditorTextChange { rangeOffset = 99, text = "x" } },
                new[] { new EditorTextChange { rangeOffset = 0, rangeLength = int.MaxValue, text = "x" } },
                new[] { new EditorTextChange { rangeOffset = 0, rangeLength = 4, text = "x" }, new EditorTextChange { rangeOffset = 2, rangeLength = 1, text = "y" } },
                new[] { new EditorTextChange { text = "\0" } }, new EditorTextChange[] { null }
            }) { Assert.IsFalse(EditorTextChange.TryApply(text, bad, out result)); Assert.AreEqual(text, result); }
        }
        private static object Message(ModernEditorToolFixture f, int baseVersion, int version, string text)
        {
            var type = typeof(ModernEditorWindow).GetNestedType("EditorMessage", BindingFlags.NonPublic);
            return f.Json.Deserialize(f.Json.Serialize(new { id = f.Document.Id, baseVersion, version,
                changes = new[] { new { rangeOffset = 0, rangeLength = f.Document.Text.Length, text } } }), type);
        }
        private static void Call(ModernEditorToolFixture f, string method, params object[] values)
            => ModernEditorDebugFixture.Wait((Task)f.Private(method, values));
        [STATestMethod]
        public void StreamFlushWritesOnlyDirtyDocumentAndRecoversAGapFromTheRenderer()
        {
            using (var f = new ModernEditorToolFixture())
            {
                var expected = f.Document.Text + "\n' streamed";
                int captures = f.Captures;
                Call(f, "AcceptEditorChange", Message(f, 1, 2, expected));
                Assert.AreEqual(expected, f.Document.Text); Assert.AreEqual(captures, f.Captures);
                f.Base.Set("lastEdit", DateTime.UtcNow.AddSeconds(-1));
                Call(f, "FlushStream");
                Assert.AreEqual(expected, f.Module.Code); Assert.AreEqual(1, f.Module.Writes);
                Assert.AreEqual(captures, f.Captures); Assert.IsFalse(f.Document.Dirty);
                var recovered = expected + "\n' recovered latest";
                f.Override = (method, args) => method == "snapshots" ? f.Json.Serialize(new[] { f.Snapshot(recovered, 8) }) : null;
                Call(f, "AcceptEditorChange", Message(f, 5, 7, "must never apply"));
                Assert.AreEqual(recovered, f.Document.Text); Assert.AreEqual(8, f.Versions[f.Document.Id]);
                Assert.AreEqual(captures + 1, f.Captures);
                Call(f, "AcceptEditorChange", Message(f, 5, 7, "stale"));
                Assert.AreEqual(recovered, f.Document.Text);
            }
        }
        [STATestMethod]
        public void UnwritableStreamKeepsRecoveryAndResumesWithoutOverwritingNativeConflicts()
        {
            using (var f = new ModernEditorToolFixture())
            {
                f.Module.CanWrite = false;
                string edit = f.Document.Text + "\n' preserve";
                Call(f, "AcceptEditorChange", Message(f, 1, 2, edit));
                f.Base.Set("lastEdit", DateTime.UtcNow.AddSeconds(-1)); Call(f, "FlushStream");
                Assert.AreEqual(0, f.Module.Writes); Assert.IsTrue(f.Document.Dirty);
                Assert.AreEqual(edit, f.Window.Drafts.Recover(f.Module.Key).Text);
                f.Module.CanWrite = true;
                ModernEditorDebugFixture.Wait(f.Window.ProcessDocuments(true));
                Assert.AreEqual(edit, f.Module.Code); Assert.IsFalse(f.Document.Dirty);
                Call(f, "AcceptEditorChange", Message(f, f.Versions[f.Document.Id], f.Versions[f.Document.Id] + 1, edit + "\n' second"));
                f.Module.Code += "\n' native edit";
                f.Base.Set("lastEdit", DateTime.UtcNow.AddSeconds(-1)); Call(f, "FlushStream");
                Assert.IsTrue(f.Document.Conflict); Assert.AreEqual(1, f.Module.Writes);
            }
        }
    }
}

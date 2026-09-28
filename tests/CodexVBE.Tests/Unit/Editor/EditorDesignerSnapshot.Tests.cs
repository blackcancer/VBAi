using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit.Editor
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorDesignerSnapshotTests
    {
        private static object Property(string name, object value, string kind = "scalar", string digest = null, string error = null) => new { Name = name, Value = value, Kind = kind, Digest = digest, Members = (object)null, Error = error, Display = (string)null };
        private static object Tree(string name, string caption, string picture) => new { Properties = new[] { Property("Name", name) }, Controls = new[] { new { Path = "Controls/Button", Type = "CommandButton", Properties = new[] { Property("Name", "Button"), Property("Caption", caption), Property("Picture", null, "object", picture) }, Children = new object[0] } } };
        [TestMethod]
        public void TemporaryFormNameIsIgnoredButCaptionAndPictureChangesAreDetected()
        {
            string before = EditorDesignerSnapshot.Capture(Tree("Original", "Caption", "image-a"));
            Assert.AreEqual(before, EditorDesignerSnapshot.Capture(Tree("Staged", "Caption", "image-a")));
            Assert.AreNotEqual(before, EditorDesignerSnapshot.Capture(Tree("Staged", "Changed", "image-a")));
            Assert.AreNotEqual(before, EditorDesignerSnapshot.Capture(Tree("Staged", "Caption", "image-b")));
        }
        [TestMethod]
        public void OpaquePropertiesAndUnknownControlsPreventReplacement()
        {
            Assert.ThrowsException<InvalidOperationException>(() => EditorDesignerSnapshot.Capture(new { Properties = new[] { Property("CustomObject", null, "object") }, Controls = new object[0] }));
            Assert.ThrowsException<InvalidOperationException>(() => EditorDesignerSnapshot.Capture(new { Properties = new object[0], Controls = new[] { new { Path = "Controls/Custom", Type = "ThirdPartyControl", Properties = new object[0], Children = new object[0] } } }));
            Assert.ThrowsException<InvalidOperationException>(() => EditorDesignerSnapshot.Capture(new { Properties = new[] { Property("Caption", null, error: "Read failed") }, Controls = new object[0] }));
        }
    }
}

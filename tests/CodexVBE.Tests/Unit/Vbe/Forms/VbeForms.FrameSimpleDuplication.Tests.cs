namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeFormsValueDuplicationTests
    {
        [TestMethod]
        public void SimpleFrameCopiesLabelAndTextBoxValueWithDistinctPaths()
        {
            var f = Create("Frame");
            var label = f.Source.Controls.AddExisting("Label", "Title");
            label.Caption = "Heading";
            var text = f.Source.Controls.AddExisting("TextBox", "Entry");
            text.Value = "payload";
            var request = f.Request(name: "FrameCopy");
            dynamic result = f.Service.DuplicateFrameWithSimpleChildren(request);
            var copy = f.Form.Designer.Controls.Item("FrameCopy");
            Assert.AreEqual("Heading", copy.Controls.Item("FrameCopy_Title").Caption);
            Assert.AreEqual("payload", copy.Controls.Item("FrameCopy_Entry").Value);
            Assert.AreEqual(1, (int)result.DirectLabelsCopied);
            Assert.AreEqual(1, (int)result.DirectTextBoxesCopied);
            Assert.AreEqual(2, ((IEnumerable)result.CopiedChildPaths).Cast<string>().Count());
        }

        [TestMethod]
        public void SimpleFrameRefusesNoTextBoxOrNonTextValueBeforeMutation()
        {
            var f = Create("Frame");
            f.Source.Controls.AddExisting("Label", "Title");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameWithSimpleChildren(f.Request(name: "FrameCopy")));
            var text = f.Source.Controls.AddExisting("TextBox", "Entry");
            text.Value = 42;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameWithSimpleChildren(f.Request(name: "FrameCopy")));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void SimpleFrameRemovesCreatedChildAndFrameAfterChildSetterFailure()
        {
            var f = Create("Frame");
            f.Source.Controls.AddExisting("Label", "Title");
            f.Source.Controls.AddExisting("TextBox", "Entry").Value = "payload";
            f.Form.Designer.Controls.FailNextChildCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameWithSimpleChildren(f.Request(name: "FrameCopy")));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
        }
    }
}

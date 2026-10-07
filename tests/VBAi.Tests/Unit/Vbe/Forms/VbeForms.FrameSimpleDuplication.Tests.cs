namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections;
    using System.Linq;

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
namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    public sealed partial class VbeFormsFrameDuplicationTests
    {
        [TestMethod] public void SimpleFrameChecksPreflightGeometryFontsAndConcurrentTreeChanges() { FramePreflightFailures("simple"); }
        [TestMethod] public void SimpleFrameChecksEverySupportedReadbackAndHierarchyField() { FrameReadbackFailures("simple"); }
        [TestMethod] public void SimpleFrameReportsNativeFailuresAndIncompleteChildOrRootRollback() { FrameNativeAndRollbackFailures("simple"); }
        [TestMethod]
        public void SimpleFrameRejectsIneligiblePlanAndPreservesNullTextBoxValues()
        {
            var f = Create(); f.Frame.Controls.AddExisting("CheckBox", "Check");
            var error = Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameWithSimpleChildren(f.Request()));
            StringAssert.Contains(error.Message, "plan is ineligible");
            var nullable = FrameWithChildren("simple"); nullable.Frame.Controls.Item("Entry").Value = null;
            dynamic result = nullable.Service.DuplicateFrameWithSimpleChildren(nullable.Request());
            Assert.AreEqual(1, (int)result.DirectTextBoxesCopied);
            Assert.IsNull(nullable.Form.Designer.Controls.Item("FrameCopy").Controls.Item("FrameCopy_Entry").Value);
            var mismatch = FrameWithChildren("simple"); var request = mismatch.Request();
            mismatch.Form.Designer.Controls.ConfigureAdded = frame => frame.Controls.ConfigureAdded = c => { if (c.Name.EndsWith("_Entry", StringComparison.Ordinal)) c.ReadOverrides["Value"] = "changed"; };
            AssertRolledBack(mismatch, request, "simple", "Copied TextBox value differs");
        }
    }
}

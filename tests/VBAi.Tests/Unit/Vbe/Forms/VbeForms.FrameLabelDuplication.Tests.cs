namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;

    public sealed partial class VbeFormsFrameDuplicationTests
    {
        [TestMethod]
        public void LabelFrameDuplicationCopiesChildrenAndSourceRemainsUntouched()
        {
            var f = Create();
            var label = f.Frame.Controls.AddExisting("Label", "Title");
            label.Caption = "Before";
            label.Left = 2;
            label.Top = 3;
            label.Width = 50;
            label.Height = 12;
            label.BackColor = 255;
            label.Font.Name = "Calibri";
            label.Font.Size = 11;
            label.Font.Bold = true;
            var request = f.Request();
            dynamic result = f.Service.DuplicateFrameWithLabels(request);
            var copy = f.Form.Designer.Controls.Item("FrameCopy");
            var child = copy.Controls.Item("FrameCopy_Title");
            Assert.AreEqual("Controls/FrameCopy", (string)result.NewPath);
            Assert.AreEqual(1, (int)result.DirectLabelsCopied);
            Assert.AreEqual("Before", child.Caption);
            Assert.AreEqual(2d, child.Left);
            Assert.AreEqual(255, child.BackColor);
            Assert.AreEqual("Calibri", child.Font.Name);
            Assert.IsTrue(child.Font.Bold);
            Assert.AreEqual("Title", label.Name);
            Assert.AreNotEqual((string)request.ExpectedTreeVersion, (string)result.Tree.TreeVersion);
        }

        [TestMethod]
        public void LabelFrameDuplicationRejectsEmptyOrUnsupportedChildrenBeforeAdd()
        {
            var f = Create();
            var request = f.Request();
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameWithLabels(request));
            f.Frame.Controls.AddExisting("TextBox", "Entry");
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameWithLabels(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void LabelFrameRollsBackWhenFrameSetterFails()
        {
            var f = Create();
            f.Frame.Controls.AddExisting("Label", "Title");
            var request = f.Request();
            f.Form.Designer.Controls.FailNextCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameWithLabels(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
            Assert.AreEqual((string)request.ExpectedTreeVersion, (string)((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion);
        }
    }
}
namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Drawing;
    using System.Reflection;
    public sealed partial class VbeFormsFrameDuplicationTests
    {
        [TestMethod] public void LabelFrameChecksPreflightGeometryFontsAndConcurrentTreeChanges() { FramePreflightFailures("labels"); }
        [TestMethod] public void LabelFrameChecksEverySupportedReadbackAndHierarchyField() { FrameReadbackFailures("labels"); }
        [TestMethod] public void LabelFrameReportsNativeFailuresAndIncompleteChildOrRootRollback() { FrameNativeAndRollbackFailures("labels"); }
        [TestMethod]
        public void LabelColorConverterAcceptsNativeOleIntegersAndManagedColors()
        {
            var method = typeof(VBAi.VbeForms).GetMethod("CopyOleColor", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.AreEqual(ColorTranslator.ToOle(Color.Red), method.Invoke(null, new object[] { Color.Red }));
            Assert.AreEqual(255, method.Invoke(null, new object[] { 255 }));
        }
    }
}

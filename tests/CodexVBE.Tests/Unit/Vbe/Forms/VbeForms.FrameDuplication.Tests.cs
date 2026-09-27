namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeFormsFrameDuplicationTests
    {
        [TestMethod]
        public void EmptyFrameDuplicationCopiesSupportedPropertiesAndChangesVersion()
        {
            var f = Create();
            var request = f.Request();
            dynamic result = f.Service.DuplicateEmptyFrame(request);
            var copy = f.Form.Designer.Controls.Item("FrameCopy");
            Assert.AreEqual("Controls/FrameCopy", (string)result.NewPath);
            Assert.AreEqual("Group", copy.Caption);
            Assert.AreEqual(4d, copy.Left);
            Assert.AreEqual(6d, copy.Top);
            Assert.AreEqual(90d, copy.Width);
            Assert.AreEqual(60d, copy.Height);
            Assert.AreEqual(0, copy.Controls.Count);
            Assert.AreEqual(0, (int)result.SourceChildCount);
            Assert.AreNotEqual((string)request.ExpectedTreeVersion, (string)result.Tree.TreeVersion);
        }

        [TestMethod]
        public void EmptyFrameDuplicationRefusesChildrenStaleVersionAndGeometry()
        {
            var f = Create();
            var request = f.Request();
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateEmptyFrame(request));
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            f.Frame.Controls.AddExisting("Label", "Title");
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateEmptyFrame(request));
            f.Frame.Controls.Remove("Title");
            f.Frame.Width = 0;
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateEmptyFrame(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void EmptyFrameRollsBackWhenCopiedCaptionSetterFails()
        {
            var f = Create();
            var request = f.Request();
            f.Form.Designer.Controls.FailNextCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateEmptyFrame(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
            Assert.AreEqual((string)request.ExpectedTreeVersion, (string)((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion);
        }
    }
}

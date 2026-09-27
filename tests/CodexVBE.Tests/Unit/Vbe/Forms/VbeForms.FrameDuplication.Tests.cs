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
namespace CodexVBE.Tests.Unit
{
    using System;
    using System.ComponentModel;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    public sealed partial class VbeFormsFrameDuplicationTests
    {
        [TestMethod]
        public void EmptyFrameRejectsRequiredFieldsNoncanonicalPathWrongTypeAndCollision()
        {
            var f=Create();
            foreach(var field in new[]{"ControlPath","ExpectedTreeVersion"}) {
                var r=f.Request();typeof(CodexVBE.Request).GetProperty(field).SetValue(r,null);
                Assert.ThrowsException<ArgumentException>(()=>f.Service.DuplicateEmptyFrame(r));
            }
            var missing=f.Request();missing.ControlPath="Controls/missing";
            Assert.ThrowsException<InvalidOperationException>(()=>f.Service.DuplicateEmptyFrame(missing));
            f.Form.Designer.Controls.AddExisting("Label","Title");
            var wrong=f.Request();wrong.ControlPath="Controls/Title";
            Assert.ThrowsException<InvalidOperationException>(()=>f.Service.DuplicateEmptyFrame(wrong));
            Assert.ThrowsException<InvalidOperationException>(()=>f.Service.DuplicateEmptyFrame(f.Request("Title")));
            f.Frame.Pages.Add("P","Page");var page=f.Request();page.ControlPath="Controls/Frame1/Pages/P";
            Assert.ThrowsException<ArgumentException>(()=>f.Service.DuplicateEmptyFrame(page));
        }

        [TestMethod]
        public void EmptyFrameRejectsEveryGeometryBoundaryBeforeNativeAdd()
        {
            foreach(var box in InvalidBoxes) {
                var f=Create();ChangeGeometry(f.Frame,(string)box[0],(double)box[1]);
                var error=Assert.ThrowsException<InvalidOperationException>(()=>f.Service.DuplicateEmptyFrame(f.Request()));
                StringAssert.Contains(error.Message,"geometry");
                Assert.AreEqual(1,f.Form.Designer.Controls.Count);
            }
        }

        [TestMethod]
        public void EmptyFrameCopiesInsideNestedFrameAndRejectsMissingParentDescriptor()
        {
            var f=Create();f.Frame.Controls.AddExisting("Frame","Nested");
            var r=f.Request();r.ControlPath="Controls/Frame1/Controls/Nested";
            dynamic result=f.Service.DuplicateEmptyFrame(r);
            Assert.AreEqual("Controls/Frame1/Controls/FrameCopy",(string)result.NewPath);
            Assert.AreEqual(2,f.Frame.Controls.Count);
            var root=Create();var rootRequest=root.Request();var provider=new HideControlsProvider(root.Form.Designer);
            TypeDescriptor.AddProvider(provider,root.Form.Designer);
            try { var error=Assert.ThrowsException<InvalidOperationException>(()=>root.Service.DuplicateEmptyFrame(rootRequest));StringAssert.Contains(error.Message,"no Controls collection"); }
            finally { TypeDescriptor.RemoveProvider(provider,root.Form.Designer); }
        }

        [TestMethod]
        public void EmptyFrameDetectsEverySupportedReadbackMismatchAndRollsBack()
        {
            foreach(var field in new[]{"Caption","Children","Left","Top","Width","Height"}) {
                var f=Create();var r=f.Request();
                f.Form.Designer.Controls.ConfigureAdded=c=>{
                    if(field=="Caption") c.ReadOverrides["Caption"]="changed";
                    else if(field=="Children") c.Controls.CountOverride=1;
                    else c.ReadOverrides[field]=100d;
                };
                AssertRolledBack(f,r,"empty","did not retain");
            }
            var hidden=Create();var request=hidden.Request();hidden.Form.Designer.Controls.HideAdded=true;
            var error=Assert.ThrowsException<InvalidOperationException>(()=>hidden.Service.DuplicateEmptyFrame(request)); StringAssert.Contains(error.Message,"rollback also failed"); Assert.AreEqual(1,hidden.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void EmptyFrameReportsNativeAddAndRollbackFailures()
        {
            var f=Create();var r=f.Request();f.Form.Designer.Controls.FailAdd=true;
            AssertRolledBack(f,r,"empty","Native Add failed");
            Assert.AreEqual(0,f.Form.Designer.Controls.RemoveCount);
            f=Create();r=f.Request();f.Form.Designer.Controls.FailNextCaption=true;f.Form.Designer.Controls.FailNextRemove=true;
            var error=Assert.ThrowsException<InvalidOperationException>(()=>f.Service.DuplicateEmptyFrame(r));
            StringAssert.Contains(error.Message,"rollback also failed");
            Assert.AreEqual(2,f.Form.Designer.Controls.Count);
        }
    }
}

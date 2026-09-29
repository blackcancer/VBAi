namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Linq;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeFormsValueDuplicationTests
    {
        [TestMethod]
        public void OptionButtonCopiesShellButLeavesSelectionAndGroupOut()
        {
            var f = Create("OptionButton");
            f.Source.Value = true;
            dynamic result = f.Service.DuplicateOptionButton(f.Request());
            var copy = f.Form.Designer.Controls.Item("Copy");
            Assert.AreEqual("Source", copy.Caption);
            Assert.AreEqual(false, copy.Value);
            Assert.IsFalse((bool)result.SelectionCopied);
            Assert.IsFalse((bool)result.GroupCopied);
            Assert.AreEqual(true, f.Source.Value);
        }

        [TestMethod]
        public void OptionButtonRejectsWrongTypeAndRollsBackSetterFailure()
        {
            var f = Create("OptionButton");
            var request = f.Request();
            request.ControlPath = "Controls/Missing";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateOptionButton(request));
            f.Form.Designer.Controls.AddExisting("Label", "Label1");
            request.ControlPath = "Controls/Label1";
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateOptionButton(request));
            request.ControlPath = "Controls/OptionButton1";
            f.Form.Designer.Controls.FailNextCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateOptionButton(request));
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
        }
    }
}

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeFormsValueDuplicationTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void OptionButtonCompleteGuardAndRollbackMatrix()
        {
            VerifyDuplicationGuards("OptionButton");
        }
    }
}
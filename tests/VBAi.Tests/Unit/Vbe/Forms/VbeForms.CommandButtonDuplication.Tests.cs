namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeFormsFrameDuplicationTests
    {
        [TestMethod]
        public void CommandButtonDuplicationCopiesNarrowProfileAndLeavesEventsOut()
        {
            var f = Create();
            var button = f.Form.Designer.Controls.AddExisting("CommandButton", "RunButton");
            button.Caption = "Run";
            button.Left = 7;
            button.Top = 9;
            button.Width = 45;
            button.Height = 18;
            var request = f.Request("RunCopy");
            request.ControlPath = "Controls/RunButton";
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            dynamic result = f.Service.DuplicateCommandButton(request);
            var copy = f.Form.Designer.Controls.Item("RunCopy");
            Assert.AreEqual("Run", copy.Caption);
            Assert.AreEqual(7d, copy.Left);
            Assert.AreEqual(18d, copy.Height);
            Assert.IsFalse((bool)result.EventsCopied);
            Assert.AreEqual("Partial", (string)result.Completeness);
        }

        [TestMethod]
        public void CommandButtonDuplicationRefusesWrongTypeAndRollsBackSetterFailure()
        {
            var f = Create();
            var request = f.Request("RunCopy");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateCommandButton(request));
            f.Form.Designer.Controls.AddExisting("CommandButton", "RunButton");
            request.ControlPath = "Controls/RunButton";
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            f.Form.Designer.Controls.FailNextCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateCommandButton(request));
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
        }
    }
}

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeFormsValueDuplicationTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void CommandButtonCompleteGuardAndRollbackMatrix()
        {
            VerifyDuplicationGuards("CommandButton");
        }
    }
}
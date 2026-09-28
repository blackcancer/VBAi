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
        public void ToggleButtonCopiesFalseDefaultWithoutWritingSelection()
        {
            var f = Create("ToggleButton");
            var request = f.Request();
            dynamic result = f.Service.DuplicateToggleButton(request);
            Assert.AreEqual(false, f.Form.Designer.Controls.Item("Copy").Value);
            Assert.AreEqual("Source", f.Form.Designer.Controls.Item("Copy").Caption);
            Assert.IsTrue(((string[])result.CopiedProperties).Last().Contains("no setter call"));
            f.Source.Value = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateToggleButton(f.Request(name: "BadCopy")));
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void ToggleButtonRollsBackFailedCaption()
        {
            var f = Create("ToggleButton");
            f.Form.Designer.Controls.FailNextCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateToggleButton(f.Request()));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    public sealed partial class VbeFormsValueDuplicationTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void ToggleButtonCompleteGuardAndRollbackMatrix()
        {
            VerifyDuplicationGuards("ToggleButton");
        }
    }
}
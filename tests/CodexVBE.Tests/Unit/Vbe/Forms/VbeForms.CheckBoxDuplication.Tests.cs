namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeFormsValueDuplicationTests
    {
        [TestMethod]
        public void CheckBoxCopiesBooleanTrueAndRefusesTriState()
        {
            var f = Create("CheckBox");
            f.Source.Value = true;
            var request = f.Request();
            dynamic result = f.Service.DuplicateCheckBox(request);
            var copy = f.Form.Designer.Controls.Item("Copy");
            Assert.AreEqual(true, copy.Value);
            Assert.AreEqual("Source", copy.Caption);
            Assert.AreEqual("Partial", (string)result.Completeness);
            f.Source.Value = null;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateCheckBox(f.Request(name: "BadCopy")));
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void CheckBoxRollsBackWhenBooleanSetterFails()
        {
            var f = Create("CheckBox");
            f.Source.Value = true;
            f.Form.Designer.Controls.FailNextValue = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateCheckBox(f.Request()));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
        }
    }
}

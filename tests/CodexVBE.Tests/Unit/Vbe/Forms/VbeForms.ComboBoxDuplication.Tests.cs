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
        public void ComboBoxCopiesListWidthButNotItemsOrBindings()
        {
            var f = Create("ComboBox");
            f.Source.ListWidth = "83 pt";
            f.Source.Items.Add("selection");
            f.Source.RowSource = "Sheet1!A1:A2";
            dynamic result = f.Service.DuplicateComboBox(f.Request());
            var copy = f.Form.Designer.Controls.Item("Copy");
            Assert.AreEqual("83 pt", copy.ListWidth);
            Assert.AreEqual(0, copy.Items.Count);
            Assert.AreEqual("", copy.RowSource);
            Assert.IsFalse((bool)result.ItemsCopied);
            Assert.IsFalse((bool)result.BindingsCopied);
        }

        [TestMethod]
        public void ComboBoxRejectsNonTextWidthAndRollsBackSetterFailure()
        {
            var f = Create("ComboBox");
            f.Source.ListWidth = 83;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateComboBox(f.Request()));
            f.Source.ListWidth = "83 pt";
            f.Form.Designer.Controls.FailNextListWidth = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateComboBox(f.Request()));
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
        public void ComboBoxCompleteGuardAndRollbackMatrix()
        {
            VerifyDuplicationGuards("ComboBox");
        }
    }
}
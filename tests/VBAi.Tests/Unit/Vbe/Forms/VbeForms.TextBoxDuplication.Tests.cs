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
        public void TextBoxCopiesTextAndGeometryWithoutChangingTheSource()
        {
            var f = Create("TextBox");
            f.Source.Value = "alpha";
            var request = f.Request();
            dynamic result = f.Service.DuplicateTextBox(request);
            var copy = f.Form.Designer.Controls.Item("Copy");
            Assert.AreEqual("alpha", copy.Value);
            Assert.AreEqual(3d, copy.Left);
            Assert.AreEqual(50d, copy.Width);
            Assert.AreEqual("alpha", f.Source.Value);
            Assert.AreNotEqual((string)request.ExpectedTreeVersion, (string)result.Tree.TreeVersion);
        }

        [TestMethod]
        public void TextBoxAcceptsEmptyValueButRejectsNonTextBeforeAdd()
        {
            var f = Create("TextBox");
            dynamic empty = f.Service.DuplicateTextBox(f.Request());
            Assert.IsNull(f.Form.Designer.Controls.Item("Copy").Value);
            f.Source.Value = 12;
            var request = f.Request(name: "BadCopy");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateTextBox(request));
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
            Assert.AreEqual("Controls/Copy", (string)empty.NewPath);
        }

        [TestMethod]
        public void TextBoxRejectsStaleTreeAndRollsBackFailedValueSetter()
        {
            var f = Create("TextBox");
            f.Source.Value = "alpha";
            var request = f.Request();
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateTextBox(request));
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            f.Form.Designer.Controls.FailNextValue = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateTextBox(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
        }
    }
}

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeFormsValueDuplicationTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void TextBoxCompleteGuardAndRollbackMatrix()
        {
            VerifyDuplicationGuards("TextBox");
        }
    }
}
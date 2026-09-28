using System;

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Reflection;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeFormsPartialTests
    {
        [TestMethod]
        public void DuplicateLabelCopiesSupportedPropertiesAndChangesTree()
        {
            var f = Create();
            f.Control.Caption = "Source";
            f.Control.Left = 4;
            f.Control.Top = 7;
            f.Control.Width = 48;
            f.Control.Height = 19;
            f.Control.BackColor = 255;
            f.Control.Font.Name = "Calibri";
            f.Control.Font.Size = 11;
            f.Control.Font.Bold = true;
            dynamic before = f.Service.Tree(f.Project.Name, f.Form.Name);
            dynamic result = f.Service.DuplicateLabel(new Request { Project = f.Project.Name, Form = f.Form.Name, ControlPath = "Controls/Label1", NewName = "LabelCopy", ExpectedTreeVersion = before.TreeVersion });
            var copy = f.Form.Designer.Controls.Item("LabelCopy");
            Assert.AreEqual("Controls/LabelCopy", (string)result.NewPath);
            Assert.AreEqual("Source", copy.Caption);
            Assert.AreEqual(4d, copy.Left);
            Assert.AreEqual(48d, copy.Width);
            Assert.AreEqual(255, copy.BackColor);
            Assert.AreEqual("Calibri", copy.Font.Name);
            Assert.IsTrue(copy.Font.Bold);
            Assert.AreNotEqual((string)before.TreeVersion, (string)result.Tree.TreeVersion);
        }

        [TestMethod]
        public void DuplicateLabelRollsBackNewControlWhenNativeSetterFails()
        {
            var f = Create();
            f.Form.Designer.Controls.FailNextCaption = true;
            dynamic before = f.Service.Tree(f.Project.Name, f.Form.Name);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateLabel(new Request { Project = f.Project.Name, Form = f.Form.Name, ControlPath = "Controls/Label1", NewName = "BrokenCopy", ExpectedTreeVersion = before.TreeVersion }));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
            Assert.AreEqual((string)before.TreeVersion, (string)((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion);
        }

        [TestMethod]
        public void DuplicateLabelRejectsStaleTreeAndNameCollisionBeforeAdd()
        {
            var f = Create();
            var request = new Request
            {
                Project = f.Project.Name,
                Form = f.Form.Name,
                ControlPath = "Controls/Label1",
                NewName = "LabelCopy",
                ExpectedTreeVersion = "stale"
            };
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateLabel(request));
            dynamic tree = f.Service.Tree(f.Project.Name, f.Form.Name);
            request.ExpectedTreeVersion = tree.TreeVersion;
            request.NewName = "Label1";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateLabel(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    public sealed partial class VbeFormsValueDuplicationTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void LabelCompleteGuardAndRollbackMatrix()
        {
            VerifyDuplicationGuards("Label");
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void LabelFontAndOleColorValidationAndReadbackMatrix()
        {
            foreach (double size in new[] { 0d, 201d, double.NaN, double.PositiveInfinity })
            {
                var f = Create("Label"); f.Source.Font.Size = size;
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, "Label", f.Request()));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            }
            {
                var f = Create("Label"); f.Source.BackColor = System.Drawing.Color.Red;
                f.Form.Designer.Controls.ConfigureAdded = copy => copy.ReadOverrides["BackColor"] = System.Drawing.Color.Red;
                Duplicate(f, "Label", f.Request());
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(System.Drawing.Color.Red, f.Form.Designer.Controls.Item("Copy").BackColor);
            }
            foreach (string property in new[] { "BackColor", "Font.Name", "Font.Size", "Font.Bold" })
            {
                var f = Create("Label");
                f.Form.Designer.Controls.ConfigureAdded = copy =>
                {
                    if (property == "BackColor") copy.ReadOverrides[property] = 255;
                    else copy.Font.ReadOverrides[property.Substring(5)] = property == "Font.Name" ? (object)"Other" : property == "Font.Size" ? (object)24d : true;
                };
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<InvalidOperationException>(() => Duplicate(f, "Label", f.Request()), property);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            }
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void PropertyAccessorTargetsAndNestedFontMatrix()
        {
            var f = Create("Label");
            foreach (string path in new[] { null, "", "UserForm", "Controls/Label1" })
            {
                var request = f.Request(); request.ControlPath = path; request.Property = "Caption";
                dynamic result = f.Service.PropertyAccessors(request);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Caption", (string)result.Property);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse((bool)result.Metadata.MetadataComplete);
            }
            foreach (string property in new[] { "", "Font.Name.More", "Other.Name" })
            {
                var request = f.Request(); request.Property = property;
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<ArgumentException>(() => f.Service.PropertyAccessors(request));
            }
            {
                var request = f.Request("Controls/Missing"); request.Property = "Caption";
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<InvalidOperationException>(() => f.Service.PropertyAccessors(request));
                request = f.Request(); request.Property = "Font.Name";
                dynamic result = f.Service.PropertyAccessors(request);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Name", (string)result.Metadata.Property);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("Font.Name", (string)result.Property);
            }
        }
    }
}

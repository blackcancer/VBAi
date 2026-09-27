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

    public sealed partial class VbeFormsPartialTests
    {
        [TestMethod]
        public void AppendListItemUsesCanonicalTreeAndVerifiesCount()
        {
            var f = Create("ComboBox");
            dynamic tree = f.Service.Tree(f.Project.Name, f.Form.Name);
            dynamic result = f.Service.AppendListItem(new Request { Project = f.Project.Name, Form = f.Form.Name, ControlPath = "Controls/ComboBox1", ExpectedTreeVersion = tree.TreeVersion, Text = "first" });
            Assert.IsTrue((bool)result.Applied);
            Assert.IsTrue((bool)result.Verified);
            Assert.AreEqual(0, (int)result.CountBefore);
            Assert.AreEqual(1, (int)result.CountAfter);
            Assert.AreEqual("first", f.Control.Items.Single());
        }

        [TestMethod]
        public void AppendListItemRejectsBoundMulticolumnAndStaleListsWithoutMutation()
        {
            var f = Create("ListBox");
            dynamic tree = f.Service.Tree(f.Project.Name, f.Form.Name);
            var request = new Request
            {
                Project = f.Project.Name,
                Form = f.Form.Name,
                ControlPath = "Controls/ListBox1",
                ExpectedTreeVersion = tree.TreeVersion,
                Text = "value"
            };
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AppendListItem(request));
            request.ExpectedTreeVersion = tree.TreeVersion;
            f.Control.RowSource = "Sheet1!A1:A3";
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AppendListItem(request));
            f.Control.RowSource = "";
            f.Control.ColumnCount = 2;
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AppendListItem(request));
            Assert.AreEqual(0, f.Control.ListCount);
        }

        [TestMethod]
        public void AddListItemRequiresVersionAndRejectsBoundControlBeforeNativeAdd()
        {
            var f = Create("ComboBox");
            dynamic tree = f.Service.Tree(f.Project.Name, f.Form.Name);
            var request = new Request
            {
                Project = f.Project.Name,
                Form = f.Form.Name,
                ControlPath = "Controls/ComboBox1",
                ExpectedTreeVersion = tree.TreeVersion,
                Text = "value"
            };
            Assert.ThrowsException<ArgumentException>(() => f.Service.AddListItem(request));
            request.ExpectedListVersion = "current";
            f.Control.RowSource = "Sheet1!A1:A3";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddListItem(request));
            Assert.AreEqual(0, f.Control.ListCount);
        }
    }
}

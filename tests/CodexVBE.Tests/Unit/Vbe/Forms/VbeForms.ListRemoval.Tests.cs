namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    public sealed partial class VbeFormsTests
    {
        [TestMethod]
        public void RemoveListItemUsesBothRevisionsAndVerifiesTheRemainingValues()
        {
            var fixture = NewFixture();
            var control = fixture.Form.Designer.Controls.AddExisting("Choices");
            control.ListRows.Add(new object[] { "A" });
            control.ListRows.Add(new object[] { "B" });
            control.ListRows.Add(new object[] { "C" });
            var provider = new NamedControlProvider("ListBox");
            TypeDescriptor.AddProvider(provider, control);
            try
            {
                dynamic before = fixture.Service.ListItems(new Request { Project = fixture.Project.Name, Form = fixture.Form.Name, ControlPath = "Controls/Choices", Limit = 64 });
                var request = new Request
                {
                    Project = fixture.Project.Name,
                    Form = fixture.Form.Name,
                    ControlPath = "Controls/Choices",
                    RowIndex = 1,
                    ExpectedTreeVersion = before.TreeVersion,
                    ExpectedListVersion = before.ListVersion
                };
                dynamic result = fixture.Service.RemoveListItem(request);
                Assert.IsTrue((bool)result.Applied);
                Assert.IsTrue((bool)result.Verified);
                Assert.AreEqual("B", (string)result.RemovedValue);
                Assert.AreEqual(1, control.RemoveAttempts);
                CollectionAssert.AreEqual(new[] { "A", "C" }, control.ListRows.Select(row => (string)row[0]).ToArray());
                Assert.AreNotEqual((string)before.ListVersion, (string)result.ListVersionAfter);
            }
            finally
            {
                TypeDescriptor.RemoveProvider(provider, control);
            }
        }

        [TestMethod]
        public void RemoveListItemRefusesBoundMulticolumnStaleAndUnverifiedNativeEffects()
        {
            var fixture = NewFixture();
            var control = fixture.Form.Designer.Controls.AddExisting("Choices");
            control.ListRows.Add(new object[] { "A" });
            var provider = new NamedControlProvider("ComboBox");
            TypeDescriptor.AddProvider(provider, control);
            try
            {
                dynamic before = fixture.Service.ListItems(new Request { Project = fixture.Project.Name, Form = fixture.Form.Name, ControlPath = "Controls/Choices", Limit = 64 });
                var request = new Request
                {
                    Project = fixture.Project.Name,
                    Form = fixture.Form.Name,
                    ControlPath = "Controls/Choices",
                    RowIndex = 0,
                    ExpectedTreeVersion = before.TreeVersion,
                    ExpectedListVersion = before.ListVersion
                };
                var missingVersion = request.ExpectedListVersion;
                request.ExpectedListVersion = null;
                Assert.ThrowsException<ArgumentException>(() => fixture.Service.RemoveListItem(request));
                request.ExpectedListVersion = missingVersion;
                request.RowIndex = -1;
                Assert.ThrowsException<ArgumentOutOfRangeException>(() => fixture.Service.RemoveListItem(request));
                request.RowIndex = 0;
                control.RowSource = "Sheet1!A1:A2";
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RemoveListItem(request));
                control.RowSource = null;
                control.ColumnCount = 2;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RemoveListItem(request));
                control.ColumnCount = 1;
                request.RowIndex = 2;
                Assert.ThrowsException<ArgumentOutOfRangeException>(() => fixture.Service.RemoveListItem(request));
                request.RowIndex = 0;
                request.ExpectedListVersion = "stale";
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RemoveListItem(request));
                request.ExpectedListVersion = before.ListVersion;
                control.ThrowRemove = true;
                dynamic failed = fixture.Service.RemoveListItem(request);
                Assert.IsFalse((bool)failed.Verified);
                Assert.IsTrue((bool)failed.VerificationPending);
                Assert.IsNull((object)failed.Applied);
                control.ThrowRemove = false;
                control.SkipRemove = true;
                dynamic pending = fixture.Service.RemoveListItem(request);
                Assert.IsTrue((bool)pending.Applied);
                Assert.IsFalse((bool)pending.Verified);
                Assert.AreEqual(1, control.ListRows.Count);
                for (int index = 0; index < 64; index++)
                    control.ListRows.Add(new object[] { "item" + index });
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RemoveListItem(request));
            }
            finally
            {
                TypeDescriptor.RemoveProvider(provider, control);
            }
        }
    }
}

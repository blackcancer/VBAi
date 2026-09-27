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
namespace CodexVBE.Tests.Unit
{
    using System;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    public sealed partial class VbeFormsTests
    {
        [TestMethod]
        public void ListRemovalValidatesEveryRequiredFieldAndWrongControlType()
        {
            var service = new VbeForms(new FakeVbe());
            MissingFields(request => service.RemoveListItem(request), "Project", "Form", "ControlPath", "ExpectedTreeVersion", "ExpectedListVersion", "RowIndex");
            var f=NewFixture(); f.Form.Designer.Controls.AddExisting("Choices");
            var r=RequiredListRequest(); r.Project=f.Project.Name; r.Form=f.Form.Name; r.ControlPath="Controls/Choices";
            Assert.ThrowsException<InvalidOperationException>(()=>f.Service.RemoveListItem(r));
            WithList("ListBox",(fixture,c,request)=>{
                c.ThrowOnCell=true;
                Assert.ThrowsException<InvalidOperationException>(()=>fixture.Service.RemoveListItem(request));
                c.ThrowOnCell=false;
                request.ExpectedTreeVersion="stale";
                Assert.ThrowsException<InvalidOperationException>(()=>fixture.Service.RemoveListItem(request));
                Assert.AreEqual(0,c.RemoveAttempts);
            });
        }

        [TestMethod]
        public void ListRemovalDetectsUnavailableOrChangedReadbackAndReportsNativeReadError()
        {
            foreach(var type in new[]{"ListBox","ComboBox"})
                for(int scenario=0;scenario<3;scenario++) WithList(type,(f,c,r)=>{
                    if(scenario==0) c.AfterRemove=control=>control.ThrowOnCell=true;
                    if(scenario==1) c.AfterRemove=control=>control.ListRows[0][0]="changed";
                    if(scenario==2) c.AfterRemove=control=>control.ThrowListCount=true;
                    r.ExpectedTreeVersion=((dynamic)f.Service.Tree(r.Project,r.Form)).TreeVersion; dynamic result=f.Service.RemoveListItem(r);
                    Assert.IsTrue((bool)result.Applied);
                    Assert.IsFalse((bool)result.Verified);
                    Assert.IsTrue((bool)result.VerificationPending);
                    if(scenario==2) StringAssert.Contains((string)result.NativeError,"ListCount unavailable");
                    Assert.AreEqual(1,c.RemoveAttempts);
                });
        }
    }
}

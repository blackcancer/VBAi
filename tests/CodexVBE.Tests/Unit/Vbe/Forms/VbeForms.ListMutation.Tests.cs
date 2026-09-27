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
namespace CodexVBE.Tests.Unit
{
    using System;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    public sealed partial class VbeFormsTests
    {
        [TestMethod]
        public void ListAddAndAppendValidateEveryRequiredFieldAndTextBoundary()
        {
            var service = new VbeForms(new FakeVbe());
            MissingFields(r => service.AddListItem(r), "Project", "Form", "ControlPath", "ExpectedTreeVersion", "ExpectedListVersion", "Text");
            MissingFields(r => service.AppendListItem(r), "Project", "Form", "ControlPath", "ExpectedTreeVersion", "Text");
            var oversized = RequiredListRequest(); oversized.Text = new string('x', 4097);
            Assert.ThrowsException<ArgumentException>(() => service.AddListItem(oversized));
            Assert.ThrowsException<ArgumentException>(() => service.AppendListItem(oversized));
        }

        [TestMethod]
        public void ListAddRejectsWrongTypeColumnsUnreadableAndStaleRevisions()
        {
            var f = NewFixture(); f.Form.Designer.Controls.AddExisting("Choices");
            var r = RequiredListRequest(); r.Project = f.Project.Name; r.Form = f.Form.Name; r.ControlPath = "Controls/Choices";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddListItem(r));
            foreach (var type in new[] { "ComboBox", "ListBox" }) WithList(type, (fixture,c,request) => {
                c.ColumnCount = 2;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AddListItem(request));
                c.ColumnCount = 1; c.ThrowOnCell = true;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AddListItem(request));
                c.ThrowOnCell = false;
                request.ExpectedTreeVersion = "stale";
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AddListItem(request));
                request.ExpectedTreeVersion = ((dynamic)fixture.Service.ListItems(request)).TreeVersion;
                request.ExpectedListVersion = "stale";
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AddListItem(request));
                Assert.AreEqual(0, c.AddAttempts);
            });
            WithList("ComboBox", (fixture,c,request) => {
                request.ExpectedListVersion = "version";
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AddListItem(request));
            }, 64);
        }

        [TestMethod]
        public void ListAddReportsNativeAndReadbackOutcomesWithoutAssumingSuccess()
        {
            foreach (var type in new[] { "ComboBox", "ListBox" }) {
                for (int scenario=0; scenario<6; scenario++) WithList(type, (f,c,r) => {
                    if (scenario == 1) c.ThrowAdd = true;
                    if (scenario == 2) c.SkipAdd = true;
                    if (scenario == 3) c.AfterAdd = control => control.ThrowOnCell = true;
                    if (scenario == 4) c.AfterAdd = control => control.ListRows[0][0] = "changed";
                    if (scenario == 5) c.AfterAdd = control => control.ThrowListCount = true;
                    r.ExpectedTreeVersion=((dynamic)f.Service.Tree(r.Project,r.Form)).TreeVersion; dynamic result = f.Service.AddListItem(r);
                    Assert.AreEqual(scenario == 1 ? (bool?)null : true, (bool?)result.Applied);
                    Assert.AreEqual(scenario == 0, (bool)result.Verified);
                    Assert.AreEqual(scenario != 0, (bool)result.VerificationPending);
                    if (scenario == 0) Assert.AreEqual("new", (string)result.AddedValue);
                    if (scenario == 1 || scenario == 5) Assert.IsNotNull((string)result.NativeError);
                    Assert.AreEqual(1, c.AddAttempts);
                });
            }
        }

        [TestMethod]
        public void ListAppendRejectsNoncanonicalWrongTypeAndCountBounds()
        {
            WithList("ComboBox", (fixture,c,request) => {
                request.ControlPath = "Controls/missing";
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AppendListItem(request));
            });
            var f = NewFixture(); f.Form.Designer.Controls.AddExisting("Choices");
            var r = RequiredListRequest(); r.Project=f.Project.Name; r.Form=f.Form.Name; r.ControlPath="Controls/Choices";
            r.ExpectedTreeVersion = ((dynamic)f.Service.Tree(r.Project,r.Form)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AppendListItem(r));
            foreach (int count in new[] { -1, 1024 }) WithList("ListBox", (fixture,c,request) => {
                c.ListCountOverride=count;
                request.ExpectedTreeVersion=((dynamic)fixture.Service.Tree(request.Project, request.Form)).TreeVersion;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AppendListItem(request));
                Assert.AreEqual(0,c.AddAttempts);
            });
        }

        [TestMethod]
        public void ListAppendReturnsVerifiedMismatchAndReadbackFailure()
        {
            for (int scenario=0; scenario<3; scenario++) WithList("ListBox", (f,c,r) => {
                if(scenario==1) c.SkipAdd=true;
                if(scenario==2) c.AfterAdd=control=>control.ThrowListCount=true;
                r.ExpectedTreeVersion=((dynamic)f.Service.Tree(r.Project,r.Form)).TreeVersion; dynamic result=f.Service.AppendListItem(r);
                Assert.IsTrue((bool)result.Applied);
                Assert.AreEqual(scenario==0,(bool)result.Verified);
                Assert.AreEqual(scenario!=0,(bool)result.VerificationPending);
                if(scenario==2) { Assert.IsNull((object)result.CountAfter); Assert.IsNotNull((string)result.ReadbackError); }
                else Assert.AreEqual(scenario==0?3:2,(int)result.CountAfter);
            });
        }
    }
}

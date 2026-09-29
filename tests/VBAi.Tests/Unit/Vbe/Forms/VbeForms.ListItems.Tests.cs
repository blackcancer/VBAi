namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeFormsTests
    {
        [TestMethod]
        public void ListItemsPaginatesIndexedCellsAndOnlyVersionsACompleteRead()
        {
            var fixture = NewFixture();
            var control = fixture.Form.Designer.Controls.AddExisting("Choices");
            control.ColumnCount = 2;
            control.ListRows.Add(new object[] { "un", 1 });
            control.ListRows.Add(new object[] { "deux", 2 });
            var provider = new NamedControlProvider("ListBox");
            TypeDescriptor.AddProvider(provider, control);
            try
            {
                var request = new Request
                {
                    Project = fixture.Project.Name,
                    Form = fixture.Form.Name,
                    ControlPath = "Controls/Choices",
                    Offset = 0,
                    Limit = 1
                };
                dynamic page = fixture.Service.ListItems(request);
                Assert.AreEqual(2, (int)page.TotalRows);
                Assert.AreEqual(1, (int)page.ReturnedRows);
                Assert.IsTrue((bool)page.HasMore);
                Assert.IsNull((string)page.ListVersion);
                request.Limit = 64;
                dynamic complete = fixture.Service.ListItems(request);
                Assert.AreEqual(2, (int)complete.ReturnedRows);
                Assert.IsFalse((bool)complete.HasMore);
                Assert.AreEqual(64, ((string)complete.ListVersion).Length);
                var rows = ((IEnumerable)complete.Rows).Cast<object>().ToArray();
                var cells = ((IEnumerable)((dynamic)rows[0]).Cells).Cast<object>().ToArray();
                Assert.AreEqual("un", (string)((dynamic)cells[0]).Value);
                Assert.AreEqual(1, (int)((dynamic)cells[1]).Value);
                control.ListRows[0][0] = "changed";
                dynamic changed = fixture.Service.ListItems(request);
                Assert.AreNotEqual((string)complete.ListVersion, (string)changed.ListVersion);
                control.ListRows.Clear();
                control.ColumnCount = 0;
                request.Offset = 100;
                request.Limit = 0;
                dynamic empty = fixture.Service.ListItems(request);
                Assert.AreEqual(0, (int)empty.ReturnedRows);
                Assert.AreEqual(0, (int)empty.ColumnCount);
                Assert.IsFalse((bool)empty.HasMore);
                Assert.AreEqual(64, ((string)empty.ListVersion).Length);
            }
            finally
            {
                TypeDescriptor.RemoveProvider(provider, control);
            }
        }

        [TestMethod]
        public void ListItemsReportsCellReadErrorsAndUnversionableValues()
        {
            var fixture = NewFixture();
            var control = fixture.Form.Designer.Controls.AddExisting("Choices");
            control.ListRows.Add(new object[] { "first" });
            var provider = new NamedControlProvider("ComboBox");
            TypeDescriptor.AddProvider(provider, control);
            try
            {
                var request = new Request
                {
                    Project = fixture.Project.Name,
                    Form = fixture.Form.Name,
                    ControlPath = "Controls/Choices",
                    Limit = 10
                };
                control.ThrowOnCell = true;
                dynamic failed = fixture.Service.ListItems(request);
                Assert.IsNull((string)failed.ListVersion);
                StringAssert.Contains((string)failed.ListVersionError, "could not be read");
                control.ThrowOnCell = false;
                control.ListRows[0][0] = new object ();
                dynamic objectValue = fixture.Service.ListItems(request);
                Assert.IsNull((string)objectValue.ListVersion);
                StringAssert.Contains((string)objectValue.ListVersionError, "non-scalar");
                control.ListRows[0][0] = new string ('x', 4097);
                dynamic oversized = fixture.Service.ListItems(request);
                Assert.IsNull((string)oversized.ListVersion);
                StringAssert.Contains((string)oversized.ListVersionError, "4096");
            }
            finally
            {
                TypeDescriptor.RemoveProvider(provider, control);
            }
        }

        [TestMethod]
        public void ListItemsRejectsInvalidPaginationPathControlAndColumnCount()
        {
            var fixture = NewFixture();
            var control = fixture.Form.Designer.Controls.AddExisting("Choices");
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ControlPath = "Controls/Choices",
                Offset = -1
            };
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => fixture.Service.ListItems(request));
            request.Offset = 0;
            request.Limit = -1;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => fixture.Service.ListItems(request));
            request.Limit = 0;
            request.ControlPath = "Choices";
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ListItems(request));
            request.ControlPath = "Controls/Choices";
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ListItems(request));
            var provider = new NamedControlProvider("ListBox");
            TypeDescriptor.AddProvider(provider, control);
            try
            {
                control.ColumnCount = 33;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ListItems(request));
                control.ColumnCount = -1;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ListItems(request));
            }
            finally
            {
                TypeDescriptor.RemoveProvider(provider, control);
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    using System;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    public sealed partial class VbeFormsTests
    {
        [TestMethod]
        public void ListReaderValidatesRequiredFieldsNegativeRowsAndNonzeroOffset()
        {
            var service = new VbeForms(new FakeVbe());
            MissingFields(r=>service.ListItems(r),"Project","Form","ControlPath");
            WithList("ListBox",(f,c,r)=>{
                c.ListCountOverride=-1;
                Assert.ThrowsException<InvalidOperationException>(()=>f.Service.ListItems(r));
                c.ListCountOverride=null; r.Offset=1;
                dynamic result=f.Service.ListItems(r);
                Assert.AreEqual(1,(int)result.ReturnedRows);
                Assert.IsNull((string)result.ListVersion);
                StringAssert.Contains((string)result.ListVersionError,"complete list");
            });
        }

        [TestMethod]
        public void ListReaderPreservesNullAndReportsConversionAndSerializationFailures()
        {
            WithList("ComboBox",(f,c,r)=>{
                c.ListRows[0][0]=null;
                dynamic nullable=f.Service.ListItems(r);
                Assert.IsNotNull((string)nullable.ListVersion);
                foreach(var invocation in new[]{false,true}) {
                    c.ListRows[0][0]=new CellTextFailure(invocation);
                    dynamic failed=f.Service.ListItems(r);
                    Assert.IsNull((string)failed.ListVersion);
                    StringAssert.Contains((string)failed.ListVersionError,"could not be read");
                }
                c.ListRows.Clear(); c.ColumnCount=32;
                for(int i=0;i<4;i++) {
                    var cells=new object[32];
                    for(int j=0;j<32;j++) cells[j]=new string('\0',4096);
                    c.ListRows.Add(cells);
                }
                dynamic oversized=f.Service.ListItems(r);
                Assert.IsNull((string)oversized.ListVersion);
                StringAssert.Contains((string)oversized.ListVersionError,"could not be calculated");
            });
        }
    }
}

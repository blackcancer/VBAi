namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Security.Cryptography;
    using System.Text;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeFormsInitializerTests
    {
        [TestMethod]
        public void InitializerWritesManagedBlockInsideExistingEventAndPreservesUserCode()
        {
            var f = Create();
            var request = RequestFor(f, "first", "a\"b");
            dynamic result = f.Service.SetListInitializer(request);
            Assert.IsTrue((bool)result.Applied);
            Assert.IsTrue((bool)result.Verified);
            Assert.IsTrue((bool)result.UserCodePreserved);
            Assert.IsTrue((bool)result.RuntimeVerificationPending);
            StringAssert.Contains(f.Form.CodeModule.Code, "    Debug.Print \"keep\"");
            StringAssert.Contains(f.Form.CodeModule.Code, "Me.ComboBox1.AddItem \"a\"\"b\"");
            Assert.IsTrue(f.Form.CodeModule.Code.IndexOf("CodexVBE BEGIN LIST", StringComparison.Ordinal) < f.Form.CodeModule.Code.IndexOf("End Sub", StringComparison.Ordinal));
            Assert.AreEqual(Sha(f.Form.CodeModule.Code), (string)result.Sha256After);
            Assert.AreEqual(0, f.List.AddCount);
        }

        [TestMethod]
        public void NestedInitializerTargetsExactPageAndRetainsManagedBlockProtection()
        {
            var f = Create();
            var frame = f.Form.Designer.Controls.Add("Frame", "Frame1");
            var pages = frame.Controls.Add("MultiPage", "MultiPage1");
            var page = pages.Pages.Add("Page", "Page1");
            var list = page.Controls.Add("ComboBox", "ComboBox1");
            const string path = "Controls/Frame1/Controls/MultiPage1/Pages/Page1/Controls/ComboBox1";
            var request = RequestFor(f, "nested"); request.ControlPath = path;
            dynamic result = f.Service.SetListInitializer(request);
            Assert.IsTrue((bool)result.Verified);
            StringAssert.Contains(f.Form.CodeModule.Code, "Me.Controls(\"Frame1\").Controls(\"MultiPage1\").Pages(\"Page1\").Controls(\"ComboBox1\").AddItem \"nested\"");
            Assert.AreEqual(0, list.AddCount); Assert.AreEqual(0, f.List.AddCount);
            request = RequestFor(f, "replaced"); request.ControlPath = path;
            Assert.IsTrue((bool)((dynamic)f.Service.SetListInitializer(request)).Verified);
            Assert.IsFalse(f.Form.CodeModule.Code.Contains("AddItem \"nested\""));
            f.Form.CodeModule.ReplaceText("AddItem \"replaced\"", "AddItem \"manual\"");
            string edited = f.Form.CodeModule.Code;
            request = RequestFor(f, "next"); request.ControlPath = path;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetListInitializer(request));
            Assert.AreEqual(edited, f.Form.CodeModule.Code);
        }

        [TestMethod]
        public void MatrixInitializerWritesColumnsAndCanReplaceItsOwnBlock()
        {
            var f = Create(); f.List.ColumnCount = 2;
            var request = RequestFor(f); request.Items = null;
            request.Rows = new[] { new[] { "first", "a\"b" }, new[] { "second", "" } };
            dynamic result = f.Service.SetListInitializer(request);
            Assert.IsTrue((bool)result.Verified);
            Assert.AreEqual(2, (int)result.Columns);
            StringAssert.Contains(f.Form.CodeModule.Code, "Me.ComboBox1.List(0, 1) = \"a\"\"b\"");
            StringAssert.Contains(f.Form.CodeModule.Code, "Me.ComboBox1.List(1, 1) = \"\"");
            Assert.AreEqual(0, f.List.AddCount);
            request.ExpectedSha256 = Sha(f.Form.CodeModule.Code);
            request.Rows = new[] { new[] { "new", "column" } };
            Assert.IsTrue((bool)((dynamic)f.Service.SetListInitializer(request)).Verified);
            Assert.IsFalse(f.Form.CodeModule.Code.Contains("List(1, 1)"));
            request.ExpectedSha256 = Sha(f.Form.CodeModule.Code); request.Rows = new string[0][];
            Assert.IsTrue((bool)((dynamic)f.Service.SetListInitializer(request)).Verified);
            Assert.IsFalse(f.Form.CodeModule.Code.Contains("AddItem"));
        }

        [TestMethod]
        public void MatrixInitializerRejectsMalformedOrMismatchingRowsWithoutWrites()
        {
            var f = Create(); f.List.ColumnCount = 2;
            var request = RequestFor(f); request.Items = null;
            request.Rows = new[] { new[] { "one", "two" }, new[] { "ragged" } };
            Assert.ThrowsException<ArgumentException>(() => f.Service.SetListInitializer(request));
            request.Rows = new[] { new[] { "one" } };
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetListInitializer(request));
            request.Items = new[] { "ambiguous" };
            Assert.ThrowsException<ArgumentException>(() => f.Service.SetListInitializer(request));
            Assert.AreEqual(0, f.Form.CodeModule.InsertCount);
        }

        [TestMethod]
        public void BindingReplacesManagedItemsAndRejectsManualChanges()
        {
            var f = Create(); f.Project.FileName = @"C:\Tests\Macro.xlsm"; f.List.ColumnCount = 2;
            var initial = RequestFor(f); initial.Items = null; initial.Rows = new[] { new[] { "a", "b" } };
            f.Service.SetListInitializer(initial);
            var request = RequestFor(f); request.Items = null; request.ExpectedHostPath = f.Project.FileName;
            request.SheetName = "Source \"quoted\""; request.RangeAddress = "$A$1:$B$2";
            dynamic result = f.Service.SetListBinding(request);
            Assert.IsTrue((bool)result.Verified);
            StringAssert.Contains(f.Form.CodeModule.Code, "ThisWorkbook.Worksheets(\"Source \"\"quoted\"\"\").Range(\"$A$1:$B$2\").Address(External:=True)");
            Assert.IsFalse(f.Form.CodeModule.Code.Contains(".Clear"));
            Assert.IsFalse(f.Form.CodeModule.Code.Contains(".AddItem"));
            request.ExpectedSha256 = Sha(f.Form.CodeModule.Code); request.RangeAddress = "A3:B4";
            Assert.IsTrue((bool)((dynamic)f.Service.SetListBinding(request)).Verified);
            f.Form.CodeModule.ReplaceText("A3:B4", "A5:B6");
            request.ExpectedSha256 = Sha(f.Form.CodeModule.Code);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetListBinding(request));
        }

        [TestMethod]
        public void BindingRejectsOtherHostAndColumnMismatchBeforeWriting()
        {
            var f = Create(); f.Project.FileName = @"C:\Tests\Macro.swp";
            var request = RequestFor(f); request.ExpectedHostPath = f.Project.FileName; request.SheetName = "Sheet1"; request.RangeAddress = "A1:B2";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetListBinding(request));
            f.Project.FileName = @"C:\Tests\Macro.xlsm"; request.ExpectedHostPath = f.Project.FileName;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetListBinding(request));
            Assert.AreEqual(0, f.Form.CodeModule.InsertCount);
        }

        [TestMethod]
        public void BindingRangeRejectsExpressionsAndOutOfBoundsReferences()
        {
            Assert.AreEqual(2, VbeForms.ExcelBindingColumns("Sheet1", "$A$1:$B$100"));
            Assert.AreEqual(1, VbeForms.ExcelBindingColumns("Sheet1", "XFD1048576"));
            Assert.ThrowsException<ArgumentException>(() => VbeForms.ExcelBindingColumns("Sheet'", "A1"));
            foreach (string invalid in new[] { "A:A", "A1,B2", "[other.xlsm]Sheet1!A1", "OFFSET(A1,0,0)", "B2:A1", "XFE1", "A1048577", "A1:K1", "A1:A10001" })
                Assert.ThrowsException<ArgumentException>(() => VbeForms.ExcelBindingColumns("Sheet1", invalid));
        }

        [TestMethod]
        public void GeneratedFormEditsPublishDiffAgainstTheFormModule()
        {
            var f = Create();
            var tools = new LlmVbeTools(new VbeSession(f.Host), null, new LlmSettings { VbeEditApproval = "Automatic" });
            CodeChange diff = null; tools.CodeEdited += change => diff = change;
            var request = RequestFor(f, "new item");
            var json = new System.Web.Script.Serialization.JavaScriptSerializer();
            var response = json.Deserialize<Response>(tools.Invoke("set_form_list_initializer", json.Serialize(new {
                request.Project, request.Form, request.ControlPath, request.ExpectedTreeVersion, request.ExpectedSha256, request.Items
            })));
            Assert.IsTrue(response.Ok, response.Error); Assert.IsNotNull(diff);
            Assert.AreEqual(f.Form.Name, diff.Module);
            StringAssert.Contains(diff.After, "new item");
            Assert.AreEqual(f.Form.CodeModule.CountOfLines, diff.AfterLineCount);
            Assert.AreEqual(request.ExpectedSha256, diff.BeforeSha256);
            Assert.AreEqual(Sha(f.Form.CodeModule.Code), diff.AfterSha256);
        }

        [TestMethod]
        public void InitializerCreatesMissingEventWithoutTouchingDesignerList()
        {
            var f = Create("");
            var request = RequestFor(f, "one");
            dynamic result = f.Service.SetListInitializer(request);
            Assert.IsTrue((bool)result.Applied);
            Assert.IsTrue((bool)result.Verified);
            Assert.AreEqual(1, f.Form.CodeModule.CreateEventCount);
            StringAssert.Contains(f.Form.CodeModule.Code, "Private Sub UserForm_Initialize()");
            StringAssert.Contains(f.Form.CodeModule.Code, "Me.ComboBox1.AddItem \"one\"");
            Assert.AreEqual(0, f.List.AddCount);
        }

        [TestMethod]
        public void InitializerIsIdempotentForCurrentIdenticalBlock()
        {
            var f = Create();
            f.Service.SetListInitializer(RequestFor(f, "one"));
            string before = f.Form.CodeModule.Code;
            dynamic again = f.Service.SetListInitializer(RequestFor(f, "one"));
            Assert.IsFalse((bool)again.Applied);
            Assert.IsTrue((bool)again.Verified);
            Assert.AreEqual(before, f.Form.CodeModule.Code);
            Assert.AreEqual(Sha(before), (string)again.Sha256After);
        }

        [TestMethod]
        public void InitializerReplacesOnlyPreviouslyManagedBlock()
        {
            var f = Create();
            f.Service.SetListInitializer(RequestFor(f, "old"));
            dynamic changed = f.Service.SetListInitializer(RequestFor(f, "new"));
            Assert.IsTrue((bool)changed.Applied);
            Assert.IsTrue((bool)changed.Verified);
            Assert.IsTrue((bool)changed.UserCodePreserved);
            StringAssert.Contains(f.Form.CodeModule.Code, "Debug.Print \"keep\"");
            StringAssert.Contains(f.Form.CodeModule.Code, "Me.ComboBox1.AddItem \"new\"");
            Assert.IsFalse(f.Form.CodeModule.Code.Contains("Me.ComboBox1.AddItem \"old\""));
            Assert.AreEqual(1, f.Form.CodeModule.Code.Split(new[] { "CodexVBE BEGIN LIST" }, StringSplitOptions.None).Length - 1);
        }

        [TestMethod]
        public void InitializerRefusesEditedManagedBlockWithoutChangingCode()
        {
            var f = Create();
            f.Service.SetListInitializer(RequestFor(f, "old"));
            f.Form.CodeModule.ReplaceText("Me.ComboBox1.AddItem \"old\"", "Me.ComboBox1.AddItem \"user edit\"");
            string edited = f.Form.CodeModule.Code;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetListInitializer(RequestFor(f, "new")));
            Assert.AreEqual(edited, f.Form.CodeModule.Code);
        }

        [TestMethod]
        public void InitializerReportsNativeInsertFailureAsPendingWithOriginalCodeIntact()
        {
            var f = Create();
            string before = f.Form.CodeModule.Code;
            f.Form.CodeModule.FailInsert = true;
            dynamic result = f.Service.SetListInitializer(RequestFor(f, "one"));
            Assert.IsNull((bool? )result.Applied);
            Assert.IsFalse((bool)result.Verified);
            Assert.IsTrue((bool)result.VerificationPending);
            StringAssert.Contains((string)result.NativeError, "insert refused");
            Assert.AreEqual(before, f.Form.CodeModule.Code);
        }

        [TestMethod]
        public void InitializerChecksTreeCodeAndBindingBeforeMutatingModule()
        {
            var f = Create();
            var request = RequestFor(f, "one");
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetListInitializer(request));
            request.ExpectedTreeVersion = RequestFor(f, "one").ExpectedTreeVersion;
            request.ExpectedSha256 = Sha("other code");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetListInitializer(request));
            request.ExpectedSha256 = Sha(f.Form.CodeModule.Code);
            f.List.RowSource = "Sheet1!A1:A3";
            request.ExpectedTreeVersion = RequestFor(f, "one").ExpectedTreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetListInitializer(request));
            Assert.AreEqual(0, f.Form.CodeModule.InsertCount);
        }
    }
}

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
        public void ListInitializerRejectsInvalidItemsAndInjectedPathsBeforeVbideAccess()
        {
            var service = new VbeForms(new FakeVbe());
            var request = new Request
            {
                Project = "VBAProject",
                Form = "Form1",
                ControlPath = "Controls/ComboBox1",
                ExpectedTreeVersion = "tree",
                ExpectedSha256 = "sha",
                Items = new[]
                {
                    "first\nsecond"
                }
            };
            Assert.ThrowsException<ArgumentException>(() => service.SetListInitializer(request));
            request.Items = new[]
            {
                "valid"
            };
            request.ControlPath = "Controls/Frame1/Controls/ComboBox1\"): Kill \"*";
            Assert.ThrowsException<ArgumentException>(() => service.SetListInitializer(request));
        }

        [TestMethod]
        public void ManagedListBlockEscapesQuotesAndRejectsEditedBody()
        {
            const string prefix = "' CodexVBE BEGIN LIST Controls/ComboBox1 SHA256=";
            const string end = "' CodexVBE END LIST Controls/ComboBox1";
            var generate = typeof(VbeForms).GetMethod("GenerateListBlock", BindingFlags.NonPublic | BindingFlags.Static);
            var validate = typeof(VbeForms).GetMethod("ValidateManagedBlock", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(generate);
            Assert.IsNotNull(validate);
            var lines = (string[])generate.Invoke(null, new object[] { "ComboBox1", new[] { "a\"b", "second" }, prefix, end });
            Assert.AreEqual("    Me.ComboBox1.AddItem \"a\"\"b\"", lines[2]);
            validate.Invoke(null, new object[] { lines, 0, lines.Length - 1, "ComboBox1", prefix });
            lines[2] = "    Me.ComboBox1.AddItem \"changed\"";
            var error = Assert.ThrowsException<TargetInvocationException>(() => validate.Invoke(null, new object[] { lines, 0, lines.Length - 1, "ComboBox1", prefix }));
            Assert.IsInstanceOfType(error.InnerException, typeof(InvalidOperationException));
        }
    }
}

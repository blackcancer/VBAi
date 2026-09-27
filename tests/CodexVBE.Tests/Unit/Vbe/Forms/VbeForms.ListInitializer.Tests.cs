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
        public void ListInitializerRejectsInvalidItemsAndNestedPathsBeforeVbideAccess()
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
            request.ControlPath = "Controls/Frame1/Controls/ComboBox1";
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

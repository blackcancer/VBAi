using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class GitFailureDiagnosticTests
    {
        [TestMethod]
        public void DescribeReportsComAndManagedTypeAndHResultWithoutExceptionText()
        {
            var com = new COMException("SECRET COM content C:\\private\\form.frm", unchecked((int)0x80020003));
            var managed = new InvalidOperationException("SECRET managed content C:\\private\\form.frx");

            string comDescription = GitFailureDiagnostic.Describe(com);
            string managedDescription = GitFailureDiagnostic.Describe(managed);

            StringAssert.Contains(comDescription, "COMException 0x80020003");
            StringAssert.Contains(managedDescription, "InvalidOperationException 0x" + managed.HResult.ToString("X8"));
            Assert.IsFalse(comDescription.Contains("SECRET"));
            Assert.IsFalse(managedDescription.Contains("SECRET"));
            Assert.IsFalse(comDescription.Contains("C:\\private"));
            Assert.IsFalse(managedDescription.Contains("C:\\private"));
        }

        [TestMethod]
        public void DescribeKeepsOriginalInnerChainAndBoundsItsOutput()
        {
            var deepest = new OverflowException("SECRET last C:\\private\\last.frm");
            var omitted = new NotSupportedException("SECRET fifth C:\\private\\fifth.frm", deepest);
            var fourth = new FormatException("SECRET fourth C:\\private\\fourth.frm", omitted);
            var third = new ArgumentException("SECRET third C:\\private\\third.frm", fourth);
            var root = new InvalidOperationException("SECRET first C:\\private\\first.frm",
                new ArgumentException("SECRET second C:\\private\\second.frm", third));

            string description = GitFailureDiagnostic.Describe(root);

            StringAssert.Contains(description, "InvalidOperationException 0x" + root.HResult.ToString("X8"));
            StringAssert.Contains(description, "ArgumentException 0x" + root.InnerException.HResult.ToString("X8"));
            StringAssert.Contains(description, "FormatException 0x" + fourth.HResult.ToString("X8"));
            Assert.IsFalse(description.Contains("NotSupportedException"), "Only the first four chain entries may be reported.");
            Assert.IsFalse(description.Contains("OverflowException"));
            Assert.IsFalse(description.Contains("SECRET"));
            Assert.IsFalse(description.Contains("C:\\private"));
            Assert.IsTrue(description.Length <= 1200, "Exception metadata must remain bounded.");
            Assert.AreSame(third, root.InnerException.InnerException);
            Assert.AreSame(fourth, third.InnerException);
            Assert.AreSame(omitted, fourth.InnerException);
            Assert.AreSame(deepest, omitted.InnerException);
        }

        [TestMethod]
        public void DescribeRejectsNull()
        {
            Assert.ThrowsException<ArgumentNullException>(() => GitFailureDiagnostic.Describe(null));
        }

        [TestMethod]
        public void DescribeUsesFirstProductFrameWithoutSourcePath()
        {
            InvalidOperationException failure = Assert.ThrowsException<InvalidOperationException>(
                () => FormFontRestoration.ValidateDescriptor(null));

            string description = GitFailureDiagnostic.Describe(failure);

            StringAssert.Contains(description, "ValidateDescriptor");
            Assert.IsFalse(description.Contains("FormFontRestoration.cs"), "Do not expose a source file path.");
            Assert.IsFalse(description.Contains("E:\\"), "Do not expose a checkout path.");
            Assert.IsFalse(description.Contains("Unsupported persisted standard font descriptor."),
                "Only exception metadata belongs in this diagnostic.");
        }
    }
}

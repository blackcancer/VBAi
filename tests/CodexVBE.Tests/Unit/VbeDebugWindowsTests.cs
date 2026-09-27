using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CodexVBE;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeDebugWindowsTests
    {
        [TestMethod]
        public void ImmediateCommandRejectsMultilineControlAndOversizedInputBeforeNativeUi()
        {
            foreach (string invalid in new[] {
                null, "", "   ", "Debug.Print 1\r\nDebug.Print 2",
                "Debug.Print 1\0", "Debug.Print " + (char)1,
                new string('x', 2049)
            })
            {
                var error = Assert.ThrowsException<ArgumentException>(
                    () => VbeDebugWindows.ExecuteImmediate(invalid));
                StringAssert.Contains(error.Message, "one nonempty printable line");
            }
        }

        [TestMethod]
        public void DebugTreeMutationRejectsMalformedTargetBeforeNativeUi()
        {
            var invalid = new[] {
                new Request { Pane = "locals", Action = "expand" },
                new Request { Pane = "locals", Action = "expand", PathSegments = new string[0] },
                new Request { Pane = "locals", Action = "expand", PathSegments = new[] { "root", " " } },
                new Request { Pane = "locals", Action = "expand", PathSegments = new string[17] },
                new Request { Pane = "immediate", Action = "expand", PathSegments = new[] { "root" } },
                new Request { Pane = "watches", Action = "delete", PathSegments = new[] { "root" } }
            };
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ChangeDebugItem(null));
            foreach (var request in invalid)
                Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ChangeDebugItem(request));
        }

        [TestMethod]
        public void DiagnosticResponseRequiresExactMessageAndButtonBeforeNativeUi()
        {
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.RespondDebugDialog(null));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.RespondDebugDialog(
                new Request { Diagnostic = " ", Button = "OK" }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.RespondDebugDialog(
                new Request { Diagnostic = "Compile error", Button = "" }));
        }

        [TestMethod]
        public void WatchSelectionRequiresExpressionAndContextBeforeNativeUi()
        {
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectWatch(null));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectWatch(
                new Request { Expression = "counter", Context = " " }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectWatch(
                new Request { Expression = "", Context = "Project.Module.Procedure" }));
        }

        [TestMethod]
        public void SignaturePlaceholderDetectionAcceptsOnlyKnownLocalizedLabels()
        {
            var method = typeof(VbeDebugWindows).GetMethod("IsNoCertificate",
                BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method);
            foreach (var label in new[] { "[Aucun certificat]", "[NO CERTIFICATE]", "[None]" })
                Assert.AreEqual(true, method.Invoke(null, new object[] { label }));
            foreach (var label in new[] { null, "", "None", "[My certificate]" })
                Assert.AreEqual(false, method.Invoke(null, new object[] { label }));
        }

        [TestMethod]
        public void CertificateNameParserRequiresAdjacentHeading()
        {
            var method = typeof(VbeDebugWindows).GetMethod("CertificateBeforeHeading",
                BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method);
            var headings = new[] { "Certificat actuel", "Current certificate" };
            var valid = new List<string> { "header", "My certificate", "Nom du certificat :", "Certificat actuel" };
            Assert.AreEqual("My certificate", method.Invoke(null, new object[] { valid, headings }));
            var invalid = new List<string> { "header", "My certificate", "Unrelated label", "Certificat actuel" };
            Assert.IsNull(method.Invoke(null, new object[] { invalid, headings }));
        }
    }
}

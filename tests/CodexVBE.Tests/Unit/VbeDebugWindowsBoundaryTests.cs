using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeDebugWindowsBoundaryTests
    {
        private static object CallPrivate(string name, params object[] arguments)
        {
            var method = typeof(VbeDebugWindows).GetMethod(name,
                BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method, "Missing private helper: " + name);
            return method.Invoke(null, arguments);
        }

        [TestMethod]
        public void MissingListWindowIsUnavailableRatherThanAnEmptyVerifiedCollection()
        {
            dynamic result = CallPrivate("ReadList", IntPtr.Zero);
            Assert.IsFalse((bool)result.Available);
            Assert.AreEqual("UIAExposedRowsOnly", (string)result.Coverage);
            Assert.AreEqual(0, ((IEnumerable)result.Items).Cast<object>().Count());
            StringAssert.Contains((string)result.Error, "not visible");
        }

        [TestMethod]
        public void MissingImmediateWindowReturnsNoObservedText()
        {
            dynamic result = CallPrivate("ReadImmediate", IntPtr.Zero);
            Assert.IsFalse((bool)result.Available);
            Assert.IsNull((string)result.Text);
            StringAssert.Contains((string)result.Error, "not visible");
        }

        [TestMethod]
        public void CertificatePlaceholderRecognitionIsLocalizedAndExact()
        {
            foreach (var label in new[] { "[Aucun certificat]", "[NO CERTIFICATE]", "[None]" })
                Assert.AreEqual(true, CallPrivate("IsNoCertificate", label));
            foreach (var label in new[] { null, "", "No certificate", " [None] ", "[Another certificate]" })
                Assert.AreEqual(false, CallPrivate("IsNoCertificate", label));
        }

        [TestMethod]
        public void CertificateParserAcceptsAdjacentFrenchAndEnglishHeadings()
        {
            var french = new List<string> { "ignored", "Certificat A", "Nom du certificat :",
                "Signature actuelle du projet VBA" };
            Assert.AreEqual("Certificat A", CallPrivate("CertificateBeforeHeading", french,
                new[] { "Signature actuelle du projet VBA" }));
            var english = new List<string> { "ignored", "Certificate B", "Certificate name:",
                "The VBA project is currently signed as" };
            Assert.AreEqual("Certificate B", CallPrivate("CertificateBeforeHeading", english,
                new[] { "The VBA project is currently signed as" }));
        }

        [TestMethod]
        public void CertificateParserRejectsNonadjacentOrUnrecognizedLabels()
        {
            var unrelated = new List<string> { "Certificate A", "Other label", "Sign as" };
            Assert.IsNull(CallPrivate("CertificateBeforeHeading", unrelated, new[] { "Sign as" }));
            var tooShort = new List<string> { "Certificate name:", "Sign as" };
            Assert.IsNull(CallPrivate("CertificateBeforeHeading", tooShort, new[] { "Sign as" }));
            var wrongHeading = new List<string> { "Certificate A", "Certificate name:", "Current certificate" };
            Assert.IsNull(CallPrivate("CertificateBeforeHeading", wrongHeading, new[] { "Sign as" }));
        }

        [TestMethod]
        public void ImmediateInputRejectsControlCharactersAndLengthBeforeWindowLookup()
        {
            foreach (var input in new[] {
                "Debug.Print 1\r", "Debug.Print 1\n", "Debug.Print 1\0",
                "Debug.Print " + (char)127, new string('x', 2049)
            })
            {
                var error = Assert.ThrowsException<ArgumentException>(() =>
                    VbeDebugWindows.ExecuteImmediate(input));
                StringAssert.Contains(error.Message, "one nonempty printable line");
            }
        }

        [TestMethod]
        public void DebugTreeMutationRequiresExactPaneActionAndBoundedPath()
        {
            var validPath = new[] { "root" };
            foreach (var request in new[] {
                new Request { Pane = "Locals", Action = "expand", PathSegments = validPath },
                new Request { Pane = "locals", Action = "Expand", PathSegments = validPath },
                new Request { Pane = "watches", Action = "expand", PathSegments = new string[17] },
                new Request { Pane = "locals", Action = "collapse", PathSegments = new[] { "root", "\t" } }
            })
                Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ChangeDebugItem(request));
        }

        [TestMethod]
        public void DialogAndWatchSelectionRequireExactUserEvidenceBeforeWindowLookup()
        {
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.RespondDebugDialog(
                new Request { Diagnostic = "Run-time error", Button = " " }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.RespondDebugDialog(
                new Request { Diagnostic = " ", Button = "End" }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectWatch(
                new Request { Expression = " ", Context = "VBAProject.Module1.Main" }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectWatch(
                new Request { Expression = "counter", Context = " " }));
        }
    }
}

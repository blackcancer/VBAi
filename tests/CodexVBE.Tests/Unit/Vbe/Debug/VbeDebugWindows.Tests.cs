namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeDebugWindowsBoundaryTests
    {
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
            foreach (var label in new[]
            {
                "[Aucun certificat]",
                "[NO CERTIFICATE]",
                "[None]"
            }

            )
                Assert.AreEqual(true, CallPrivate("IsNoCertificate", label));
            foreach (var label in new[]
            {
                null,
                "",
                "No certificate",
                " [None] ",
                "[Another certificate]"
            }

            )
                Assert.AreEqual(false, CallPrivate("IsNoCertificate", label));
        }

        [TestMethod]
        public void CertificateParserAcceptsAdjacentFrenchAndEnglishHeadings()
        {
            var french = new List<string>
            {
                "ignored",
                "Certificat A",
                "Nom du certificat :",
                "Signature actuelle du projet VBA"
            };
            Assert.AreEqual("Certificat A", CallPrivate("CertificateBeforeHeading", french, new[] { "Signature actuelle du projet VBA" }));
            var english = new List<string>
            {
                "ignored",
                "Certificate B",
                "Certificate name:",
                "The VBA project is currently signed as"
            };
            Assert.AreEqual("Certificate B", CallPrivate("CertificateBeforeHeading", english, new[] { "The VBA project is currently signed as" }));
        }

        [TestMethod]
        public void CertificateParserRejectsNonadjacentOrUnrecognizedLabels()
        {
            var unrelated = new List<string>
            {
                "Certificate A",
                "Other label",
                "Sign as"
            };
            Assert.IsNull(CallPrivate("CertificateBeforeHeading", unrelated, new[] { "Sign as" }));
            var tooShort = new List<string>
            {
                "Certificate name:",
                "Sign as"
            };
            Assert.IsNull(CallPrivate("CertificateBeforeHeading", tooShort, new[] { "Sign as" }));
            var wrongHeading = new List<string>
            {
                "Certificate A",
                "Certificate name:",
                "Current certificate"
            };
            Assert.IsNull(CallPrivate("CertificateBeforeHeading", wrongHeading, new[] { "Sign as" }));
        }

        [TestMethod]
        public void ImmediateInputRejectsControlCharactersAndLengthBeforeWindowLookup()
        {
            foreach (var input in new[]
            {
                "Debug.Print 1\r",
                "Debug.Print 1\n",
                "Debug.Print 1\0",
                "Debug.Print " + (char)127,
                new string ('x', 2049)
            }

            )
            {
                var error = Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ExecuteImmediate(input));
                StringAssert.Contains(error.Message, "one nonempty printable line");
            }
        }

        [TestMethod]
        public void DebugTreeMutationRequiresExactPaneActionAndBoundedPath()
        {
            var validPath = new[]
            {
                "root"
            };
            foreach (var request in new[]
            {
                new Request
                {
                    Pane = "Locals",
                    Action = "expand",
                    PathSegments = validPath
                },
                new Request
                {
                    Pane = "locals",
                    Action = "Expand",
                    PathSegments = validPath
                },
                new Request
                {
                    Pane = "watches",
                    Action = "expand",
                    PathSegments = new string[17]
                },
                new Request
                {
                    Pane = "locals",
                    Action = "collapse",
                    PathSegments = new[]
                    {
                        "root",
                        "\t"
                    }
                }
            }

            )
                Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ChangeDebugItem(request));
        }

        [TestMethod]
        public void DialogAndWatchSelectionRequireExactUserEvidenceBeforeWindowLookup()
        {
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.RespondDebugDialog(new Request { Diagnostic = "Run-time error", Button = " " }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.RespondDebugDialog(new Request { Diagnostic = " ", Button = "End" }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectWatch(new Request { Expression = " ", Context = "VBAProject.Module1.Main" }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectWatch(new Request { Expression = "counter", Context = " " }));
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeDebugWindowsImmediateProbeTests
    {
        [TestMethod]
        public void ImmediateRejectsNonPrintableMultilineOrOversizedCommandsBeforeNativeAccess()
        {
            var fake = new ImmediateFake();
            foreach (string command in new[]
            {
                null,
                "",
                " ",
                "a\nb",
                "a\rb",
                "a\0b",
                "a\tb",
                new string ('x', 2049)
            }

            )
                Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ExecuteImmediate(command, fake));
            Assert.AreEqual(0, fake.RootReads);
        }

        [TestMethod]
        public void ImmediateRequiresVisibleVbeAndPaneBeforePreparingSelection()
        {
            var fake = new ImmediateFake();
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ExecuteImmediate("? 1", fake));
            fake.Root = new IntPtr(1);
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ExecuteImmediate("? 1", fake));
            Assert.AreEqual(0, fake.Prepares);
        }

        [TestMethod]
        public void ImmediateDoesNotSendEnterUntilEveryCharacterEchoesExactly()
        {
            var fake = new ImmediateFake
            {
                Root = new IntPtr(1),
                PaneHandle = new IntPtr(2),
                RejectChar = '?'
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ExecuteImmediate("? 1", fake));
            Assert.AreEqual(0, fake.Enters);
            fake = new ImmediateFake
            {
                Root = new IntPtr(1),
                PaneHandle = new IntPtr(2)
            };
            fake.Readbacks.Add("unexpected");
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ExecuteImmediate("? 1", fake));
            Assert.AreEqual(40, fake.Pauses);
            Assert.AreEqual(0, fake.Enters);
        }

        [TestMethod]
        public void ImmediateDistinguishesRejectedEnterPendingAndChangedOutput()
        {
            var fake = Ready("? 1");
            fake.EnterSucceeds = false;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ExecuteImmediate("? 1", fake));
            Assert.AreEqual(1, fake.Enters);
            fake = Ready("? 1");
            dynamic pending = VbeDebugWindows.ExecuteImmediate("? 1", fake);
            Assert.AreEqual("Pending", (string)pending.Verification);
            Assert.IsTrue((bool)pending.VerificationPending);
            Assert.AreEqual("? 1", (string)pending.OutputDelta);
            fake = Ready("? 1");
            fake.Readbacks.Add("> ? 1\r\n1\r\n");
            dynamic changed = VbeDebugWindows.ExecuteImmediate("? 1", fake);
            Assert.AreEqual("ImmediateTextChangedAfterEnter", (string)changed.Verification);
            Assert.IsFalse((bool)changed.VerificationPending);
            Assert.AreEqual("? 1\r\n1\r\n", (string)changed.OutputDelta);
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeDebugWindowsNativeTests
    {
        [TestMethod]
        public void CaptureRejectsMissingHostWindowBeforeReadingPanes()
        {
            var native = new FakeNative();
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.Capture(false, native));
            Assert.AreEqual(0, native.ListReads.Count);
            Assert.AreEqual(0, native.CallStackReads);
        }

        [TestMethod]
        public void CaptureRoutesVisiblePanesAndOnlyReadsRequestedCallStack()
        {
            var native = new FakeNative
            {
                Root = new IntPtr(1)
            };
            native.Panes["Variables locales"] = new IntPtr(11);
            native.Panes["Espions"] = new IntPtr(12);
            native.Panes["Exécution"] = new IntPtr(13);
            dynamic withoutStack = VbeDebugWindows.Capture(false, native);
            Assert.AreEqual(0, native.CallStackReads);
            Assert.IsNull((object)withoutStack.CallStack);
            CollectionAssert.AreEqual(new[] { new IntPtr(11), new IntPtr(12) }, native.ListReads.ToArray());
            Assert.AreEqual(new IntPtr(13), native.ImmediateHandle);
            Assert.IsTrue((bool)withoutStack.Locals.Available);
            Assert.IsTrue((bool)withoutStack.Watches.Available);
            dynamic withStack = VbeDebugWindows.Capture(true, native);
            Assert.AreEqual(1, native.CallStackReads);
            Assert.AreEqual(new IntPtr(11), native.CallStackHandle);
            Assert.IsTrue((bool)withStack.CallStack.Available);
        }

        [TestMethod]
        public void CaptureDistinguishesMissingPanesFromEmptyDebuggerCollections()
        {
            var native = new FakeNative
            {
                Root = new IntPtr(1)
            };
            dynamic snapshot = VbeDebugWindows.Capture(true, native);
            Assert.IsFalse((bool)snapshot.Locals.Available);
            Assert.IsFalse((bool)snapshot.Watches.Available);
            Assert.IsFalse((bool)snapshot.Immediate.Available);
            Assert.AreEqual(IntPtr.Zero, native.CallStackHandle);
            StringAssert.Contains((string)snapshot.Limits, "missing pane");
        }

        [TestMethod]
        public void ExistingCompileOrOptionsDialogPreventsOpeningAnother()
        {
            var native = new FakeNative
            {
                DialogHandle = new IntPtr(2)
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.EnsureNoCompileDialog(native));
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.EnsureNoDebugOptionsDialog(native));
            native.DialogHandle = IntPtr.Zero;
            VbeDebugWindows.EnsureNoCompileDialog(native);
            VbeDebugWindows.EnsureNoDebugOptionsDialog(native);
        }

        [TestMethod]
        public void ReadDebugDialogReportsAbsenceAndFiltersInvisibleControls()
        {
            var native = new FakeNative();
            dynamic absent = VbeDebugWindows.ReadDebugDialog(native);
            Assert.IsFalse((bool)absent.Available);
            Assert.IsNull((string)absent.Diagnostic);
            native.DialogHandle = new IntPtr(2);
            native.Controls.Add(Control(3, "Static", "Compile error: Syntax error", false));
            native.Controls.Add(Control(4, "Static", "Run-time error '9'", true));
            native.Controls.Add(Control(5, "Button", "End", true));
            native.Controls.Add(Control(6, "Button", "hidden", false));
            dynamic found = VbeDebugWindows.ReadDebugDialog(native);
            Assert.IsTrue((bool)found.Available);
            Assert.AreEqual("Run-time error '9'", (string)found.Diagnostic);
            CollectionAssert.AreEqual(new[] { "End" }, (string[])found.Buttons);
            Assert.IsNull((string)found.Error);
        }

        [TestMethod]
        public void ReadDebugDialogRefusesAmbiguousDiagnosticText()
        {
            var native = new FakeNative
            {
                DialogHandle = new IntPtr(2)
            };
            dynamic missing = VbeDebugWindows.ReadDebugDialog(native);
            StringAssert.Contains((string)missing.Error, "found 0");
            native.Controls.Add(Control(3, "Static", "Compile error: one", true));
            native.Controls.Add(Control(4, "Static", "Compile error: two", true));
            dynamic result = VbeDebugWindows.ReadDebugDialog(native);
            Assert.IsTrue((bool)result.Available);
            Assert.IsNull((string)result.Diagnostic);
            StringAssert.Contains((string)result.Error, "found 2");
        }

        [TestMethod]
        public void RespondDebugDialogChecksExactMessageAndUniqueButtonBeforeClick()
        {
            var native = DialogWith("Compile error: Syntax error", "OK");
            var request = new Request
            {
                Diagnostic = "Compile error: other",
                Button = "OK"
            };
            native.DialogHandle = IntPtr.Zero;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.RespondDebugDialog(request, native));
            native.DialogHandle = new IntPtr(2);
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.RespondDebugDialog(request, native));
            request.Diagnostic = "Compile error: Syntax error";
            request.Button = "Cancel";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.RespondDebugDialog(request, native));
            request.Button = "OK";
            native.Controls.Add(Control(5, "Button", "OK", true));
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.RespondDebugDialog(request, native));
            Assert.AreEqual(0, native.Clicks);
        }

        [TestMethod]
        public void RespondDebugDialogRejectsUnrecognizedDiagnosticsAndFailedClick()
        {
            var native = DialogWith("Generic information", "OK");
            var request = new Request
            {
                Diagnostic = "Generic information",
                Button = "OK"
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.RespondDebugDialog(request, native));
            native.Controls[0].Text = "Compile error: Syntax error";
            request.Diagnostic = native.Controls[0].Text;
            native.ClickSucceeds = false;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.RespondDebugDialog(request, native));
            Assert.AreEqual(1, native.Clicks);
        }

        [TestMethod]
        public void RespondDebugDialogDistinguishesClosedAndPendingAfterClick()
        {
            var native = DialogWith("Compile error: Syntax error", "OK");
            var request = new Request
            {
                Diagnostic = "Compile error: Syntax error",
                Button = "OK"
            };
            native.CloseAfterClick = true;
            dynamic closed = VbeDebugWindows.RespondDebugDialog(request, native);
            Assert.AreEqual("DialogClosed", (string)closed.Verification);
            Assert.IsFalse((bool)closed.VerificationPending);
            native.CloseAfterClick = false;
            native.PausesSinceReset = 0;
            dynamic pending = VbeDebugWindows.RespondDebugDialog(request, native);
            Assert.AreEqual("Pending", (string)pending.Verification);
            Assert.IsTrue((bool)pending.VerificationPending);
            Assert.AreEqual(40, native.PausesSinceReset);
        }

        [TestMethod]
        public void AwaitCompileDialogReturnsDiagnosticOnlyAfterOwnOkAndCompletion()
        {
            var native = DialogWith("Compile error: Syntax error", "OK");
            using (var completed = new ManualResetEventSlim(true))
            {
                Assert.AreEqual("Compile error: Syntax error", VbeDebugWindows.AwaitCompileDialog(completed, native));
                Assert.AreEqual(1, native.Clicks);
            }
        }

        [TestMethod]
        public void AwaitCompileDialogRejectsMissingOkAndTimesOutWithoutDiagnostic()
        {
            var native = DialogWith("Compile error: Syntax error", "Cancel");
            using (var completed = new ManualResetEventSlim(true))
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.AwaitCompileDialog(completed, native));
            native.Controls[1].Text = "OK";
            native.ClickSucceeds = false;
            using (var completed = new ManualResetEventSlim(true))
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.AwaitCompileDialog(completed, native));
            native.DialogHandle = IntPtr.Zero;
            using (var completed = new ManualResetEventSlim(true))
                Assert.IsNull(VbeDebugWindows.AwaitCompileDialog(completed, native));
            using (var completed = new ManualResetEventSlim(false))
                Assert.ThrowsException<TimeoutException>(() => VbeDebugWindows.AwaitCompileDialog(completed, native));
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeDebugWindowsOptionsProbeTests
    {
        [TestMethod]
        public void OptionsRequireDialogAndBoundedReadableTabs()
        {
            var fake = new OptionsFake
            {
                Open = false
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadVbeOptions(fake));
            Assert.AreEqual(60, fake.Pauses);
            fake.Open = true;
            fake.Names.Clear();
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadVbeOptions(fake));
            Assert.AreEqual(1, fake.Closes);
            fake = new OptionsFake();
            fake.Names.Add(null);
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadVbeOptions(fake));
            fake = new OptionsFake();
            for (int index = 0; index < 8; index++)
                fake.Names.Add("Extra " + index);
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadVbeOptions(fake));
        }

        [TestMethod]
        public void OptionsFiltersHiddenDisabledAndBlankTextWhileKeepingOtherValues()
        {
            var fake = new OptionsFake
            {
                CloseAfterRead = true
            };
            fake.Items.Add(new VbeDebugWindows.OptionsControl { Name = "Visible", Type = "ControlType.CheckBox", Value = "On" });
            fake.Items.Add(new VbeDebugWindows.OptionsControl { Name = "Hidden", Type = "ControlType.Edit", Visible = false });
            fake.Items.Add(new VbeDebugWindows.OptionsControl { Name = "Disabled", Type = "ControlType.Edit", Enabled = false });
            fake.Items.Add(new VbeDebugWindows.OptionsControl { Name = "", Type = "ControlType.Text" });
            fake.Items.Add(new VbeDebugWindows.OptionsControl { Name = "", Type = "ControlType.Edit", Value = "editable" });
            dynamic result = VbeDebugWindows.ReadVbeOptions(fake);
            Assert.AreEqual(1, (int)result.Count);
            Assert.AreEqual(2, (int)result.Tabs[0].Count);
            Assert.AreEqual("Visible", (string)result.Tabs[0].Controls[0].Name);
            Assert.AreEqual("editable", (string)result.Tabs[0].Controls[1].Value);
            Assert.IsTrue((bool)result.DialogClosed);
        }

        [TestMethod]
        public void OptionsRejectControlExplosionAndUnclosedDialog()
        {
            var fake = new OptionsFake();
            for (int index = 0; index < 2001; index++)
                fake.Items.Add(new VbeDebugWindows.OptionsControl { Name = "Item", Type = "ControlType.Text" });
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadVbeOptions(fake));
            Assert.AreEqual(1, fake.Closes);
            fake = new OptionsFake();
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadVbeOptions(fake));
            Assert.AreEqual(20, fake.ClosePolls);
        }

        [TestMethod]
        public void DebugOptionsRequireThreeReadableNamedChoicesAndExactlyOneSelection()
        {
            var fake = new OptionsFake
            {
                Open = false
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadDebugOptions(fake));
            fake = DebugFake();
            fake.Choices.RemoveAt(2);
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadDebugOptions(fake));
            fake = DebugFake();
            fake.Choices[0].Name = "";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadDebugOptions(fake));
            fake = DebugFake();
            fake.Choices[0].Readable = false;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadDebugOptions(fake));
            fake = DebugFake();
            fake.Choices[0].Name = "Unrelated";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadDebugOptions(fake));
            fake = DebugFake();
            fake.Choices[1].Selected = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadDebugOptions(fake));
            fake = DebugFake();
            fake.Choices[0].Selected = false;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadDebugOptions(fake));
        }

        [TestMethod]
        public void DebugOptionsReturnsChosenRadioOnlyAfterOwnDialogCloses()
        {
            var fake = DebugFake();
            fake.CloseAfterRead = true;
            dynamic result = VbeDebugWindows.ReadDebugOptions(fake);
            Assert.AreEqual("Break on All Errors", (string)result.ErrorTrapping);
            CollectionAssert.AreEqual(new[] { "Break on All Errors", "Break in Class Module", "Break on Unhandled Errors" }, (string[])result.Choices);
            Assert.IsTrue((bool)result.DialogClosed);
            fake = DebugFake();
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadDebugOptions(fake));
            Assert.AreEqual(20, fake.ClosePolls);
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeDebugWindowsParsingTests
    {
        [TestMethod]
        public void FrenchAndEnglishLocalRowsExposeExpressionValueAndType()
        {
            dynamic french = VbeDebugWindows.ParseDebugRow("Expression compteur Valeur 42 Type Integer", new[] { "compteur" });
            Assert.AreEqual("compteur", (string)french.Expression);
            Assert.AreEqual("42", (string)french.Value);
            Assert.AreEqual("Integer", (string)french.Type);
            Assert.IsTrue((bool)french.Parsed);
            Assert.AreEqual(0, (int)french.Depth);
            dynamic english = VbeDebugWindows.ParseDebugRow("Expression obj Value {Object} Type Worksheet", new[] { "obj", "Name" });
            Assert.AreEqual("obj", (string)english.Expression);
            Assert.AreEqual("{Object}", (string)english.Value);
            Assert.AreEqual("Worksheet", (string)english.Type);
            Assert.AreEqual(1, (int)english.Depth);
            CollectionAssert.AreEqual(new[] { "obj", "Name" }, (string[])english.PathSegments);
        }

        [TestMethod]
        public void FrenchAndEnglishWatchRowsExposeContext()
        {
            dynamic french = VbeDebugWindows.ParseDebugRow("counter Valeur 3 Type Long Contexte VBAProject.Module1.Main", new[] { "counter" });
            Assert.AreEqual("counter", (string)french.Expression);
            Assert.AreEqual("3", (string)french.Value);
            Assert.AreEqual("VBAProject.Module1.Main", (string)french.Context);
            Assert.IsTrue((bool)french.Parsed);
            dynamic english = VbeDebugWindows.ParseDebugRow("counter Value 4 Type Long Context VBAProject.Module1.Main", new[] { "counter" });
            Assert.AreEqual("counter", (string)english.Expression);
            Assert.AreEqual("4", (string)english.Value);
            Assert.AreEqual("VBAProject.Module1.Main", (string)english.Context);
        }

        [TestMethod]
        public void NoVariablesPlaceholderIsOmittedButOtherUnparsedRowsRemainVisible()
        {
            Assert.IsNull(VbeDebugWindows.ParseDebugRow("Expression  Valeur Aucune variable Type ", new string[0]));
            Assert.IsNull(VbeDebugWindows.ParseDebugRow("Expression  Value No variables Type ", new string[0]));
            dynamic unknown = VbeDebugWindows.ParseDebugRow("Provider-specific row", null);
            Assert.IsFalse((bool)unknown.Parsed);
            Assert.AreEqual("Provider-specific row", (string)unknown.Raw);
            Assert.IsNull((string)unknown.Expression);
            Assert.AreEqual(0, ((string[])unknown.PathSegments).Length);
        }

        [TestMethod]
        public void ItemPathSegmentUnderstandsLocalAndWatchRows()
        {
            Assert.AreEqual("child", VbeDebugWindows.ParseItemPathSegment("Expression child Valeur 7 Type Long"));
            Assert.AreEqual("child", VbeDebugWindows.ParseItemPathSegment("Expression child Value 7 Type Long"));
            Assert.AreEqual("counter", VbeDebugWindows.ParseItemPathSegment("counter Value 4 Type Long Context VBAProject.Module1.Main"));
            Assert.IsNull(VbeDebugWindows.ParseItemPathSegment("Unreadable UIA row"));
            Assert.IsNull(VbeDebugWindows.ParseItemPathSegment(null));
        }

        [TestMethod]
        public void WatchContextParserRecognizesLocalizedSuffixOnly()
        {
            Assert.AreEqual("VBAProject.Module1.Main", VbeDebugWindows.ParseWatchContext("counter Valeur 3 Type Long Contexte VBAProject.Module1.Main"));
            Assert.AreEqual("VBAProject.Module1.Main", VbeDebugWindows.ParseWatchContext("counter Value 3 Type Long Context VBAProject.Module1.Main"));
            Assert.IsNull(VbeDebugWindows.ParseWatchContext("counter Value 3 Type Long"));
            Assert.IsNull(VbeDebugWindows.ParseWatchContext(null));
        }

        [TestMethod]
        public void OnlyKnownVbaDiagnosticPrefixesMayAuthorizeButtonResponse()
        {
            foreach (var message in new[]
            {
                "Erreur d'exécution '9'",
                "Run-time error '9'",
                "Erreur de compilation: Syntaxe",
                "Compile error: Syntax error"
            }

            )
                Assert.IsTrue(VbeDebugWindows.IsRecognizedDiagnostic(message), message);
            foreach (var message in new[]
            {
                null,
                "",
                "Other application error",
                "Note: Compile error",
                "Windows Security"
            }

            )
                Assert.IsFalse(VbeDebugWindows.IsRecognizedDiagnostic(message), message);
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeDebugWindowsSignatureProbeTests
    {
        [TestMethod]
        public void ReadSignatureRequiresDialogAndCancelControl()
        {
            var fake = new SignatureFake
            {
                Open = false
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadSignatureDialog("Book", fake));
            Assert.AreEqual(60, fake.Pauses);
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadSignatureDialog("Book", fake));
            Assert.AreEqual(1, fake.Closes);
        }

        [TestMethod]
        public void ReadSignatureReturnsLabelsAndClosesByCancel()
        {
            var fake = new SignatureFake
            {
                CloseAfterCancel = true
            };
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 1, Role = 41, Name = "[No certificate]" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 2, Role = 41, Name = "Certificate name" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 3, Role = 41, Name = "The VBA project is currently signed as" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 4, Role = 41, Name = "Cert A" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 5, Role = 41, Name = "Certificate name" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 6, Role = 41, Name = "Sign as" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 7, Role = 43, Name = "Cancel" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 8, Role = 99, Name = "ignored" });
            dynamic result = VbeDebugWindows.ReadSignatureDialog("Book", fake);
            Assert.AreEqual("Book", (string)result.Project);
            Assert.AreEqual("[No certificate]", (string)result.CurrentCertificate);
            Assert.AreEqual("Cert A", (string)result.SignAsCertificate);
            Assert.AreEqual(7, fake.CancelIndex);
            Assert.AreEqual(0, fake.Closes);
            Assert.IsTrue((bool)result.DialogClosed);
        }

        [TestMethod]
        public void ReadSignatureClosesOwnDialogWhenAccessibilityFailsAndRejectsStuckDialog()
        {
            var fake = new SignatureFake
            {
                FailChildren = true
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadSignatureDialog("Book", fake));
            Assert.AreEqual(1, fake.Closes);
            fake = new SignatureFake();
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 1, Role = 43, Name = "Annuler" });
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadSignatureDialog("Book", fake));
            Assert.AreEqual(20, fake.ClosePolls);
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeDebugWindowsTests
    {
        [TestMethod]
        public void ImmediateCommandRejectsMultilineControlAndOversizedInputBeforeNativeUi()
        {
            foreach (string invalid in new[]
            {
                null,
                "",
                "   ",
                "Debug.Print 1\r\nDebug.Print 2",
                "Debug.Print 1\0",
                "Debug.Print " + (char)1,
                new string ('x', 2049)
            }

            )
            {
                var error = Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ExecuteImmediate(invalid));
                StringAssert.Contains(error.Message, "one nonempty printable line");
            }
        }

        [TestMethod]
        public void DebugTreeMutationRejectsMalformedTargetBeforeNativeUi()
        {
            var invalid = new[]
            {
                new Request
                {
                    Pane = "locals",
                    Action = "expand"
                },
                new Request
                {
                    Pane = "locals",
                    Action = "expand",
                    PathSegments = new string[0]
                },
                new Request
                {
                    Pane = "locals",
                    Action = "expand",
                    PathSegments = new[]
                    {
                        "root",
                        " "
                    }
                },
                new Request
                {
                    Pane = "locals",
                    Action = "expand",
                    PathSegments = new string[17]
                },
                new Request
                {
                    Pane = "immediate",
                    Action = "expand",
                    PathSegments = new[]
                    {
                        "root"
                    }
                },
                new Request
                {
                    Pane = "watches",
                    Action = "delete",
                    PathSegments = new[]
                    {
                        "root"
                    }
                }
            };
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ChangeDebugItem(null));
            foreach (var request in invalid)
                Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ChangeDebugItem(request));
        }

        [TestMethod]
        public void DiagnosticResponseRequiresExactMessageAndButtonBeforeNativeUi()
        {
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.RespondDebugDialog(null));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.RespondDebugDialog(new Request { Diagnostic = " ", Button = "OK" }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.RespondDebugDialog(new Request { Diagnostic = "Compile error", Button = "" }));
        }

        [TestMethod]
        public void WatchSelectionRequiresExpressionAndContextBeforeNativeUi()
        {
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectWatch(null));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectWatch(new Request { Expression = "counter", Context = " " }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectWatch(new Request { Expression = "", Context = "Project.Module.Procedure" }));
        }

        [TestMethod]
        public void SignaturePlaceholderDetectionAcceptsOnlyKnownLocalizedLabels()
        {
            var method = typeof(VbeDebugWindows).GetMethod("IsNoCertificate", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method);
            foreach (var label in new[]
            {
                "[Aucun certificat]",
                "[NO CERTIFICATE]",
                "[None]"
            }

            )
                Assert.AreEqual(true, method.Invoke(null, new object[] { label }));
            foreach (var label in new[]
            {
                null,
                "",
                "None",
                "[My certificate]"
            }

            )
                Assert.AreEqual(false, method.Invoke(null, new object[] { label }));
        }

        [TestMethod]
        public void CertificateNameParserRequiresAdjacentHeading()
        {
            var method = typeof(VbeDebugWindows).GetMethod("CertificateBeforeHeading", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method);
            var headings = new[]
            {
                "Certificat actuel",
                "Current certificate"
            };
            var valid = new List<string>
            {
                "header",
                "My certificate",
                "Nom du certificat :",
                "Certificat actuel"
            };
            Assert.AreEqual("My certificate", method.Invoke(null, new object[] { valid, headings }));
            var invalid = new List<string>
            {
                "header",
                "My certificate",
                "Unrelated label",
                "Certificat actuel"
            };
            Assert.IsNull(method.Invoke(null, new object[] { invalid, headings }));
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeDebugWindowsWatchProbeTests
    {
        [TestMethod]
        public void AddWatchRequiresDialogAndExactContext()
        {
            var fake = new WatchFake
            {
                Open = false
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            Assert.AreEqual(60, fake.Pauses);
            fake.Open = true;
            fake.Texts[4858] = "Other";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            Assert.AreEqual(1, fake.Closes);
            fake.Open = true;
            fake.Texts[4858] = "Book";
            fake.Texts[4856] = "Other";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
        }

        [TestMethod]
        public void AddWatchChecksControlsTypeEditAndOk()
        {
            var fake = new WatchFake();
            fake.Missing = 4853;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            fake.Missing = 4851;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            fake.Missing = 0;
            fake.Check = false;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            fake.Check = true;
            fake.ReplaceEcho = false;
            fake.Texts[4853] = "";
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            fake.ReplaceEcho = true;
            fake.RefuseClick = 1;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
        }

        [TestMethod]
        public void AddWatchReportsNativeRejectionTimeoutAndPendingOrVisibleReadback()
        {
            var fake = new WatchFake
            {
                ErrorOpen = true,
                CloseAfterOk = false
            };
            var rejection = Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            StringAssert.Contains(rejection.Message, "invalid expression");
            Assert.AreEqual(2, fake.Closes);
            fake = new WatchFake
            {
                CloseAfterOk = false
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), fake));
            Assert.AreEqual(1, fake.Closes);
            fake = new WatchFake();
            dynamic pending = VbeDebugWindows.CompleteAddWatch(Add(), fake);
            Assert.AreEqual("Pending", (string)pending.Verification);
            Assert.AreEqual("break_when_true", (string)pending.WatchType);
            fake = new WatchFake
            {
                Root = new IntPtr(4),
                PaneHandle = new IntPtr(5)
            };
            dynamic visible = VbeDebugWindows.CompleteAddWatch(Add(), fake);
            Assert.AreEqual("ReadbackAvailable", (string)visible.Verification);
            Assert.AreEqual(new IntPtr(5), fake.ListHandle);
        }

        [TestMethod]
        public void EditWatchRefusesChangedSelectionAndNativeFailures()
        {
            var fake = new WatchFake
            {
                Open = false
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake.Open = true;
            fake.Texts[4853] = "other";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake.Texts[4853] = "x";
            fake.Texts[4857] = "Other";
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake.Texts[4857] = "Module1";
            fake.Missing = 4852;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake.Missing = 0;
            fake.Check = false;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake.Check = true;
            fake.ReplaceEcho = false;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake.ReplaceEcho = true;
            fake.RefuseClick = 1;
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
        }

        [TestMethod]
        public void EditWatchDistinguishesRejectionPendingAndVerifiedReadback()
        {
            var fake = new WatchFake
            {
                ErrorOpen = true,
                CloseAfterOk = false
            };
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake)).Message, "invalid expression");
            fake = new WatchFake
            {
                CloseAfterOk = false
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), fake));
            fake = new WatchFake();
            dynamic hidden = VbeDebugWindows.CompleteEditWatch(Edit(), fake);
            Assert.IsTrue((bool)hidden.VerificationPending);
            fake = new WatchFake
            {
                Root = new IntPtr(4),
                PaneHandle = new IntPtr(5)
            };
            dynamic pending = VbeDebugWindows.CompleteEditWatch(Edit(), fake);
            Assert.AreEqual("Pending", (string)pending.Verification);
            Assert.AreEqual(40, fake.WatchReadAttempts);
            fake = new WatchFake
            {
                Root = new IntPtr(4),
                PaneHandle = new IntPtr(5),
                NewMatches = 1
            };
            dynamic verified = VbeDebugWindows.CompleteEditWatch(Edit(), fake);
            Assert.AreEqual("ReadbackVerified", (string)verified.Verification);
            Assert.IsFalse((bool)verified.VerificationPending);
            fake = new WatchFake
            {
                Root = new IntPtr(4),
                PaneHandle = new IntPtr(5),
                NewMatches = 1,
                OldMatches = 1
            };
            dynamic oldStillPresent = VbeDebugWindows.CompleteEditWatch(Edit(), fake);
            Assert.AreEqual("Pending", (string)oldStillPresent.Verification);
        }

        [TestMethod]
        public void QuickWatchRequiresDialogControlsExpressionAndExactContext()
        {
            var request = new Request
            {
                Project = "Book",
                Module = "Module1",
                Procedure = "Run",
                Expression = "x"
            };
            var fake = new WatchFake
            {
                Open = false
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteQuickWatch(request, fake));
            Assert.AreEqual(60, fake.Pauses);
            foreach (int missing in new[]
            {
                4751,
                4752,
                4753,
                2
            }

            )
            {
                fake = QuickFake();
                fake.Missing = missing;
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteQuickWatch(request, fake));
                Assert.AreEqual(1, fake.Closes);
            }

            fake = QuickFake();
            fake.Texts[4751] = "other";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteQuickWatch(request, fake));
            fake = QuickFake();
            fake.Texts[4753] = "Other.Module1.Run";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteQuickWatch(request, fake));
            fake = QuickFake();
            fake.Texts[4753] = "Book.Module1.Other";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteQuickWatch(request, fake));
        }

        [TestMethod]
        public void QuickWatchReadbackIncludesDisplayedValueAndAlwaysClosesOwnDialog()
        {
            var fake = QuickFake();
            dynamic result = VbeDebugWindows.CompleteQuickWatch(new Request { Project = "Book", Module = "Module1", Expression = "x" }, fake);
            Assert.AreEqual("42", (string)result.Value);
            Assert.AreEqual("Book.Module1.Run", (string)result.Context);
            Assert.AreEqual("NativeDialogReadback", (string)result.Verification);
            Assert.AreEqual(1, fake.Closes);
        }

        [TestMethod]
        public void SelectWatchRequiresUniqueVisibleSelectableNativeRow()
        {
            var request = new Request
            {
                Expression = "x",
                Context = "Book.Module1.Run"
            };
            var fake = new WatchFake();
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectWatch(null, fake));
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SelectWatch(request, fake));
            fake.Root = new IntPtr(4);
            fake.PaneHandle = new IntPtr(5);
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SelectWatch(request, fake));
            fake.OldMatches = 2;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SelectWatch(request, fake));
            fake.OldMatches = 1;
            fake.SelectSucceeds = false;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SelectWatch(request, fake));
            fake.SelectSucceeds = true;
            dynamic selected = VbeDebugWindows.SelectWatch(request, fake);
            Assert.IsTrue((bool)selected.Selected);
            Assert.AreEqual(2, fake.SelectAttempts);
        }

        [TestMethod]
        public void VerifyWatchRemovedDistinguishesHiddenPendingTimeoutAndObservedAbsence()
        {
            var request = new Request
            {
                Expression = "x",
                Context = "Book.Module1.Run"
            };
            var fake = new WatchFake();
            dynamic hidden = VbeDebugWindows.VerifyWatchRemoved(request, fake);
            Assert.IsTrue((bool)hidden.VerificationPending);
            Assert.IsFalse((bool)hidden.Removed);
            fake.Root = new IntPtr(4);
            fake.PaneHandle = new IntPtr(5);
            dynamic absent = VbeDebugWindows.VerifyWatchRemoved(request, fake);
            Assert.IsTrue((bool)absent.Removed);
            fake.OldMatches = 1;
            dynamic pending = VbeDebugWindows.VerifyWatchRemoved(request, fake);
            Assert.IsTrue((bool)pending.VerificationPending);
            Assert.AreEqual(20, fake.Pauses);
            fake.DisappearAfter = fake.WatchReadAttempts + 3;
            dynamic removed = VbeDebugWindows.VerifyWatchRemoved(request, fake);
            Assert.IsTrue((bool)removed.Removed);
            Assert.IsFalse((bool)removed.VerificationPending);
        }
    }
}

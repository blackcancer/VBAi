namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    [DoNotParallelize]
    public sealed partial class VbeDebugWindowsSystemTests
    {
        [TestMethod]
        public void NativeWindowDiscoveryUsesProcessClassVisibilityAndExactLocalizedTitles()
        {
            using (var scene = new SystemScene())
            {
                var foreign = scene.Add("Digital Signature"); foreign.ProcessId = 0;
                scene.Add("Digital Signature", "unrelated");
                var hidden = scene.Add("Digital Signature"); hidden.Visible = false;
                scene.Add("Unrelated");
                foreach (var title in new[] { "Digital Signature", "Signature numérique", "Signature électronique" })
                    foreach (var kind in new[] { "#32770", "bosa_sdm_dialog" })
                    {
                        var match = scene.Add(title, kind);
                        Assert.AreEqual(match.Handle, Call("FindSignatureDialog")); match.Visible = false;
                    }
                Assert.AreEqual(IntPtr.Zero, Call("FindSignatureDialog"));
                var wrongRoot = scene.Add("VBE", "wndclass_desked_gsk"); wrongRoot.ProcessId = 0;
                var invisibleRoot = scene.Add("VBE", "wndclass_desked_gsk"); invisibleRoot.Visible = false;
                var root = scene.Add("Editor", "wndclass_desked_gsk");
                var invisiblePane = scene.Add("Locals", "VbaWindow", root); invisiblePane.Visible = false;
                scene.Add("Locals", "Other", root);
                var locals = scene.Add("Variables locales", "VbaWindow", root);
                var watches = scene.Add("Espions", "VbaWindow", root);
                Assert.AreEqual(root.Handle, Call("FindVbeRoot"));
                var children = (System.Collections.Generic.List<IntPtr>)Call("ChildWindows", root.Handle);
                CollectionAssert.AreEqual(new[] { locals.Handle, watches.Handle }, children);
                Assert.AreEqual(watches.Handle, Call("FindPane", children, new[] { "Watches", "Espions" }));
                Assert.AreEqual(IntPtr.Zero, Call("FindPane", children, new[] { "Missing" }));
                var compile = scene.Add("Microsoft Visual Basic for Applications");
                Assert.AreEqual(compile.Handle, Call("FindDialog", (object)new[] { "Other", compile.Text }));
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.EnsureNoCompileDialog());
                compile.Visible = false; VbeDebugWindows.EnsureNoCompileDialog();
                var options = scene.Add("Options");
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.EnsureNoDebugOptionsDialog());
                options.Visible = false; VbeDebugWindows.EnsureNoDebugOptionsDialog();
                var signature = scene.Add("Signature numérique");
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.EnsureNoSignatureDialog());
                signature.Visible = false; VbeDebugWindows.EnsureNoSignatureDialog();
            }
        }

        [TestMethod]
        public void NativeDialogCaptureReadsVisibleControlsAndNativeWatchMessages()
        {
            using (var scene = new SystemScene())
            {
                var dialog = scene.Add("Microsoft Visual Basic pour Applications");
                var message = scene.Add("Erreur de compilation: Expected expression", "Static", dialog);
                var ok = scene.Add("OK", "Button", dialog, 1);
                var cancel = scene.Add("Cancel", "Button", dialog, 2);
                var hidden = scene.Add("Hidden", "Static", dialog); hidden.Visible = false;
                var native = Native<VbeDebugWindows.INativeProbe>("NativeProbe");
                Assert.AreEqual(dialog.Handle, native.Dialog(dialog.Text));
                Assert.AreEqual(4, native.DialogControls(dialog.Handle).Count);
                dynamic read = VbeDebugWindows.ReadDebugDialog();
                Assert.AreEqual(message.Text, (string)read.Diagnostic);
                CollectionAssert.AreEqual(new[] { "OK", "Cancel" }, (string[])read.Buttons);
                Assert.IsTrue(native.Click(ok.Handle)); Assert.IsTrue(native.Visible(dialog.Handle)); native.Pause(1);
                scene.OnMessage = (window, messageId) => { if (window == cancel) dialog.Visible = false; };
                Call("CloseDialog", dialog.Handle); Assert.IsFalse(dialog.Visible);
                Call("CloseDialog", new IntPtr(999));
                var watch = Native<VbeDebugWindows.IWatchProbe>("NativeWatchProbe");
                dialog.Visible = true;
                Assert.AreEqual(dialog.Handle, watch.Dialog(dialog.Text)); Assert.AreEqual(ok.Handle, watch.Item(dialog.Handle, 1));
                Assert.AreEqual(message.Text, watch.Text(message.Handle)); Assert.IsTrue(watch.Click(ok.Handle));
                scene.IntegerMessageResult = IntPtr.Zero; Assert.IsFalse(watch.Checked(ok.Handle));
                scene.IntegerMessageResult = new IntPtr(1); Assert.IsTrue(watch.Checked(ok.Handle));
                watch.Replace(ok.Handle, "new expression"); Assert.AreEqual("new expression", scene.Replacements.Single().Item2);
                watch.Pause(1); watch.Close(dialog.Handle);
                Assert.AreEqual(IntPtr.Zero, watch.VbeRoot()); Assert.AreEqual(0, watch.Children(IntPtr.Zero).Count);
                Assert.AreEqual(IntPtr.Zero, watch.Pane(new IntPtr[0], "Espions"));
                dynamic absent = watch.List(IntPtr.Zero); Assert.IsFalse((bool)absent.Available);
                var immediate = Native<VbeDebugWindows.IImmediateProbe>("NativeImmediateProbe");
                Assert.AreEqual(IntPtr.Zero, immediate.VbeRoot()); Assert.AreEqual(0, immediate.Children(IntPtr.Zero).Count);
                Assert.AreEqual(IntPtr.Zero, immediate.Pane(new IntPtr[0], "Immediate"));
                Assert.IsTrue(immediate.PostChar(ok.Handle, 'x')); Assert.IsTrue(immediate.PostEnter(ok.Handle));
                scene.PostSucceeds = false; Assert.IsFalse(immediate.PostEnter(ok.Handle)); immediate.Pause(1);
                Assert.AreEqual(0x0102, scene.Messages.First(item => item.Item3 == new IntPtr('x')).Item2);
                var options = Native<VbeDebugWindows.IOptionsProbe>("NativeOptionsProbe");
                Assert.AreEqual(IntPtr.Zero, options.Dialog()); options.Pause(1); options.Close(dialog.Handle);
            }
        }

        [TestMethod]
        public void SignatureMsaaReadersBoundChildrenSkipFailuresAndUseTheExactButton()
        {
            using (var scene = new SystemScene())
            {
                var dialog = scene.Add("Digital Signature");
                var root = Signature(dialog, "[None]", "[None]");
                root.Children.AddRange(new[] { Label(" "), new AccessibleNode { Label = "other", NativeRole = AccessibleRole.List },
                    new AccessibleNode { FailName = true, NativeRole = AccessibleRole.StaticText },
                    new AccessibleNode { FailRole = true } });
                var labels = (System.Collections.Generic.List<string>)Call("SignatureLabels", root);
                Assert.AreEqual(6, labels.Count);
                int clicked = 0; root.Children.Add(Button("Confirm", () => clicked++));
                Call("InvokeSignatureButton", root, new[] { "Confirm" }); Assert.AreEqual(1, clicked);
                var missing = Assert.ThrowsException<TargetInvocationException>(() => Call("InvokeSignatureButton", root, new[] { "Missing" }));
                Assert.IsInstanceOfType(missing.InnerException, typeof(InvalidOperationException));
                var native = Native<VbeDebugWindows.ISignatureProbe>("NativeSignatureProbe");
                Assert.AreEqual(dialog.Handle, native.Dialog());
                var children = native.Children(dialog.Handle); Assert.AreEqual(root.Children.Count - 2, children.Count);
                var cancel = children.Single(child => child.Name == "Cancel"); native.Cancel(dialog.Handle, cancel.Index); Assert.IsFalse(dialog.Visible);
                native.Close(dialog.Handle); native.Pause(1);
                dialog.Visible = true;
                dynamic read = VbeDebugWindows.ReadSignatureDialog("Project"); Assert.IsTrue((bool)read.DialogClosed);
                Assert.AreEqual("[None]", (string)read.CurrentCertificate);
                dialog.Visible = true;
                root.Children.Clear(); root.Children.Add(Button("Unrelated"));
                scene.OnMessage = (window, messageId) => { if (window == dialog && messageId == 0x0010) dialog.Visible = false; };
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadSignatureDialog("Project")); Assert.IsFalse(dialog.Visible);
                root.Children.Clear(); for (int i = 0; i < 70; i++) root.Children.Add(Label("Label" + i));
                Assert.AreEqual(64, ((System.Collections.Generic.List<string>)Call("SignatureLabels", root)).Count);
                Assert.AreEqual(64, native.Children(dialog.Handle).Count);
            }
        }

        [TestMethod]
        public void SignatureMsaaRefusesNativeErrorsAndObjectsWithoutTheAccessibleInterface()
        {
            using (var scene = new SystemScene())
            {
                var dialog = scene.Add("Digital Signature");
                Signature(dialog, "[None]", "[None]");
                foreach (bool errorHResult in new[] { true, false })
                {
                    scene.AccessibilityHResult = errorHResult ? unchecked((int)0x80004005) : 0;
                    scene.OverrideAccessibility = true; scene.AccessibilityOverride = new object();
                    var error = Assert.ThrowsException<TargetInvocationException>(() => Call("SignatureAccessible", dialog.Handle));
                    Assert.IsInstanceOfType(error.InnerException, typeof(COMException));
                }
            }
        }

        [TestMethod]
        public void ProjectSignatureUsesExistingReadbackOnlyForTheSelectedNamedCertificate()
        {
            foreach (var state in new[] { new[] { "[None]", "Certificate" }, new[] { "Certificate", "[None]" } })
                using (var scene = new SystemScene())
                {
                    var dialog = scene.Add("Digital Signature"); var root = Signature(dialog, state[0], state[1]);
                    int chooseClicks = 0; root.Children[6].OnInvoke = () => chooseClicks++;
                    dynamic result = VbeDebugWindows.CompleteProjectSignature("Project", "thumbprint", "Certificate", state[0] != "[None]");
                    Assert.IsTrue((bool)result.SignatureAssigned); Assert.AreEqual("NativeVbeExistingCertificate", (string)result.SelectionSource);
                    Assert.AreEqual(0, chooseClicks); Assert.IsFalse(dialog.Visible);
                    Assert.AreEqual("thumbprint", (string)result.CertificateThumbprint);
                }
        }

        [TestMethod]
        public void ProjectSignatureWaitsForPickerReadbackAndRejectsAChangedOrClosedDialog()
        {
            foreach (var outcome in new[] { "selected", "wrong", "closed", "timeout", "stuck" })
                using (var scene = new SystemScene())
                {
                    var dialog = scene.Add("Digital Signature"); var root = Signature(dialog, "[None]", "[None]");
                    bool picker = false; root.Children[6].OnInvoke = () => picker = true;
                    scene.OnPause = count => {
                        if (!picker || count < 5) return;
                        if (outcome == "selected" || outcome == "stuck") root.Children[3].Label = "Certificate";
                        else if (outcome == "wrong") root.Children[3].Label = "Other certificate";
                        else if (outcome == "closed") dialog.Visible = false;
                    };
                    if (outcome == "stuck") root.Children[7].OnInvoke = null;
                    scene.OnMessage = (window, messageId) => { if (window == dialog && messageId == 0x0010) dialog.Visible = false; };
                    if (outcome == "selected")
                    {
                        dynamic result = VbeDebugWindows.CompleteProjectSignature("Project", "thumbprint", "Certificate", false);
                        Assert.AreEqual("WindowsCertificatePicker", (string)result.SelectionSource); Assert.IsTrue((bool)result.SignatureAssigned);
                    }
                    else if (outcome == "timeout") Assert.ThrowsException<TimeoutException>(() => VbeDebugWindows.CompleteProjectSignature("Project", "thumbprint", "Certificate", false));
                    else Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteProjectSignature("Project", "thumbprint", "Certificate", false));
                    Assert.IsTrue(picker); Assert.IsFalse(dialog.Visible);
                }
        }

        [TestMethod]
        public void ProjectSignatureNeverOverwritesAnExistingOrDifferentCertificate()
        {
            foreach (var state in new[] { new[] { "Other certificate", "[None]", "false" }, new[] { "[None]", "Other certificate", "true" },
                new[] { (string)null, "[None]", "false" }, new[] { "[None]", (string)null, "true" } })
                using (var scene = new SystemScene())
                {
                    var dialog = scene.Add("Digital Signature"); var root = Signature(dialog, state[0], state[1]);
                    int buttonClicks = 0; root.Children[6].OnInvoke = root.Children[7].OnInvoke = () => buttonClicks++;
                    scene.OnMessage = (window, messageId) => { if (window == dialog) dialog.Visible = false; };
                    Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteProjectSignature("Project", "thumbprint", "Certificate", bool.Parse(state[2])));
                    Assert.AreEqual(0, buttonClicks); Assert.IsFalse(dialog.Visible);
                }
            using (var scene = new SystemScene())
            {
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteProjectSignature("Project", "thumbprint", "Certificate", true));
                Assert.AreEqual(60, scene.Pauses);
                Assert.AreEqual(0, scene.Messages.Count);
            }
        }

        [TestMethod]
        public void CallStackReadsMsaaFramesAndPreservesAUserOwnedDialog()
        {
            foreach (bool available in new[] { true, false })
                using (var scene = new SystemScene())
                {
                    var dialog = scene.Add("Call Stack"); var root = new AccessibleNode();
                    root.Children.AddRange(new[] { Label("Project.Module.TryMe"), Label("Project.Module.Main") });
                    var list = scene.Add("Frames", "ListBox", dialog); list.Accessible = root;
                    scene.Add("OK", "Button", dialog, 1);
                    if (!available) list.Class = "Other";
                    dynamic result = Call("ReadCallStack", IntPtr.Zero);
                    Assert.IsTrue((bool)result.Available); Assert.IsTrue(dialog.Visible); Assert.AreEqual(0, scene.Messages.Count);
                    if (available) { CollectionAssert.AreEqual(new[] { "Project.Module.TryMe", "Project.Module.Main" }, (string[])result.Frames); Assert.IsNull((string)result.Error); }
                    else Assert.IsNotNull((string)result.Error);
                }
        }

        [TestMethod]
        public void CallStackOwnDialogClosesAfterReadsErrorsAndNativeOpeningFailures()
        {
            foreach (var outcome in new[] { "success", "missingList", "badMsaa", "invalidMsaa", "timeout", "noButton", "postFailure" })
                using (var scene = new SystemScene())
                {
                    var locals = scene.Add("Locals", "VbaWindow"); scene.Add("Other", "Static", locals);
                    var trigger = scene.Add("Call Stack", "Button", locals, 4601);
                    var hidden = scene.Add("Call Stack", "Button", locals, 4601); hidden.Visible = false;
                    if (outcome == "noButton") trigger.ControlId = 1;
                    if (outcome == "postFailure") scene.PostSucceeds = false;
                    SystemWindow opened = null;
                    scene.OnMessage = (window, messageId) => {
                        if (window == trigger && outcome != "timeout" && outcome != "postFailure")
                        {
                            opened = scene.Add("Pile des appels");
                            scene.Add("Other", "Static", opened);
                            var list = scene.Add("Frames", outcome == "missingList" ? "Other" : "ListBox", opened);
                            var access = new AccessibleNode(); access.Children.Add(Label("TryMe")); list.Accessible = access;
                            scene.Add("OK", "Button", opened, 1); scene.Add("Cancel", "Button", opened, 2);
                            if (outcome == "badMsaa") scene.AccessibilityHResult = -1;
                            if (outcome == "invalidMsaa") { scene.OverrideAccessibility = true; scene.AccessibilityOverride = new object(); }
                        }
                        else if (window != null && window.Parent == opened?.Handle && window.ControlId == 2) opened.Visible = false;
                    };
                    dynamic result = Call("ReadCallStack", locals.Handle);
                    if (outcome == "success") { Assert.IsTrue((bool)result.Available); Assert.IsNull((string)result.Error); Assert.IsFalse(opened.Visible); }
                    else Assert.IsNotNull((string)result.Error);
                    if (opened != null) Assert.IsFalse(opened.Visible);
                }
            using (var scene = new SystemScene())
            {
                dynamic result = Call("ReadCallStack", IntPtr.Zero); Assert.IsFalse((bool)result.Available);
                var foreign = scene.Add("Call Stack"); foreign.ProcessId = 0;
                var invisible = scene.Add("Call Stack"); invisible.Visible = false;
                scene.Add("Call Stack", "Other"); scene.Add("Other");
                Assert.AreEqual(IntPtr.Zero, Call("FindCallStackDialog"));
            }
        }

        [TestMethod]
        public void PublicNativeWrappersRejectAbsentUiBeforeAnyAction()
        {
            using (var scene = new SystemScene())
            {
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.Capture(true));
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeDebugItem(new Request { Pane = "locals", Action = "expand", PathSegments = new[] { "x" } }));
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ExecuteImmediate("Debug.Print 1"));
                var root = scene.Add("VBE", "wndclass_desked_gsk");
                dynamic capture = VbeDebugWindows.Capture(true); Assert.IsNotNull((object)capture.CallStack);
                foreach (var pane in new[] { "locals", "watches" })
                    Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeDebugItem(new Request { Pane = pane, Action = "expand", PathSegments = new[] { "x" } }));
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ExecuteImmediate("Debug.Print 1"));
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(new Request()));
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(new Request()));
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteQuickWatch(new Request()));
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadDebugOptions());
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadVbeOptions());
                dynamic dialog = VbeDebugWindows.ReadDebugDialog(); Assert.IsFalse((bool)dialog.Available);
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadSignatureDialog("Project"));
            }
        }

        [TestMethod]
        public void UiaListsReadValuesHierarchiesPlaceholdersAndUnparsedRows()
        {
            var root = new AutomationNode { Name = "Locals", Kind = System.Windows.Automation.ControlType.List };
            root.Add(new AutomationNode { Name = "Expression x Value 42 Type Long" });
            root.Add(new AutomationNode { Name = "Expression x Value 42 Type Long" });
            root.Add(new AutomationNode { Name = "Expression x Value 43 Type Long" });
            root.Add(new AutomationNode { Name = "Expression  Value No variables Type " });
            root.Add(new AutomationNode { Name = "unparsed native row" });
            root.Add(new AutomationNode { Name = "Expression display Value old Type String", Text = "Expression actual Value new Type String" }.With(System.Windows.Automation.ValuePattern.Pattern));
            var parent = root.Add(new AutomationNode { Name = "Expression container Value {...} Type Collection" });
            parent.Add(new AutomationNode { Name = "Expression child Value 1 Type Integer" });
            parent.Add(new AutomationNode { Name = "Expression x Value 42 Type Long" });
            using (var host = new AutomationHost(root))
            {
                dynamic result = Call("ReadList", host.Handle);
                Assert.IsNull((string)result.Error);
                var rows = ((IEnumerable)result.Items).Cast<object>().ToArray(); Assert.AreEqual(7, rows.Length);
                Assert.AreEqual(1, (int)result.DuplicateRowsOmitted);
                var xRows = rows.Where(row => (string)((dynamic)row).Expression == "x").ToArray();
                Assert.AreEqual(3, xRows.Length, "Changed values and different ancestry must remain distinct observations.");
                Assert.AreEqual(2, xRows.Count(row => (string)((dynamic)row).Value == "42"));
                Assert.AreEqual(1, xRows.Count(row => (int)((dynamic)row).Depth == 1));
                dynamic value = rows.Single(row => (string)((dynamic)row).Expression == "actual");
                Assert.AreEqual("new", (string)value.Value);
                dynamic child = rows.Single(row => (string)((dynamic)row).Expression == "child");
                CollectionAssert.AreEqual(new[] { "container", "child" }, (string[])child.PathSegments);
                Assert.AreEqual(1, (int)child.Depth);
                Assert.IsFalse((bool)((dynamic)rows.Single(row => (string)((dynamic)row).Raw == "unparsed native row")).Parsed);
            }
            using (var scene = new SystemScene())
            {
                dynamic failed = Call("ReadList", new IntPtr(999)); Assert.IsTrue((bool)failed.Available); Assert.IsNotNull((string)failed.Error);
            }
        }

        [TestMethod]
        public void DebugRowIdentityPreservesRawValuesAndFullAncestryWithoutSeparatorCollisions()
        {
            Assert.AreEqual(VbeDebugWindows.DebugRowIdentity("row", null),
                VbeDebugWindows.DebugRowIdentity("row", new string[0]));
            Assert.AreNotEqual(VbeDebugWindows.DebugRowIdentity("row", new[] { "a", "b" }),
                VbeDebugWindows.DebugRowIdentity("row", new[] { "a/b" }));
            Assert.AreNotEqual(VbeDebugWindows.DebugRowIdentity("row", new[] { "parent", "child" }),
                VbeDebugWindows.DebugRowIdentity("row", new[] { "other", "child" }));
            Assert.AreNotEqual(VbeDebugWindows.DebugRowIdentity("row", new[] { "child" }),
                VbeDebugWindows.DebugRowIdentity("changed", new[] { "child" }));
        }

        [TestMethod]
        public void UiaWatchesSelectOnlyTheExactExpressionAndContextAndVerifyRemoval()
        {
            var root = new AutomationNode { Name = "Watches", Kind = System.Windows.Automation.ControlType.List };
            root.Add(new AutomationNode { Name = "not a native watch" });
            root.Add(new AutomationNode { Name = "other Value 1 Type Long Context Project.Module" });
            root.Add(new AutomationNode { Name = "counter Value 1 Type Long Context Other.Module" });
            var selected = root.Add(new AutomationNode { Name = "counter Value 7 Type Long Context Project.Module" }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            using (var host = new AutomationHost(root))
                using (var scene = new SystemScene())
                {
                    var editor = scene.Add("Editor", "wndclass_desked_gsk");
                    var pane = scene.Add("Espions", "VbaWindow", editor); pane.Handle = host.Handle;
                    var native = Native<VbeDebugWindows.IWatchProbe>("NativeWatchProbe");
                    Assert.AreEqual(1, native.WatchMatches(host.Handle, "counter", "project.module"));
                    Assert.IsTrue(native.SelectWatchRow(host.Handle, "counter", "Project.Module"));
                    Assert.IsTrue(selected.Selected); Assert.IsTrue(selected.FocusCount > 0);
                    var request = new Request { Expression = "counter", Context = "Project.Module" };
                    dynamic result = VbeDebugWindows.SelectWatch(request); Assert.IsTrue((bool)result.Selected);
                    selected.Patterns.Clear(); Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SelectWatch(request));
                    selected.Patterns.Add(System.Windows.Automation.SelectionItemPattern.Pattern.Id);
                    dynamic stillPresent = VbeDebugWindows.VerifyWatchRemoved(request); Assert.IsTrue((bool)stillPresent.VerificationPending);
                    root.Children.Remove(selected);
                    dynamic removed = VbeDebugWindows.VerifyWatchRemoved(request); Assert.IsFalse((bool)removed.VerificationPending);
                }
        }

        [TestMethod]
        public void UiaTreeExpansionMatchesTheCanonicalPathAndWatchRootContext()
        {
            foreach (var paneName in new[] { "locals", "watches" })
            {
                var root = new AutomationNode { Kind = System.Windows.Automation.ControlType.List, Name = "Debug items" };
                string name = paneName == "locals" ? "Expression group Value {...} Type Collection" : "group Value {...} Type Collection Context Project.Module";
                var group = root.Add(new AutomationNode { Name = name, HideCollapsedChildren = true, Expanded = false }.With(System.Windows.Automation.ExpandCollapsePattern.Pattern));
                group.Add(new AutomationNode { Name = "Expression child Value 1 Type Long" });
                using (var host = new AutomationHost(root))
                    using (var scene = new SystemScene())
                    {
                        var editor = scene.Add("Editor", "wndclass_desked_gsk");
                        var pane = scene.Add(paneName == "locals" ? "Locals" : "Watch", "VbaWindow", editor); pane.Handle = host.Handle;
                        var request = new Request { Pane = paneName, Action = "expand", PathSegments = new[] { "group" }, Context = "Project.Module" };
                        dynamic expanded = VbeDebugWindows.ChangeDebugItem(request); Assert.AreEqual("Observed", (string)expanded.Verification);
                        request.Action = "collapse";
                        dynamic collapsed = VbeDebugWindows.ChangeDebugItem(request); Assert.AreEqual("Observed", (string)collapsed.Verification);
                        request.Action = "expand"; request.PathSegments = new[] { "missing" };
                        Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeDebugItem(request));
                        request.PathSegments = new[] { "group" };
                        if (paneName == "watches")
                        {
                            request.Context = "Other.Module"; Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeDebugItem(request));
                            request.Context = "Project.Module";
                        }
                        group.Patterns.Clear(); Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeDebugItem(request));
                        group.Patterns.Add(System.Windows.Automation.ExpandCollapsePattern.Pattern.Id);
                        group.Children.Clear();
                        dynamic pending = VbeDebugWindows.ChangeDebugItem(request); Assert.AreEqual("Pending", (string)pending.Verification);
                        group.Children.Add(new AutomationNode { Parent = group, Name = "Expression child Value 1 Type Long" });
                        group.HideCollapsedChildren = false; request.Action = "collapse";
                        dynamic pendingCollapse = VbeDebugWindows.ChangeDebugItem(request); Assert.IsTrue((bool)pendingCollapse.VerificationPending);
                    }
            }
        }

        [TestMethod]
        public void UiaImmediateDocumentUsesTextPatternAndPlacesTheCaretAtTheEnd()
        {
            var root = new AutomationNode { Name = "Immediate pane", Kind = System.Windows.Automation.ControlType.Pane };
            var document = root.Add(new AutomationNode { Name = "Immediate", Kind = System.Windows.Automation.ControlType.Document, Text = "previous output\r\n" }.With(System.Windows.Automation.TextPattern.Pattern));
            using (var host = new AutomationHost(root))
                using (var scene = new SystemScene())
                {
                    var editor = scene.Add("Editor", "wndclass_desked_gsk"); var pane = scene.Add("Immediate", "VbaWindow", editor); pane.Handle = host.Handle;
                    var native = Native<VbeDebugWindows.IImmediateProbe>("NativeImmediateProbe");
                    Assert.AreEqual(document.Text, native.Prepare(host.Handle)); Assert.IsTrue(document.FocusCount > 0); Assert.IsTrue(document.SelectionCount > 0);
                    Assert.AreEqual(document.Text, native.Text(host.Handle));
                    dynamic read = Call("ReadImmediate", host.Handle); Assert.IsTrue((bool)read.Available); Assert.AreEqual(document.Text, (string)read.Text);
                    scene.OnMessage = (window, message) => {
                        if (message == 0x0102) document.Text += (char)scene.Messages.Last().Item3.ToInt32();
                        else if (message == 0x0101) document.Text += "\r\n1\r\n";
                    };
                    dynamic executed = VbeDebugWindows.ExecuteImmediate("Debug.Print 1"); Assert.AreEqual("ImmediateTextChangedAfterEnter", (string)executed.Verification);
                    document.Patterns.Clear();
                    var patternError = Assert.ThrowsException<TargetInvocationException>(() => Call("ImmediateDocument", host.Handle)); Assert.IsInstanceOfType(patternError.InnerException, typeof(InvalidOperationException));
                    document.Patterns.Add(System.Windows.Automation.TextPattern.Pattern.Id);
                    root.Add(new AutomationNode { Name = "Second", Kind = System.Windows.Automation.ControlType.Document });
                    var countError = Assert.ThrowsException<TargetInvocationException>(() => Call("ImmediateDocument", host.Handle)); Assert.IsInstanceOfType(countError.InnerException, typeof(InvalidOperationException));
                    dynamic failed = Call("ReadImmediate", host.Handle); Assert.IsNotNull((string)failed.Error);
                }
        }

        [TestMethod]
        public void UiaOptionsReadNativePatternsWithoutReadingPasswordsOrDisabledControls()
        {
            var root = new AutomationNode { Name = "Options", Kind = System.Windows.Automation.ControlType.Window };
            root.Add(new AutomationNode { Name = "Editor", Kind = System.Windows.Automation.ControlType.TabItem }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            root.Add(new AutomationNode { Name = "Enabled option", Kind = System.Windows.Automation.ControlType.CheckBox, ToggleState = System.Windows.Automation.ToggleState.On }.With(System.Windows.Automation.TogglePattern.Pattern));
            root.Add(new AutomationNode { Name = "No toggle", Kind = System.Windows.Automation.ControlType.CheckBox });
            root.Add(new AutomationNode { Name = "Selected", Kind = System.Windows.Automation.ControlType.RadioButton, Selected = true }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            root.Add(new AutomationNode { Name = "List entry", Kind = System.Windows.Automation.ControlType.ListItem }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            root.Add(new AutomationNode { Name = "Edit", Kind = System.Windows.Automation.ControlType.Edit, Text = "visible text" }.With(System.Windows.Automation.ValuePattern.Pattern));
            root.Add(new AutomationNode { Name = "Password", Kind = System.Windows.Automation.ControlType.Edit, Password = true, Text = "never read" }.With(System.Windows.Automation.ValuePattern.Pattern));
            root.Add(new AutomationNode { Name = "Combo", Kind = System.Windows.Automation.ControlType.ComboBox, Text = "chosen" }.With(System.Windows.Automation.ValuePattern.Pattern));
            root.Add(new AutomationNode { Name = "Slider", Kind = System.Windows.Automation.ControlType.Slider, Number = 37 }.With(System.Windows.Automation.RangeValuePattern.Pattern));
            root.Add(new AutomationNode { Name = "No range", Kind = System.Windows.Automation.ControlType.Slider });
            root.Add(new AutomationNode { Name = "Plain text", Kind = System.Windows.Automation.ControlType.Text });
            root.Add(new AutomationNode { Name = "Disabled", Kind = System.Windows.Automation.ControlType.Edit, Enabled = false });
            root.Add(new AutomationNode { Name = "Offscreen", Kind = System.Windows.Automation.ControlType.Edit, Offscreen = true });
            root.Add(new AutomationNode { Name = "Bad value", Kind = System.Windows.Automation.ControlType.Edit, FailValue = true }.With(System.Windows.Automation.ValuePattern.Pattern));
            using (var host = new AutomationHost(root))
                using (var scene = new SystemScene())
                {
                    var options = scene.Add("Options"); options.Handle = host.Handle;
                    var cancel = scene.Add("Cancel", "Button", options, 2);
                    scene.OnMessage = (window, message) => { if (window == cancel) options.Visible = false; };
                    var native = Native<VbeDebugWindows.IOptionsProbe>("NativeOptionsProbe");
                    CollectionAssert.AreEqual(new[] { "Editor" }, native.Tabs(host.Handle).ToArray());
                    var controls = native.Controls(host.Handle, 0);
                    Assert.AreEqual("On", controls.Single(control => control.Name == "Enabled option").Value);
                    Assert.AreEqual(true, controls.Single(control => control.Name == "Selected").Value);
                    Assert.AreEqual("visible text", controls.Single(control => control.Name == "Edit").Value);
                    Assert.IsNull(controls.Single(control => control.Name == "Password").Value);
                    Assert.AreEqual("chosen", controls.Single(control => control.Name == "Combo").Value);
                    Assert.AreEqual(37d, controls.Single(control => control.Name == "Slider").Value);
                    Assert.IsNotNull(controls.Single(control => control.Name == "Bad value").Error);
                    dynamic result = VbeDebugWindows.ReadVbeOptions(); Assert.IsTrue((bool)result.DialogClosed);
                }
        }

        [TestMethod]
        public void UiaGeneralTabRequiresOneSelectableTabAndReadsEachRadioSelection()
        {
            var root = new AutomationNode { Name = "Options", Kind = System.Windows.Automation.ControlType.Window };
            var general = root.Add(new AutomationNode { Name = "General", Kind = System.Windows.Automation.ControlType.TabItem }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            root.Add(new AutomationNode { Name = "Break on All Errors", Kind = System.Windows.Automation.ControlType.RadioButton }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            root.Add(new AutomationNode { Name = "Break in Class Module", Kind = System.Windows.Automation.ControlType.RadioButton }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            root.Add(new AutomationNode { Name = "Break on Unhandled Errors", Kind = System.Windows.Automation.ControlType.RadioButton, Selected = true }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            using (var host = new AutomationHost(root))
                using (var scene = new SystemScene())
                {
                    var options = scene.Add("Options"); options.Handle = host.Handle;
                    var cancel = scene.Add("Cancel", "Button", options, 2);
                    scene.OnMessage = (window, message) => { if (window == cancel) options.Visible = false; };
                    var native = Native<VbeDebugWindows.IOptionsProbe>("NativeOptionsProbe");
                    var choices = native.ErrorChoices(host.Handle); Assert.AreEqual(3, choices.Count); Assert.AreEqual(1, choices.Count(choice => choice.Selected));
                    dynamic read = VbeDebugWindows.ReadDebugOptions(); Assert.AreEqual("Break on Unhandled Errors", (string)read.ErrorTrapping);
                    options.Visible = true;
                    root.Children.Last().Patterns.Clear(); Assert.IsFalse(native.ErrorChoices(host.Handle).Last().Readable);
                    general.Patterns.Clear(); Assert.ThrowsException<InvalidOperationException>(() => native.ErrorChoices(host.Handle));
                    general.Patterns.Add(System.Windows.Automation.SelectionItemPattern.Pattern.Id);
                    general.Name = "Other"; Assert.ThrowsException<InvalidOperationException>(() => native.ErrorChoices(host.Handle));
                    general.Name = "General";
                    native.Tabs(host.Handle); general.Patterns.Clear(); Assert.ThrowsException<InvalidOperationException>(() => native.Controls(host.Handle, 0));
                }
        }

        [TestMethod]
        public void UiaDiagnosticReaderSkipsButtonsAndReportsTheFirstActualMessage()
        {
            var root = new AutomationNode { Kind = System.Windows.Automation.ControlType.Window, Name = "Diagnostic" };
            foreach (string text in new[] { " ", "OK", "Aide", "Help" }) root.Add(new AutomationNode { Kind = System.Windows.Automation.ControlType.Text, Name = text });
            using (var host = new AutomationHost(root))
            {
                Assert.AreEqual("Native validation error.", Call("AccessibleDialogMessage", host.Handle));
                root.Add(new AutomationNode { Kind = System.Windows.Automation.ControlType.Text, Name = "Type mismatch" });
                Assert.AreEqual("Type mismatch", Call("AccessibleDialogMessage", host.Handle));
                var native = Native<VbeDebugWindows.INativeProbe>("NativeProbe"); Assert.AreEqual("Type mismatch", native.DialogMessage(host.Handle));
                var watch = Native<VbeDebugWindows.IWatchProbe>("NativeWatchProbe"); Assert.AreEqual("Type mismatch", watch.Message(host.Handle));
            }
        }

        [TestMethod]
        public void UiaOptionsCapsTheNativeCollectionAndSkipsUnavailableElements()
        {
            var root = new AutomationNode { Kind = System.Windows.Automation.ControlType.Window, Name = "Options" };
            root.Add(new AutomationNode { Name = "Editor", Kind = System.Windows.Automation.ControlType.TabItem }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            root.Add(new AutomationNode { Name = "unavailable", Kind = System.Windows.Automation.ControlType.Edit, FailName = true });
            using (var host = new AutomationHost(root, optionsDialog: true))
                using (var scene = new SystemScene())
                {
                    BindOwnedOptionsDialog(scene, host);
                    var native = Native<VbeDebugWindows.IOptionsProbe>("NativeOptionsProbe"); native.Tabs(host.Handle);
                    Assert.AreEqual(0, native.Controls(host.Handle, 0).Count);
                    root.Children.RemoveAt(1);
                    for (int i = 0; i < 2001; i++) root.Add(new AutomationNode { Name = "Value" + i, Kind = System.Windows.Automation.ControlType.Text });
                    Assert.ThrowsException<InvalidOperationException>(() => native.Controls(host.Handle, 0));
                }
        }

        [TestMethod]
        public void UiaHierarchyBoundsDeepPathsAndUsesTheTopmostWatchContext()
        {
            var root = new AutomationNode { Name = "Watches", Kind = System.Windows.Automation.ControlType.List };
            var outer = root.Add(new AutomationNode { Name = "outer Value {...} Type Collection Context Root.Module" });
            var child = outer.Add(new AutomationNode { Name = "child Value 1 Type Long Context Other.Module" });
            var deep = child;
            for (int i = 0; i < 18; i++) deep = deep.Add(new AutomationNode { Name = "Expression level" + i + " Value 1 Type Long" });
            root.Add(new AutomationNode { Name = null });
            root.Add(new AutomationNode { Name = "Expression absent Value 2 Type Long", Text = null }.With(System.Windows.Automation.ValuePattern.Pattern));
            using (var host = new AutomationHost(root))
            {
                var window = System.Windows.Automation.AutomationElement.FromHandle(host.Handle);
                var target = window.FindFirst(System.Windows.Automation.TreeScope.Descendants,
                    new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.NameProperty, child.Name));
                Assert.AreEqual("Root.Module", Call("WatchRootContext", target));
                Assert.IsNull(Call("WatchRootContext", System.Windows.Automation.AutomationElement.RootElement));
                var emptyName = window.FindFirst(System.Windows.Automation.TreeScope.Descendants,
                    new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.NameProperty, ""));
                Assert.AreEqual("", emptyName.Current.Name);
                var emptyValue = window.FindFirst(System.Windows.Automation.TreeScope.Descendants,
                    new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.NameProperty, "Expression absent Value 2 Type Long"));
                Assert.AreEqual("", ((System.Windows.Automation.ValuePattern)emptyValue.GetCurrentPattern(System.Windows.Automation.ValuePattern.Pattern)).Current.Value);
                var last = window.FindFirst(System.Windows.Automation.TreeScope.Descendants,
                    new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.NameProperty, deep.Name));
                Assert.AreEqual(16, ((string[])Call("ItemPath", last)).Length);
                Assert.AreEqual(0, ((string[])Call("ItemPath", new object[] { null })).Length);
                var native = Native<VbeDebugWindows.IWatchProbe>("NativeWatchProbe");
                Assert.AreEqual(0, native.WatchMatches(host.Handle, "absent", "Root.Module"));
                dynamic read = Call("ReadList", host.Handle); Assert.IsNull((string)read.Error);
                dynamic parsed = VbeDebugWindows.ParseDebugRow(null, null); Assert.IsFalse((bool)parsed.Parsed);
            }
        }

        [TestMethod]
        public void NativePublicWatchDialogWrappersCompleteOnlyTheirOwnDialogs()
        {
            foreach (var kind in new[] { "Add Watch", "Edit Watch", "Quick Watch" })
                using (var scene = new SystemScene())
                {
                    var dialog = scene.Add(kind);
                    var ok = scene.Add("OK", "Button", dialog, 1); var cancel = scene.Add("Cancel", "Button", dialog, 2);
                    scene.OnMessage = (window, message) => { if (window == ok || window == cancel) dialog.Visible = false; };
                    var request = new Request { Project = "Book", Module = "Module1", Procedure = "Run", Expression = "x", NewExpression = "y", Context = "Module1.Run" };
                    if (kind == "Quick Watch")
                    {
                        scene.Add("x", "Edit", dialog, 4751); scene.Add("42", "Static", dialog, 4752); scene.Add("Book.Module1.Run", "Static", dialog, 4753);
                        dynamic result = VbeDebugWindows.CompleteQuickWatch(request); Assert.AreEqual("42", (string)result.Value);
                    }
                    else
                    {
                        scene.Add("x", "Edit", dialog, 4853); scene.Add("Book", "ComboBox", dialog, 4858);
                        scene.Add("Module1", "ComboBox", dialog, 4857); scene.Add("Run", "ComboBox", dialog, 4856);
                        scene.Add("Watch expression", "Button", dialog, 4850); scene.IntegerMessageResult = new IntPtr(1);
                        if (kind == "Add Watch") { dynamic result = VbeDebugWindows.CompleteAddWatch(request); Assert.IsTrue((bool)result.Added); }
                        else { dynamic result = VbeDebugWindows.CompleteEditWatch(request); Assert.IsTrue((bool)result.Edited); }
                    }
                    Assert.IsFalse(dialog.Visible);
                }
        }

        [TestMethod]
        public void NativePublicDiagnosticResponseAndCompileWaitUseTheObservedMessage()
        {
            const string diagnostic = "Run-time error '13': Type mismatch";
            var root = new AutomationNode { Kind = System.Windows.Automation.ControlType.Window, Name = "Diagnostic" };
            root.Add(new AutomationNode { Kind = System.Windows.Automation.ControlType.Text, Name = diagnostic });
            using (var host = new AutomationHost(root))
                using (var scene = new SystemScene())
                {
                    var dialog = scene.Add("Microsoft Visual Basic for Applications"); dialog.Handle = host.Handle;
                    scene.Add(diagnostic, "Static", dialog);
                    var ok = scene.Add("OK", "Button", dialog, 1);
                    scene.OnMessage = (window, message) => { if (window == ok) dialog.Visible = false; };
                    dynamic response = VbeDebugWindows.RespondDebugDialog(new Request { Diagnostic = diagnostic, Button = "OK" });
                    Assert.AreEqual("DialogClosed", (string)response.Verification);
                    dialog.Visible = true;
                    using (var completed = new System.Threading.ManualResetEventSlim(true))
                        Assert.AreEqual(diagnostic, VbeDebugWindows.AwaitCompileDialog(completed));
                }
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Threading;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeDebugWindowsNativeTests
    {
        [TestMethod]
        public void NativeCaptureRejectsNullProbeAndCompileWaitHandlesLateAndStuckDiagnostics()
        {
            Assert.ThrowsException<ArgumentNullException>(() => VbeDebugWindows.Capture(false, null));
            using (var completed = new ManualResetEventSlim(true))
            {
                var late = new FakeNative { AccessibleMessage = "Compile error: late" };
                late.Controls.Add(Control(4, "Button", "OK", true));
                late.OnPause = count => { if (count == 1) late.DialogHandle = new IntPtr(2); };
                Assert.AreEqual("Compile error: late", VbeDebugWindows.AwaitCompileDialog(completed, late));
            }
            using (var completed = new ManualResetEventSlim(false))
            {
                var stuck = DialogWith("Compile error: stuck", "OK");
                Assert.ThrowsException<TimeoutException>(() => VbeDebugWindows.AwaitCompileDialog(completed, stuck));
            }
        }
    }

    public sealed partial class VbeDebugWindowsImmediateProbeTests
    {
        [TestMethod]
        public void ImmediateOutputReplacementDoesNotInventAnAppendedOutputDelta()
        {
            var native = Ready("Debug.Print 1"); native.Readbacks.Add("replacement buffer");
            dynamic result = VbeDebugWindows.ExecuteImmediate("Debug.Print 1", native);
            Assert.IsNull((string)result.OutputDelta); Assert.AreEqual("ImmediateTextChangedAfterEnter", (string)result.Verification);
        }
    }

    public sealed partial class VbeDebugWindowsWatchProbeTests
    {
        [TestMethod]
        public void AddWatchCoversEveryTypeOptionalProcedureAndContextField()
        {
            foreach (var type in new[] { null, "expression", "break_when_true", "break_when_changed" })
            {
                var native = new WatchFake { Missing = 4856 }; var request = Add(); request.Procedure = null; request.WatchType = type;
                dynamic result = VbeDebugWindows.CompleteAddWatch(request, native);
                Assert.IsTrue((bool)result.Added); Assert.IsNull((string)result.Context.Procedure);
                Assert.AreEqual(type ?? "expression", (string)result.WatchType);
            }
            var mismatch = new WatchFake(); mismatch.Texts[4856] = "Other";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteAddWatch(Add(), mismatch));
        }

        [TestMethod]
        public void EditWatchCoversEveryNativeFieldContextAndTypeWithoutLosingItsRollback()
        {
            foreach (int missing in new[] { 4853, 4858, 4857, 1 })
            {
                var native = new WatchFake { Missing = missing };
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), native));
                Assert.AreEqual(1, native.Closes);
            }
            foreach (int field in new[] { 4858, 4857 })
            {
                var native = new WatchFake(); native.Texts[field] = "Changed";
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.CompleteEditWatch(Edit(), native));
            }
            foreach (var type in new[] { null, "expression", "break_when_true", "break_when_changed" })
            {
                var native = new WatchFake { Missing = 4856, Root = new IntPtr(1), PaneHandle = new IntPtr(2), NewMatches = 1, OldMatches = 0 };
                var request = Edit(); request.Context = "Module1"; request.WatchType = type;
                dynamic edited = VbeDebugWindows.CompleteEditWatch(request, native); Assert.IsFalse((bool)edited.VerificationPending);
            }
            var same = new WatchFake { Root = new IntPtr(1), PaneHandle = new IntPtr(2), OldMatches = 1 };
            var unchanged = Edit(); unchanged.NewExpression = unchanged.Expression;
            dynamic retained = VbeDebugWindows.CompleteEditWatch(unchanged, same); Assert.IsFalse((bool)retained.VerificationPending);
        }
    }
}

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
                "Compile error: Syntax error",
                "L'identificateur sous le curseur n'est pas reconnu",
                "Impossible d'aller à 'Range' qui est caché"
            }

            )
                Assert.IsTrue(VbeDebugWindows.IsRecognizedDiagnostic(message), message);
            foreach (var message in new[]
            {
                null,
                "",
                "Other application error",
                "Note: Compile error",
                "Windows Security",
                "Impossible d'aller à 'Range' qui est caché : autre dialogue",
                "Impossible d'aller à 'Range' qui est caché\n",
                "L'identificateur sous le curseur n'est pas reconnu : autre dialogue"
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

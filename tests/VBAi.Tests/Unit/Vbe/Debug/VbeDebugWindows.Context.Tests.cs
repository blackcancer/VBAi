using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeDebugWindowsContextTests
    {
        [TestMethod]
        public void ReadLocalsContextRequiresOnePaneAndAnOwnedReadOnlyEditWithItsButton()
        {
            foreach (string fault in new[] { "none", "duplicate", "missing edit", "foreign edit", "editable", "hidden edit", "missing button", "wrong button", "empty" })
            {
                using (var scene = new NativeDebugScene())
                {
                    var originalGetItem = VbeDebugWindows.GetDlgItem;
                    var originalStyle = VbeDebugWindows.ContextWindowStyle;
                    try
                    {
                        var root = scene.Add("wndclass_desked_gsk", "VBE");
                        var pane = fault == "none" ? null : scene.Add("VbaWindow", "Variables locales", root);
                        if (fault == "duplicate") scene.Add("VbaWindow", "Locals", root);
                        var edit = scene.Add("Edit", fault == "empty" ? "" : "VBAProject.Module1.Inspect", pane);
                        var button = scene.Add(fault == "wrong button" ? "Static" : "Button", "...", pane);
                        if (fault == "foreign edit") edit.Process = 999999;
                        if (fault == "hidden edit") edit.Visible = false;
                        VbeDebugWindows.GetDlgItem = (handle, id) => handle != pane?.Handle ? IntPtr.Zero :
                            id == 4604 && fault != "missing edit" ? edit.Handle :
                            id == 4601 && fault != "missing button" ? button.Handle : IntPtr.Zero;
                        VbeDebugWindows.ContextWindowStyle = handle => fault == "editable" ? 0 : 0x800;
                        Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadLocalsContext(), fault);
                    }
                    finally
                    {
                        VbeDebugWindows.GetDlgItem = originalGetItem;
                        VbeDebugWindows.ContextWindowStyle = originalStyle;
                    }
                }
            }
        }

        [TestMethod]
        public void ReadLocalsContextReturnsExactNativeContextWithoutChangingControls()
        {
            using (var scene = new NativeDebugScene())
            {
                var originalGetItem = VbeDebugWindows.GetDlgItem;
                var originalStyle = VbeDebugWindows.ContextWindowStyle;
                try
                {
                    var root = scene.Add("wndclass_desked_gsk", "VBE");
                    var pane = scene.Add("VbaWindow", "Locals", root);
                    var edit = scene.Add("Edit", "VBAProject.Module1.Inspect", pane);
                    var button = scene.Add("Button", "...", pane);
                    VbeDebugWindows.GetDlgItem = (handle, id) => handle == pane.Handle ?
                        id == 4604 ? edit.Handle : id == 4601 ? button.Handle : IntPtr.Zero : IntPtr.Zero;
                    VbeDebugWindows.ContextWindowStyle = handle => 0x800;
                    Assert.AreEqual("VBAProject.Module1.Inspect", VbeDebugWindows.ReadLocalsContext());
                    Assert.AreEqual(0, scene.Messages.Count);
                }
                finally
                {
                    VbeDebugWindows.GetDlgItem = originalGetItem;
                    VbeDebugWindows.ContextWindowStyle = originalStyle;
                }
            }
        }

        [TestMethod]
        public void ExistingQuickWatchOrVbaDiagnosticFailsPreflightWithoutClosingEither()
        {
            foreach (string title in new[] { "Espion express", "Microsoft Visual Basic pour Applications" })
                using (var scene = new NativeDebugScene())
                {
                    scene.Add("#32770", title);
                    Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.EnsureNoQuickWatchDialog(), title);
                    Assert.AreEqual(0, scene.Messages.Count);
                }
            using (var scene = new NativeDebugScene())
                VbeDebugWindows.EnsureNoQuickWatchDialog();
        }

        [TestMethod]
        public void OnlyExactKnownNoSelectionDiagnosticIsAcknowledged()
        {
            foreach (string scenario in new[] { "known", "unknown", "wrong button" })
                using (var scene = new NativeDebugScene())
                {
                    var originalGetItem = VbeDebugWindows.GetDlgItem;
                    var originalPost = VbeDebugWindows.PostMessage;
                    var originalPause = VbeDebugWindows.PauseNative;
                    try
                    {
                        var dialog = scene.Add("#32770", "Microsoft Visual Basic pour Applications");
                        string message = scenario == "unknown" ? "Une autre erreur VBA" : "Pas d'expression espionne sélectionnée";
                        scene.Add("Static", message, dialog);
                        var ok = scene.Add(scenario == "wrong button" ? "Static" : "Button", "OK", dialog);
                        VbeDebugWindows.GetDlgItem = (handle, id) => handle == dialog.Handle && id == 2 ? ok.Handle : IntPtr.Zero;
                        var clicked = new System.Collections.Generic.List<IntPtr>();
                        VbeDebugWindows.PostMessage = (handle, command, w, l) => { clicked.Add(handle); dialog.Visible = false; return true; };
                        VbeDebugWindows.PauseNative = milliseconds => { };
                        var failure = Assert.ThrowsException<InvalidOperationException>(() =>
                            VbeDebugWindows.ReadScalarQuickWatch(new Request { Expression = "first" }));
                        StringAssert.Contains(failure.Message, message);
                        Assert.AreEqual(scenario == "known" ? 1 : 0, clicked.Count);
                        if (clicked.Count == 1) Assert.AreEqual(ok.Handle, clicked[0]);
                    }
                    finally
                    {
                        VbeDebugWindows.GetDlgItem = originalGetItem;
                        VbeDebugWindows.PostMessage = originalPost;
                        VbeDebugWindows.PauseNative = originalPause;
                    }
                }
        }

        [TestMethod]
        public void ScalarReadWaitsForOwnedQuickWatchToCloseAndRejectsAnUnclosedDialog()
        {
            foreach (bool closes in new[] { true, false })
                using (var scene = new NativeDebugScene())
                {
                    var originalGetItem = VbeDebugWindows.GetDlgItem;
                    var originalPost = VbeDebugWindows.PostMessage;
                    var originalPause = VbeDebugWindows.PauseNative;
                    try
                    {
                        var dialog = scene.Add("#32770", "Espion express");
                        var expression = scene.Add("Edit", "first", dialog);
                        var value = scene.Add("Static", "42", dialog);
                        var context = scene.Add("Edit", "VBAProject.Module1.Inspect", dialog);
                        var cancel = scene.Add("Button", "Annuler", dialog);
                        VbeDebugWindows.GetDlgItem = (handle, id) => handle != dialog.Handle ? IntPtr.Zero :
                            id == 4751 ? expression.Handle : id == 4752 ? value.Handle :
                            id == 4753 ? context.Handle : id == 2 ? cancel.Handle : IntPtr.Zero;
                        int clicks = 0, pauses = 0;
                        VbeDebugWindows.PostMessage = (handle, command, w, l) =>
                        {
                            Assert.AreEqual(cancel.Handle, handle);
                            clicks++;
                            if (closes) dialog.Visible = false;
                            return true;
                        };
                        VbeDebugWindows.PauseNative = milliseconds => { pauses++; };
                        var request = new Request { Project = "VBAProject", Module = "Module1", Procedure = "Inspect", Expression = "first" };
                        if (closes)
                        {
                            dynamic result = VbeDebugWindows.ReadScalarQuickWatch(request);
                            Assert.AreEqual("42", (string)result.Value);
                            Assert.AreEqual("VBAProject.Module1.Inspect", (string)result.Context);
                        }
                        else Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadScalarQuickWatch(request));
                        Assert.AreEqual(1, clicks);
                        Assert.AreEqual(closes ? 1 : 41, pauses);
                    }
                    finally
                    {
                        VbeDebugWindows.GetDlgItem = originalGetItem;
                        VbeDebugWindows.PostMessage = originalPost;
                        VbeDebugWindows.PauseNative = originalPause;
                    }
                }
        }

        [TestMethod]
        public void ReadFailureStillWaitsForQuickWatchClosureBeforeReportingIt()
        {
            foreach (bool closes in new[] { true, false })
                using (var scene = new NativeDebugScene())
                {
                    var originalGetItem = VbeDebugWindows.GetDlgItem;
                    var originalPost = VbeDebugWindows.PostMessage;
                    var originalPause = VbeDebugWindows.PauseNative;
                    try
                    {
                        var dialog = scene.Add("#32770", "Quick Watch");
                        var expression = scene.Add("Edit", "different", dialog);
                        var value = scene.Add("Static", "42", dialog);
                        var context = scene.Add("Edit", "VBAProject.Module1.Inspect", dialog);
                        var cancel = scene.Add("Button", "Cancel", dialog);
                        VbeDebugWindows.GetDlgItem = (handle, id) => handle != dialog.Handle ? IntPtr.Zero :
                            id == 4751 ? expression.Handle : id == 4752 ? value.Handle :
                            id == 4753 ? context.Handle : id == 2 ? cancel.Handle : IntPtr.Zero;
                        int clicks = 0;
                        VbeDebugWindows.PostMessage = (handle, command, w, l) =>
                        {
                            clicks++;
                            if (closes) dialog.Visible = false;
                            return true;
                        };
                        VbeDebugWindows.PauseNative = milliseconds => { };
                        var request = new Request { Project = "VBAProject", Module = "Module1", Procedure = "Inspect", Expression = "first" };
                        var failure = Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadScalarQuickWatch(request));
                        StringAssert.Contains(failure.Message, closes ? "differs" : "did not close");
                        Assert.AreEqual(1, clicks);
                    }
                    finally
                    {
                        VbeDebugWindows.GetDlgItem = originalGetItem;
                        VbeDebugWindows.PostMessage = originalPost;
                        VbeDebugWindows.PauseNative = originalPause;
                    }
                }
        }

        [TestMethod]
        public void MissingQuickWatchTimesOutOnceWithoutStartingASecondRead()
        {
            using (var scene = new NativeDebugScene())
            {
                var originalPause = VbeDebugWindows.PauseNative;
                try
                {
                    int pauses = 0;
                    VbeDebugWindows.PauseNative = milliseconds => { pauses++; };
                    var failure = Assert.ThrowsException<InvalidOperationException>(() =>
                        VbeDebugWindows.ReadScalarQuickWatch(new Request { Expression = "first" }));
                    StringAssert.Contains(failure.Message, "did not open");
                    Assert.AreEqual(60, pauses);
                }
                finally { VbeDebugWindows.PauseNative = originalPause; }
            }
        }
    }
}

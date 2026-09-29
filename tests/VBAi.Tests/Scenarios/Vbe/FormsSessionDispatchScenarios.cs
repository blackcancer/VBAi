namespace VBAi.Tests.Unit
{
    using System.Threading;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    // Exercise protocol dispatch against the same live Designer contracts as the service tests.
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeFormsSessionDispatchTests
    {
        private static dynamic Call(VbeSession session, Request request, string command)
        {
            request.Command = command;
            var result = session.Execute(request);
            Assert.IsTrue(result.Ok, result.Error);
            return result.Data;
        }

        [TestMethod]
        public void SessionRoutesClipboardSelectionCutRestoreAndRecoveryToItsOwnFormsService()
        {
            using (var f = new FormsWorkflowFixture())
            {
                var session = new VbeSession(f.Vbe);
                dynamic state = Call(session, f.Request(), "form_clipboard_state");
                Assert.IsNotNull((string)state.SelectionVersion);
                var selection = f.ClipboardRequest(); selection.Items = new[] { "A" };
                Call(session, selection, "select_form_controls");
                Assert.IsTrue(f.Form.Designer.Controls.Item("A").InSelection);
                dynamic cut = Call(session, f.ClipboardRequest("cut"), "native_form_clipboard");
                Assert.IsTrue((bool)cut.Executed);
                string id = cut.DesignerClipboardRecoveryId;
                var restore = f.ClipboardRequest(); restore.DesignerClipboardRecoveryId = id;
                Assert.IsTrue(session.CanRecoverFormCut(restore));
                Assert.IsTrue((bool)Call(session, restore, "restore_form_clipboard").Restored);
                restore = f.ClipboardRequest(); restore.DesignerClipboardRecoveryId = id;
                Assert.IsTrue((bool)Call(session, restore, "recover_form_cut").RestoredNamesGeometryAndTabOrder);
                Assert.AreEqual(2, f.Form.Designer.Controls.Count);
                Assert.IsFalse(session.CanRecoverFormCut(restore));
            }
        }

        [TestMethod]
        public void SessionRoutesLayoutTabOrderHistoryAndQueuedRunWithObservableNativeEffects()
        {
            var previous = SynchronizationContext.Current;
            var context = new NativeNavigationContext();
            SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                using (var f = new FormsWorkflowFixture())
                {
                    var session = new VbeSession(f.Vbe);
                    var layout = f.Request("align_left"); layout.Items = new[] { "Controls/A", "Controls/B" };
                    Assert.IsNotNull((object)Call(session, layout, "preview_form_layout").After);
                    Assert.AreEqual(40, f.Form.Designer.Controls.Item("B").Left);
                    Assert.IsTrue((bool)Call(session, layout, "apply_form_layout").Verified);
                    Assert.AreEqual(0, f.Form.Designer.Controls.Item("B").Left);
                    var tabs = f.Request(); tabs.Items = new[] { "B", "A" };
                    Assert.IsTrue((bool)Call(session, tabs, "set_form_tab_order").Verified);
                    Assert.AreEqual(0, f.Form.Designer.Controls.Item("B").TabIndex);
                    Call(session, f.Request("undo"), "native_form_history");
                    StringAssert.Contains(f.Form.Designer.Caption, "undo");
                    dynamic queued = Call(session, f.Request(), "run_form");
                    string operation = queued.OperationId;
                    context.RunAll();
                    dynamic status = Call(session, new Request { Project = f.Project.Name, Query = operation }, "form_run_status");
                    Assert.IsTrue((bool)status.CommandCompleted);
                    Assert.AreEqual(1, ((FormsWorkflowFixture.RunControl)f.Vbe.CommandBars.Control).Calls);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }
    }

    public sealed partial class VbeFormsInitializerTests
    {
        [TestMethod]
        public void SessionRoutesListBindingThroughRealCodeGenerationAndReadback()
        {
            var f = Create(); f.Project.FileName = @"C:\Tests\Macro.xlsm";
            var request = Binding(f); request.Command = "set_form_list_binding";
            var response = new VbeSession(f.Host).Execute(request);
            Assert.IsTrue(response.Ok, response.Error);
            Assert.IsTrue((bool)((dynamic)response.Data).Verified);
            Assert.IsTrue(f.Form.CodeModule.InsertCount > 0);
            StringAssert.Contains(f.Form.CodeModule.Code, "RowSource");
        }
    }
}

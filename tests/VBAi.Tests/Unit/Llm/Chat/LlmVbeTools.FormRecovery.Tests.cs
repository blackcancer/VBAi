using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace VBAi.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class LlmVbeToolsFormRecoveryTests
    {
        public sealed class RecoveryState
        {
            public string SelectionVersion { get; set; } = "selection";
            public string ClipboardVersion { get; set; } = "clipboard";
        }
        public sealed class RecoveryResult
        {
            public bool RecoveryAttempted { get; set; }
            public bool RestoredNamesGeometryAndTabOrder { get; set; }
            public string NativeError { get; set; }
        }
        [TestMethod]
        public void RecoveryValidatesLiveOwnershipScopeAndSingleAttemptBeforeNativeAccess()
        {
            var tools = new LlmVbeTools(new VbeSession(new object()), null, new LlmSettings()) { BoundProject = "P" };
            Assert.IsFalse(tools.CanRecoverFormCut(new FormCutChange { Owner = tools, Project = "P", RecoveryId = "absent" }));
            int native = 0;
            tools.CanRecoverDesignerCut = r => { native++; return true; };
            Func<FormCutChange> valid = () => new FormCutChange { Owner = tools, Project = "p", Form = "F", ParentPath = "Frame", RecoveryId = "token" };
            var other = new LlmVbeTools(null, null, new LlmSettings());
            foreach (var invalid in new FormCutChange[] { null,
                new FormCutChange { Project = "P" }, new FormCutChange { Owner = other, Project = "P" },
                new FormCutChange { Owner = tools, Project = "P", Attempted = true },
                new FormCutChange { Owner = tools, Project = "P", Restored = true },
                new FormCutChange { Owner = tools, Project = "Other" } })
            {
                Assert.IsFalse(tools.CanRecoverFormCut(invalid));
                Assert.IsFalse(tools.RecoverFormCut(invalid).Ok);
            }
            Assert.AreEqual(0, native);
            Assert.IsTrue(tools.CanRecoverFormCut(valid()));
            tools.CanRecoverDesignerCut = r => false;
            Assert.IsFalse(tools.CanRecoverFormCut(valid()));
            tools.CanRecoverDesignerCut = r => { throw new InvalidOperationException("designer closed"); };
            Assert.IsFalse(tools.CanRecoverFormCut(valid()));
            tools.ValidateScope = () => { throw new InvalidOperationException("scope closed"); };
            StringAssert.Contains(tools.RecoverFormCut(valid()).Error, "scope closed");
            tools.ValidateScope = null;
            tools.Execute = r => Response.Failure("clipboard unavailable");
            StringAssert.Contains(tools.RecoverFormCut(valid()).Error, "clipboard unavailable");
            tools.Execute = r => Response.Success(new { });
            Assert.IsFalse(tools.RecoverFormCut(valid()).Ok);
            tools.Execute = r => { throw new InvalidOperationException("disconnected"); };
            StringAssert.Contains(tools.RecoverFormCut(valid()).Error, "disconnected");
        }

        [TestMethod]
        public void RecoveryTransfersBothRevisionsAndReportsNativeResultWithoutInventingSuccess()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings()) { BoundProject = "P" };
            foreach (bool success in new[] { false, true })
                foreach (bool restored in new[] { false, true })
                {
                    var change = new FormCutChange { Owner = tools, Project = "P", Form = "F", ParentPath = "Frame", RecoveryId = "token" };
                    int calls = 0, validations = 0;
                    tools.ValidateScope = () => validations++;
                    tools.Execute = r =>
                    {
                        calls++;
                        Assert.AreEqual("P", r.Project); Assert.AreEqual("F", r.Form);
                        Assert.AreEqual("Frame", r.ParentPath); Assert.AreEqual("token", r.DesignerClipboardRecoveryId);
                        if (calls == 1) { Assert.AreEqual("form_clipboard_state", r.Command); return Response.Success(new RecoveryState()); }
                        Assert.AreEqual("recover_form_cut", r.Command);
                        Assert.AreEqual("selection", r.ExpectedDesignerSelectionVersion);
                        Assert.AreEqual("clipboard", r.ExpectedClipboardVersion);
                        return success ? Response.Success(new RecoveryResult { RecoveryAttempted = true, RestoredNamesGeometryAndTabOrder = restored }) : Response.Failure("native failure");
                    };
                    Assert.AreEqual(success, tools.RecoverFormCut(change).Ok);
                    Assert.AreEqual(success, change.Attempted);
                    Assert.AreEqual(success && restored, change.Restored);
                    Assert.AreEqual(2, calls); Assert.AreEqual(1, validations);
                }
            tools.ValidateScope = null;
            tools.Execute = r => r.Command == "form_clipboard_state" ? Response.Success(new RecoveryState()) : Response.Success(new { });
            Assert.IsFalse(tools.RecoverFormCut(new FormCutChange { Owner = tools, Project = "P" }).Ok);
        }
    }
}

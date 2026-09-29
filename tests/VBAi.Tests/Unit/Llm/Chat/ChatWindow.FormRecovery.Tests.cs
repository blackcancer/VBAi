using System;
using System.Collections.Generic;
using System.Windows.Controls;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod]
        public void FormRecoveryCardsReflectOwnershipBusyStateAndVerifiedNativeOutcome()
        {
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                var tools = Get<LlmVbeTools>(window, "tools");
                tools.BoundProject = "P"; tools.ValidateScope = null;
                bool available = true;
                tools.CanRecoverDesignerCut = r => available;
                var cards = Get<Dictionary<FormCutChange, System.Windows.Forms.Button>>(window, "formCutButtons");
                var live = new FormCutChange { Owner = tools, Project = "P", Form = "F", ControlCount = 2 };
                Call(window, "RenderFormCut", live);
                Assert.IsTrue(cards[live].Enabled);
                Set(window, "busy", true); Call(window, "RefreshFormCutCards");
                Assert.IsFalse(cards[live].Enabled); WpfClick(cards[live]);
                Set(window, "busy", false); available = false;
                Call(window, "RefreshFormCutCards"); Assert.IsFalse(cards[live].Enabled); WpfClick(cards[live]);
                Set(window, "tools", null); WpfClick(cards[live]); Call(window, "RefreshFormCutCards");
                Set(window, "tools", tools); available = true;
                var historical = new FormCutChange { Form = "Historical", ParentPath = "Frame" };
                Call(window, "RenderFormCut", historical);
                Assert.IsFalse(cards[historical].Enabled);
                foreach (int outcome in new[] { 0, 1, 2, 3 })
                {
                    var change = new FormCutChange { Owner = tools, Project = "P", Form = "F", ParentPath = "Frame" };
                    Call(window, "RenderFormCut", change);
                    tools.Execute = r => r.Command == "form_clipboard_state" ? Response.Success(new LlmVbeToolsFormRecoveryTests.RecoveryState()) :
                        outcome == 0 ? Response.Failure("recovery rejected") : Response.Success(new LlmVbeToolsFormRecoveryTests.RecoveryResult { RecoveryAttempted = outcome != 3, RestoredNamesGeometryAndTabOrder = outcome == 1, NativeError = outcome == 3 ? (string)null : "native detail" });
                    WpfClick(cards[change]);
                    Assert.AreEqual(outcome == 1, change.Restored);
                    Assert.AreEqual(outcome == 1 || outcome == 2, change.Attempted);
                    Assert.IsFalse(cards[change].Enabled && change.Attempted);
                    string status = Get<System.Windows.Forms.Label>(window, "status").Text;
                    if (outcome == 0) StringAssert.Contains(status, "recovery rejected");
                    if (outcome == 2) StringAssert.Contains(status, "native detail");
                }
            }
        }
    }
}

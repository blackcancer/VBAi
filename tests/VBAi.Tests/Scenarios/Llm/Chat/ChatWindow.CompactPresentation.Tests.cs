using System;
using System.Collections.Generic;
using System.Windows.Forms;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void ApprovalPickerPersistsPolicyAndRefusesChangesDuringRunOrFailedSave()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var picker = Get<ComboBox>(window, "approvalPicker");
                picker.SelectedIndex = 2;
                Assert.AreEqual("ReadOnly", runtime.Settings.VbeEditApproval);
                Call(window, "SetBusy", true);
                Assert.IsFalse(picker.Enabled);
                picker.SelectedIndex = 0;
                Assert.AreEqual(2, picker.SelectedIndex);
                Assert.AreEqual("ReadOnly", runtime.Settings.VbeEditApproval);
                Call(window, "SetBusy", false);
                ChatWindow.WriteSettings = s => { throw new InvalidOperationException("save failed"); };
                picker.SelectedIndex = 1;
                Assert.AreEqual("ReadOnly", runtime.Settings.VbeEditApproval);
                Assert.AreEqual(2, picker.SelectedIndex);
                ChatWindow.WriteSettings = s => runtime.Saves++;
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void ModelSummaryRevealsExistingDesignerSelectorsWithoutChangingTheModel()
        {
            using (var window = Surfaces())
            {
                var picker = Get<ComboBox>(window, "modelPicker");
                var model = new LlmModelOption("opaque-id", "Readable model");
                picker.Items.Add(model); picker.SelectedItem = model;
                Call(window, "RefreshModelSummary");
                StringAssert.Contains(Get<ChatActionButton>(window, "modelSummary").Text, "Readable model");
                Call(window, "ModelSummary_Click", null, EventArgs.Empty);
                Assert.IsTrue(Get<TableLayoutPanel>(window, "rootLayout").RowStyles[6].Height > 0);
                Assert.AreSame(model, picker.SelectedItem);
                Call(window, "ModelSummary_Click", null, EventArgs.Empty);
                Assert.AreEqual(0F, Get<TableLayoutPanel>(window, "rootLayout").RowStyles[6].Height);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void RunningReasoningOpensInlineAndRespectsUserCollapseAcrossUpdates()
        {
            using (var window = Surfaces())
            {
                var entry = new ChatEntry { Speaker = "Réflexion", Activity = new CodexAgentActivity { Kind = "reasoning", Status = "inProgress", Detail = "Provider summary" } };
                var entries = new List<ChatEntry> { entry };
                using (var group = (ChatDesignerHost)Call(window, "RenderActivityGroup", entry, entries))
                {
                    var view = (ChatActivityGroupView)group.View;
                    Assert.IsTrue(view.section.Expanded);
                    view.section.Expanded = false;
                }
                using (var group = (ChatDesignerHost)Call(window, "RenderActivityGroup", entry, entries))
                    Assert.IsFalse(((ChatActivityGroupView)group.View).section.Expanded);
                using (var step = (ChatDesignerHost)Call(window, "RenderActivityStep", entry))
                {
                    var view = (ChatActivityStepView)step.View;
                    Assert.IsTrue(view.section.Expanded);
                    Assert.AreEqual("Provider summary", view.detail.content.Text);
                    view.section.Expanded = false;
                }
                using (var step = (ChatDesignerHost)Call(window, "RenderActivityStep", entry))
                    Assert.IsFalse(((ChatActivityStepView)step.View).section.Expanded);
                entry.Activity.Status = "completed";
                using (var group = (ChatDesignerHost)Call(window, "RenderActivityGroup", entry, entries))
                    Assert.IsFalse(((ChatActivityGroupView)group.View).section.Expanded);
            }
        }
    }
}

using Blackcancer.ChatUI.Views;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using VBAi.Tests.Integration;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaOfficeUiControlTests
    {
        [STATestMethod]
        [DataRow("options", "settingsButton"), DataRow("send", "sendButton"), DataRow("stop", "stopButton")]
        [DataRow("promptHost", "inputTextBox"), DataRow("transcriptPanel", "TranscriptHost")]
        [DataRow("scopePicker", "scopePicker"), DataRow("modelSummary", "modelSummary")]
        [DataRow("providerPicker", "providerPicker"), DataRow("modelPicker", "modelPicker")]
        [DataRow("effortPicker", "effortPicker"), DataRow("modePicker", "modePicker")]
        [DataRow("approvalPicker", "approvalPicker")]
        public void ObserverRolesResolveToExactlyOneActualDesignerControl(string role, string expected)
        {
            using (var view = new ChatWorkspaceView(CultureInfo.GetCultureInfo("fr-FR")))
            {
                string identity = OllamaOfficeUi.AssistantControlId(role);
                Assert.AreEqual(expected, identity);
                Assert.AreEqual(1, view.Controls.Find(identity, true).Length);
                if (role != expected && role != "transcriptPanel") Assert.AreEqual(0, view.Controls.Find(role, true).Length,
                    "Legacy roles cannot be used as actual AutomationIds.");
            }
        }

        [TestMethod]
        [DataRow(false, true, false, true, true, false, false)] // hidden send
        [DataRow(true, false, false, true, true, false, false)] // disabled send
        [DataRow(false, true, true, false, false, false, false)] // streaming with empty draft
        [DataRow(true, true, true, false, false, false, false)] // streaming while drafting
        [DataRow(true, true, false, false, true, true, false)] // provider still settling
        [DataRow(true, true, false, true, false, true, false)] // mode still settling
        [DataRow(true, true, false, true, true, true, true)] // complete idle recovery
        public void ActionReadinessDistinguishesHiddenSendLiveStopAndSettlingSelectors(bool sendVisible,
            bool sendEnabled, bool stopVisible, bool provider, bool mode, bool canSend, bool canFinish)
        {
            var state = new OllamaOfficeUi.AssistantActionState { SendVisible = sendVisible, SendEnabled = sendEnabled,
                StopVisible = stopVisible, ProviderEnabled = provider, ModeEnabled = mode };
            Assert.AreEqual(canSend, OllamaOfficeUi.CanSend(state));
            Assert.AreEqual(canFinish, OllamaOfficeUi.CanFinish(state));
        }

        [TestMethod]
        public void MissingActionSnapshotCannotAdmitSendOrCompletion()
        {
            Assert.IsFalse(OllamaOfficeUi.CanSend(null));
            Assert.IsFalse(OllamaOfficeUi.CanFinish(null));
        }

        [TestMethod]
        [DataRow("stop", 0, true), DataRow("stop", 1, false), DataRow("stop", 2, false)]
        [DataRow("send", 0, false), DataRow("unknown", 0, false), DataRow(null, 0, false)]
        public void DeferredInitialStopDoesNotPermitAbsentOrAmbiguousRequiredControls(string role, int peers, bool allowed)
            => Assert.AreEqual(allowed, OllamaOfficeUi.InitialStopCanBeAbsent(role, peers));

        [TestMethod]
        [DataRow("send"), DataRow("stop")]
        public void ClickPreflightRequiresThePhysicalIdentityEvenWhenTheRoleIsKnown(string role)
        {
            string physical = OllamaOfficeUi.AssistantControlId(role);
            Assert.IsTrue(PrivateDesktopUiAction.MatchesButton(physical, physical, "Action", "Action",
                "ControlType.Button", 42, 42, 77, 42, 77, 100, 100, true, false, true, true, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesButton(role, physical, "Action", "Action",
                "ControlType.Button", 42, 42, 77, 42, 77, 100, 100, true, false, true, true, true));
        }

        [TestMethod]
        [DataRow("same"), DataRow("replaced"), DataRow("foreign-process"), DataRow("foreign-thread"), DataRow("not-child")]
        public void HiddenObservationRetainsOriginalOwnershipAndCannotAdmitHiddenActions(string scenario)
        {
            long handle = scenario == "replaced" ? 43 : 42;
            int pid = scenario == "foreign-process" ? 8 : 7;
            uint thread = scenario == "foreign-thread" ? 10U : 9U;
            bool child = scenario != "not-child";
            Assert.AreEqual(scenario == "same", OllamaOfficeUi.FixedControlIdentityMatches(42, handle,
                "stopButton", "stopButton", 7, 9, pid, pid, thread, child, false, true));
            Assert.IsFalse(OllamaOfficeUi.FixedControlIdentityMatches(42, handle,
                "stopButton", "stopButton", 7, 9, pid, pid, thread, child, false));
        }

        [STATestMethod]
        [DataRow("en-US"), DataRow("fr-FR"), DataRow("de-DE"), DataRow("es-ES")]
        [DataRow("it-IT"), DataRow("pt-BR"), DataRow("ru-RU"), DataRow("ar-SA")]
        [DataRow("hi-IN"), DataRow("ja-JP"), DataRow("ko-KR"), DataRow("zh-CN"), DataRow("zh-TW")]
        public void ApprovalChoiceMatchesTheConstructedLocalizedCatalogue(string culture)
        {
            using (new LocalizationScope(culture))
            using (var form = new ChatWindow())
            {
                var picker = (ComboBox)form.Controls.Find("approvalPicker", true).Single();
                Assert.IsTrue(picker.Items.Contains(OllamaOfficeChatScenario.ReadOnlyApprovalChoice));
                Assert.AreEqual(UiText.Get("Read-only"), OllamaOfficeChatScenario.ReadOnlyApprovalChoice);
                var field = typeof(ChatWindow).GetField("transcriptPanel", BindingFlags.Instance | BindingFlags.NonPublic);
                Assert.AreEqual("TranscriptHost", ((Control)field.GetValue(form)).Name);
            }
        }
    }
}

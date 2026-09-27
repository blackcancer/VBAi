namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Net.Http;
    using System.Reflection;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class ChatWindowStateTests
    {
        [TestMethod]
        [STATestMethod]
        public void EditorActionSeedsCommandWhileBusyStateBlocksIt()
        {
            using (var window = Surfaces())
            {
                Call(window, "SetBusy", true);
                window.PrepareEditorAction("/corriger");
                var prompt = Get<object>(window, "prompt");
                Assert.AreEqual("", prompt.GetType().GetProperty("Text").GetValue(prompt, null));
                Call(window, "SetBusy", false);
                window.PrepareEditorAction("/corriger");
                Assert.AreEqual("/corriger ", prompt.GetType().GetProperty("Text").GetValue(prompt, null));
                Assert.AreEqual(ChatMode.Agent, Get<ComboBox>(window, "modePicker").SelectedItem);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ModeSelectionUpdatesSessionOnlyWhenIdle()
        {
            using (var window = Surfaces())
            {
                var session = new ChatSessionState
                {
                    Mode = ChatMode.Agent
                };
                Set(window, "currentSession", session);
                Get<ComboBox>(window, "modePicker").SelectedItem = ChatMode.Plan;
                Call(window, "ModePicker_SelectedIndexChanged", null, EventArgs.Empty);
                Assert.AreEqual(ChatMode.Plan, session.Mode);
                Call(window, "SetBusy", true);
                Get<ComboBox>(window, "modePicker").SelectedItem = ChatMode.Discussion;
                Call(window, "ModePicker_SelectedIndexChanged", null, EventArgs.Empty);
                Assert.AreEqual(ChatMode.Plan, session.Mode);
                Set(window, "currentSession", null);
            }
        }
    }
}

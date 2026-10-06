using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaOfficeStreamOracleTests
    {
        [TestMethod]
        public void PromptEchoCannotProveCompletedRecovery()
        {
            Assert.IsFalse(OllamaOfficeStreamOracle.IsReadyResponse("Reply with exactly UI_READY_42 and nothing else."));
            Assert.IsFalse(OllamaOfficeStreamOracle.IsReadyResponse("UI_READY_42 plus something else"));
            Assert.IsFalse(OllamaOfficeStreamOracle.IsReadyResponse(null));
            Assert.IsTrue(OllamaOfficeStreamOracle.IsReadyResponse("\r\nUI_READY_42\r\n"));
        }

        [TestMethod]
        public void SuccessfulDiscoveryConditionIsNotObservedTwice()
        {
            int calls = 0, records = 0;
            new OllamaOfficeUi(0, observation => records++).Wait(() => { calls++; return true; }, 1, "synthetic discovery");
            Assert.AreEqual(1, calls, "A successful condition may publish an identity receipt; evaluating twice duplicates it.");
            Assert.AreEqual(1, records);
        }

        [TestMethod]
        public void EchoedComposerAndSettingsCannotProveAssistantStreaming()
        {
            Assert.IsFalse(OllamaOfficeStreamOracle.IsNumberedResponse("Write a numbered list beginning immediately with item 1."));
            Assert.IsFalse(OllamaOfficeStreamOracle.IsNumberedResponse("qwen2.5:7b-instruct | Temperature 0"));
            Assert.IsFalse(OllamaOfficeStreamOracle.IsNumberedResponse("No text response."));
            Assert.IsTrue(OllamaOfficeStreamOracle.IsNumberedResponse("1. Chair\n2. Table\n3. Pencil\n"));
            Assert.IsTrue(OllamaOfficeStreamOracle.IsNumberedResponse("1) Chair\n2) Table\n3) Pencil\n"));
        }

        [STATestMethod]
        public void NativeTranscriptReadExcludesEditableHiddenAndClippedFields()
        {
            using (var form = new ObserverForm())
            using (var panel = new Panel { Dock = DockStyle.Fill })
            using (var response = new RichTextBox { ReadOnly = true, Text = "1. Object 1\n2. Object 2\n", Bounds = new Rectangle(0, 0, 350, 60) })
            using (var composer = new RichTextBox { Text = "COMPOSER_MUST_NOT_BE_READ", Bounds = new Rectangle(0, 65, 350, 40) })
            using (var hidden = new RichTextBox { ReadOnly = true, Text = "HIDDEN_MUST_NOT_BE_READ", Visible = false })
            using (var clipped = new RichTextBox { ReadOnly = true, Text = "CLIPPED_MUST_NOT_BE_READ", Bounds = new Rectangle(-1000, 0, 100, 30) })
            {
                panel.Controls.AddRange(new Control[] { response, composer, hidden, clipped });
                form.Controls.Add(panel); form.Show(); Application.DoEvents();
                var snapshot = OllamaOfficeUi.ReadNativeTranscript(panel.Handle, Process.GetCurrentProcess().Id, GetCurrentThreadId());
                Assert.AreEqual(1, snapshot.Length);
                Assert.AreEqual(response.Text.Replace("\r", ""), snapshot[0].Replace("\r", ""));
                Assert.IsTrue(OllamaOfficeStreamOracle.IsNumberedResponse(snapshot[0]));
            }
        }

        [STATestMethod]
        public void NativeTranscriptReadRejectsForeignPanelIdentityBeforeReading()
        {
            using (var form = new ObserverForm())
            {
                form.Show(); Application.DoEvents();
                int pid = Process.GetCurrentProcess().Id; uint thread = GetCurrentThreadId();
                Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeUi.ReadNativeTranscript(form.Handle, pid + 1, thread));
                Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeUi.ReadNativeTranscript(form.Handle, pid, thread + 1));
            }
        }

        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        private sealed class ObserverForm : Form
        {
            protected override bool ShowWithoutActivation => true;
            internal ObserverForm()
            { ShowInTaskbar = false; StartPosition = FormStartPosition.Manual; Location = new Point(-32000, -32000); Size = new Size(400, 220); }
        }
    }
}

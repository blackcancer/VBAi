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
        public void TranscriptStreamsUpdateInPlaceAndAvoidDuplicateFinalMessage()
        {
            using (var window = Surfaces())
            {
                Call(window, "ReceiveChatUpdate", "summary", "empty", " ", false);
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.AreEqual(0, entries.Count);
                Call(window, "ReceiveChatUpdate", "final", "answer", "Hel", false);
                Call(window, "ReceiveChatUpdate", "final", "answer", "lo", true);
                Assert.AreEqual(1, entries.Count);
                Assert.AreEqual("lo", entries[0].Text);
                Call(window, "CompleteAssistantResponse", "lo");
                Assert.AreEqual(1, entries.Count);
                Call(window, "CompleteAssistantResponse", "different");
                Assert.AreEqual(2, entries.Count);
                Call(window, "ClearTranscript");
                Assert.AreEqual(0, entries.Count);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void StreamingSummaryAndToolEntriesCompleteInPlace()
        {
            using (var window = Surfaces())
            {
                Call(window, "ReceiveChatUpdate", "summary", "summary-1", "Réflexion ", false);
                Call(window, "ReceiveChatUpdate", "summary", "summary-1", "terminée", true);
                Call(window, "ReceiveChatUpdate", "tool", "tool-1", "Lecture", false);
                Call(window, "ReceiveChatUpdate", "tool", "tool-1", null, true);
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.AreEqual(2, entries.Count);
                Assert.AreEqual("Réflexion", entries[0].Speaker);
                Assert.AreEqual("terminée", entries[0].Text);
                Assert.AreEqual("Outil", entries[1].Speaker);
                Assert.AreEqual("Lecture", entries[1].Text);
                Assert.AreEqual(2, Get<HashSet<string>>(window, "completedStreams").Count);
            }
        }
    }
}

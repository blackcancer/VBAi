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
        public void CurrentReferencesKeepLatestExactTokenAndIgnoreEmbeddedMatches()
        {
            using (var window = Surfaces())
            {
                var selected = Get<List<VbeChatReference>>(window, "selectedReferences");
                selected.Add(new VbeChatReference { Project = "P", Module = "M", Kind = "Module", Sha256 = "old" });
                selected.Add(new VbeChatReference { Project = "P", Module = "M", Kind = "Module", Sha256 = "new" });
                var exact = (VbeChatReference[])Call(window, "CurrentReferences", "Use #P.M for this change");
                Assert.AreEqual(1, exact.Length);
                Assert.AreEqual("new", exact[0].Sha256);
                var embedded = (VbeChatReference[])Call(window, "CurrentReferences", "prefix#P.Msuffix");
                Assert.AreEqual(0, embedded.Length);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ContextChipsRepresentMemoryAndDraftAttachmentsAndCanRemoveDraft()
        {
            using (var window = Surfaces())
            {
                Set(window, "projectMemory", "local note");
                Get<CheckBox>(window, "attachMemory").Checked = true;
                var attachments = Get<List<ChatAttachment>>(window, "draftAttachments");
                attachments.Add(new ChatAttachment { Label = "Selected code", Text = "Sub A()" });
                Call(window, "RefreshContextChips");
                var chips = Get<FlowLayoutPanel>(window, "contextChips");
                Assert.AreEqual(2, chips.Controls.Count);
                typeof(Control).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(chips.Controls[1], new object[] { EventArgs.Empty });
                Assert.AreEqual(0, attachments.Count);
                Assert.AreEqual(1, chips.Controls.Count);
            }
        }
    }
}

namespace VBAi.Tests.Unit
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
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie la conservation des entrées du transcript et de l’aperçu de contexte.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Ajoute une entrée Markdown et vérifie l’affichage structuré du transcript et des pièces jointes.</summary>
        [TestMethod]
        [STATestMethod]
        public void MarkdownTranscriptAndContextPreviewRetainStructuredEntries()
        {
            using (var window = Surfaces())
            {
                Call(window, "AddTranscriptMessage", "Assistant", "# Heading\n- item\n```vba\nDebug.Print 1\n```");
                Assert.AreEqual(1, Get<List<ChatEntry>>(window, "transcriptEntries").Count);
                Assert.AreEqual(1, ((ICollection)Get<object>(window, "visibleEntries")).Count);
                Get<List<ChatAttachment>>(window, "draftAttachments").Add(new ChatAttachment { Label = "Selection", Text = "VBA code" });
                Call(window, "RefreshContextPreview");
                Assert.AreEqual(2, Get<FlowLayoutPanel>(window, "contextPreview").Controls.Count);
            }
        }
    }
}

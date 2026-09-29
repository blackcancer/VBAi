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

    /// <summary>Vérifie les limites des jetons de référence et des pièces jointes du compositeur.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Accepte les jetons isolés et refuse les pièces jointes dépassant la limite de contenu.</summary>
        [TestMethod]
        [STATestMethod]
        public void ComposerTokenBoundaryAndDraftContextLimitAreEnforced()
        {
            using (var window = Surfaces())
            {
                var token = "#Project.Module";
                Assert.AreEqual(true, Call(window, "ContainsToken", "Use " + token + " now", token));
                Assert.AreEqual(false, Call(window, "ContainsToken", "x" + token + "suffix", token));
                Assert.AreEqual(true, Call(window, "IsReferenceChar", '_'));
                Assert.AreEqual(false, Call(window, "IsReferenceChar", '-'));
                var attachments = Get<List<ChatAttachment>>(window, "draftAttachments");
                attachments.Add(new ChatAttachment { Label = "selection", Text = "code" });
                var prepared = (ChatAttachment[])Call(window, "PrepareAttachments", "question");
                Assert.AreEqual(1, prepared.Length);
                Assert.AreEqual("selection", prepared[0].Label);
                attachments[0].Text = new string ('x', 48001);
                var error = Assert.ThrowsException<TargetInvocationException>(() => Call(window, "PrepareAttachments", "question"));
                Assert.IsInstanceOfType(error.InnerException, typeof(InvalidOperationException));
            }
        }
    }
}

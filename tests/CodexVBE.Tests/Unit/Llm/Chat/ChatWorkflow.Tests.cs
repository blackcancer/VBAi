namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class HistoryAndCodeTests
    {
        [TestMethod]
        public void CommandsExpandOnlyCompleteTokensAndRejectUnknownCommands()
        {
            var expanded = ChatCommand.Expand("/plan\nInspecter les modules");
            StringAssert.Contains(expanded, "Prépare un plan concret");
            StringAssert.Contains(expanded, "Inspecter les modules");
            Assert.AreEqual("Texte libre", ChatCommand.Expand("Texte libre"));
            Assert.ThrowsException<InvalidOperationException>(() => ChatCommand.Expand("/planifier"));
        }

        [TestMethod]
        public void HistorySearchFindsUnicodeTextAndChangedCode()
        {
            var session = new ChatSessionState
            {
                Title = "Révision",
                Scope = "Classeur.xlsx"
            };
            session.Entries.Add(new ChatEntry { Speaker = "Utilisateur", Text = "Procédure été" });
            session.Entries.Add(new ChatEntry { Speaker = "Assistant", Change = new CodeChange { Module = "ModuleCalcul", Before = "AncienneValeur", After = "NouvelleValeur" } });
            Assert.IsTrue(ChatHistory.Matches(session, "RÉVISION"));
            Assert.IsTrue(ChatHistory.Matches(session, "ÉTÉ"));
            Assert.IsTrue(ChatHistory.Matches(session, "modulecalcul"));
            Assert.IsTrue(ChatHistory.Matches(session, "nouvellevaleur"));
            Assert.IsFalse(ChatHistory.Matches(session, "introuvable"));
        }

        [TestMethod]
        public void ExportUsesLongerFenceWhenAttachmentContainsMarkdownFence()
        {
            var fence = new string ((char)96, 3);
            var longer = new string ((char)96, 4);
            var session = new ChatSessionState
            {
                Title = "Export",
                Scope = "Book.xlsm"
            };
            session.Entries.Add(new ChatEntry { Speaker = "Utilisateur", Text = "Voici le code", Attachments = new[] { new ChatAttachment { Label = "Module1", Text = "avant\n" + fence + "\naprès" } } });
            var output = ChatHistory.Export(session);
            StringAssert.Contains(output, "# Export");
            StringAssert.Contains(output, UiText.Get("Document: ") + "Book.xlsm");
            StringAssert.Contains(output, longer + "text\navant\n" + fence + "\naprès\n" + longer);
        }
    }
}

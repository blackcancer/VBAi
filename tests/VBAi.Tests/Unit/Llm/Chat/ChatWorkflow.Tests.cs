namespace VBAi.Tests.Unit
{
    using System;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

    /// <summary>Vérifie les commandes, recherches et exports de l’historique de conversation.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class HistoryAndCodeTests
    {
        /// <summary>Développe uniquement les jetons de commande complets et refuse les commandes inconnues.</summary>
        [TestMethod]
        public void CommandsExpandOnlyCompleteTokensAndRejectUnknownCommands()
        {
            var expanded = ChatCommand.Expand("/plan\nInspecter les modules");
            StringAssert.Contains(expanded, "Prépare un plan concret");
            StringAssert.Contains(expanded, "Inspecter les modules");
            Assert.AreEqual("Texte libre", ChatCommand.Expand("Texte libre"));
            Assert.ThrowsException<InvalidOperationException>(() => ChatCommand.Expand("/planifier"));
        }

        /// <summary>Recherche dans le titre, les messages et le code modifié sans ignorer les caractères Unicode.</summary>
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

        /// <summary>Choisit une clôture Markdown plus longue lorsque le code joint contient déjà une clôture.</summary>
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
        [TestMethod]
        public void CommandsExposeLocalizedTokensModesInstructionsAndAttachmentContracts()
        {
            using(var culture=new Infrastructure.LocalizationScope())
            {
                foreach(string language in new[] {"fr-FR","en-US"})
                {
                    Infrastructure.LocalizationScope.Set(language);
                    foreach(var command in ChatCommand.All)
                    {
                        Assert.AreEqual(language.StartsWith("fr")?command.Token:command.EnglishToken,command.DisplayToken);
                        Assert.AreEqual(command.Kind,command.DisplayKind);Assert.IsFalse(string.IsNullOrEmpty(command.Instruction));
                        StringAssert.StartsWith(ChatCommand.Expand(command.Token+" inspect"),command.Instruction);
                        StringAssert.StartsWith(ChatCommand.Expand(command.EnglishToken+" inspect"),command.Instruction);
                    }
                }
                Assert.AreEqual(ChatMode.Discussion,ChatCommand.All[0].Mode);Assert.AreEqual(ChatMode.Plan,ChatCommand.All.Last().Mode);
                var attachment=new ChatAttachment {Label="L",Text="T",Project="P",Module="M",Sha256="sha",StartLine=7};
                Assert.AreEqual("P",attachment.Project);Assert.AreEqual("M",attachment.Module);Assert.AreEqual("sha",attachment.Sha256);Assert.AreEqual(7,attachment.StartLine);
            }
        }

        [TestMethod]
        public void HistoryMatchesEachChangeFieldAndExportKeepsSafeFencesAndOptionalMemory()
        {
            using(var culture=new Infrastructure.LocalizationScope())
            {
                var session=new ChatSessionState {Title="T",Scope="S"};
                session.Entries.Add(new ChatEntry {Speaker="Utilisateur",Text=null});
                session.Entries.Add(new ChatEntry {Speaker="Assistant",Change=new CodeChange {Module="M",Before="before",After="after"},AttachedMemory="memory",Attachments=new[] {new ChatAttachment {Label="empty",Text=null}}});
                foreach(string query in new[] {"T","M","before","after"})Assert.IsTrue(ChatHistory.Matches(session,query));
                Assert.IsFalse(ChatHistory.Matches(session,"unmatched"));
                string output=ChatHistory.Export(session);StringAssert.Contains(output,"before");StringAssert.Contains(output,"Attached memory");StringAssert.Contains(output,"memory");
                session.Entries[1].AttachedMemory=" ";Assert.IsFalse(ChatHistory.Export(session).Contains("Attached memory"));
            }
        }

    }
}

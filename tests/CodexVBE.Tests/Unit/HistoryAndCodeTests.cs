using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CodexVBE;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class HistoryAndCodeTests
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
            var session = new ChatSessionState { Title = "Révision", Scope = "Classeur.xlsx" };
            session.Entries.Add(new ChatEntry { Speaker = "Utilisateur", Text = "Procédure été" });
            session.Entries.Add(new ChatEntry { Speaker = "Assistant", Change = new CodeChange {
                Module = "ModuleCalcul", Before = "AncienneValeur", After = "NouvelleValeur"
            } });
            Assert.IsTrue(ChatHistory.Matches(session, "RÉVISION"));
            Assert.IsTrue(ChatHistory.Matches(session, "ÉTÉ"));
            Assert.IsTrue(ChatHistory.Matches(session, "modulecalcul"));
            Assert.IsTrue(ChatHistory.Matches(session, "nouvellevaleur"));
            Assert.IsFalse(ChatHistory.Matches(session, "introuvable"));
        }

        [TestMethod]
        public void ExportUsesLongerFenceWhenAttachmentContainsMarkdownFence()
        {
            var fence = new string((char)96, 3);
            var longer = new string((char)96, 4);
            var session = new ChatSessionState { Title = "Export", Scope = "Book.xlsm" };
            session.Entries.Add(new ChatEntry {
                Speaker = "Utilisateur", Text = "Voici le code",
                Attachments = new[] { new ChatAttachment { Label = "Module1", Text = "avant\n" + fence + "\naprès" } }
            });
            var output = ChatHistory.Export(session);
            StringAssert.Contains(output, "# Export");
            StringAssert.Contains(output, UiText.Get("Document: ") + "Book.xlsm");
            StringAssert.Contains(output, longer + "text\navant\n" + fence + "\naprès\n" + longer);
        }

        [TestMethod]
        public void PreviewRejectsInvalidRangeBeforeAnyHostEdit()
        {
            var request = new Request { StartLine = 3, Count = 1, Text = "C" };
            Assert.ThrowsException<ArgumentException>(() => CodeChange.Preview("A\nB", request));
            request.StartLine = 2;
            request.Count = -1;
            Assert.ThrowsException<ArgumentException>(() => CodeChange.Preview("A\nB", request));
        }

        [TestMethod]
        public void RollbackRestoresExactChangeAndRefusesAmbiguousContext()
        {
            var before = "start\nold\nend";
            var after = "start\nnew\nend";
            var change = new CodeChange { Module = "M1", Before = before, After = after };
            Assert.AreEqual(before.Replace("\n", "\r\n"), CodeRollback.Apply(change, after));
            var conflict = Assert.ThrowsException<InvalidOperationException>(
                () => CodeRollback.Apply(change, "other\nnew\nother"));
            StringAssert.Contains(conflict.Message, "Conflit");
        }

        [TestMethod]
        public void DiffRowsPreserveOldAndNewLineNumbers()
        {
            var rows = CodeChange.BuildRows("first\nold\nlast", "first\nnew\nlast");
            var removed = rows.Single(x => x.Kind == CodeDiffKind.Removed);
            var added = rows.Single(x => x.Kind == CodeDiffKind.Added);
            Assert.AreEqual(2, removed.OldLine);
            Assert.AreEqual(2, added.NewLine);
            Assert.AreEqual("old", removed.Text);
            Assert.AreEqual("new", added.Text);
        }
    }
}

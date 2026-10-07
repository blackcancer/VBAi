using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie les transformations lexicales et leurs bornes sans modifier un hôte VBA.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTextEditsTests
    {
        /// <summary>Exerce chaque garde de plage, taille et argument de remplacement.</summary>
        [TestMethod]
        public void InvalidRangesSizesActionsAndReplacementArgumentsAreRejected()
        {
            foreach (var range in new[] { new[] { 0, 1 }, new[] { 1, 0 }, new[] { 3, 1 }, new[] { 2, 2 } })
                Assert.ThrowsException<ArgumentException>(() => VbaTextEdits.Transform("a\nb", new Request { StartLine = range[0], Count = range[1], Action = "comment" }));
            foreach (var request in new[] {
                new Request { Text = new string('x', 262145) }, new Request { Query = new string('x', 4097) },
                new Request { Action = "replace", Query = null, Text = "x" }, new Request { Action = "replace", Query = "", Text = "x" },
                new Request { Action = "replace", Query = "x", Text = null }, new Request { Action = "replace", Query = "x\nx", Text = "x" },
                new Request { Action = "replace", Query = "x\rx", Text = "x" }, new Request { Action = "unknown" } })
            {
                request.StartLine = request.Count = 1;
                Assert.ThrowsException<ArgumentException>(() => VbaTextEdits.Transform("x", request));
            }
            Assert.ThrowsException<ArgumentException>(() => VbaTextEdits.Transform(null, new Request { StartLine = 1, Count = 1, Action = "comment" }));
            Assert.AreEqual("'x", VbaTextEdits.Transform("x", new Request { StartLine = 1, Count = 1, Action = "comment", Text = new string('x', 262144), Query = new string('x', 4096) }));
        }

        /// <summary>Préserve les lignes hors portée et traite casse, mots, apostrophes et indentation.</summary>
        [TestMethod]
        public void AllTransformationsPreserveScopeAndHonorExplicitOptions()
        {
            var request = new Request { StartLine = 2, Count = 1, Action = "replace", Query = "x", Text = "$1", MatchCase = true };
            Assert.AreEqual("before\r\nX $1 $1y\r\nafter", VbaTextEdits.Transform("before\nX x xy\nafter", request));
            request.MatchCase = false; request.WholeWord = true; request.Text = "";
            Assert.AreEqual("before\r\n  xy\r\nafter", VbaTextEdits.Transform("before\nX x xy\nafter", request));
            request.Action = "uncomment";
            Assert.AreEqual("before\r\n  text\r\nafter", VbaTextEdits.Transform("before\n  'text\nafter", request));
            request.Action = "indent"; request.InsertIndex = null;
            Assert.AreEqual("before\r\n    x\r\nafter", VbaTextEdits.Transform("before\nx\nafter", request));
            request.InsertIndex = 16;
            Assert.AreEqual("before\r\n" + new string(' ', 16) + "x\r\nafter", VbaTextEdits.Transform("before\nx\nafter", request));
            request.Action = "unindent"; request.StartLine = 1; request.Count = 4; request.InsertIndex = 2;
            Assert.AreEqual("tab\r\nspace\r\n space\r\nplain", VbaTextEdits.Transform("\ttab\n space\n   space\nplain", request));
            foreach (int size in new[] { 0, 17 })
            { request.InsertIndex = size; Assert.ThrowsException<ArgumentException>(() => VbaTextEdits.Transform("\ttab\n space\n   space\nplain", request)); }
        }

        /// <summary>Refuse les identifiants invalides et remplace uniquement les jetons non opaques.</summary>
        [TestMethod]
        public void IdentifierValidationAndLexicalBoundariesCoverOpaqueAndIncompleteTokens()
        {
            foreach (string value in new[] { null, "", "1bad", "If", "Rem", new string('a', 256) })
            {
                var request = new Request { StartLine = 1, Count = 1, Action = "replace_identifier", Query = value, NewName = "replacement" };
                Assert.ThrowsException<ArgumentException>(() => VbaTextEdits.Transform("value", request));
                request.Query = "value"; request.NewName = value;
                Assert.ThrowsException<ArgumentException>(() => VbaTextEdits.Transform("value", request));
            }
            Assert.AreEqual("renamed", VbaTextEdits.Transform("VALUE", new Request { StartLine = 1, Count = 1, Action = "replace_identifier", Query = "value", NewName = "renamed" }));
            foreach (var pair in new[] {
                new[] { "", "" }, new[] { "x", "x" }, new[] { "Rem", "Rem" }, new[] { "Rem x\r\nvalue", "Rem x\r\nrenamed" },
                new[] { "Remx:value\n_value", "Remx:renamed\n_value" }, new[] { "value1 + value", "value1 + renamed" },
                new[] { "'value\rvalue", "'value\rrenamed" }, new[] { "\"value\"\"value\":value", "\"value\"\"value\":renamed" },
                new[] { "\"value\"\"", "\"value\"\"" }, new[] { "\"value", "\"value" }, new[] { "\"value\nvalue", "\"value\nrenamed" },
                new[] { "[value]:value", "[value]:renamed" }, new[] { "[value", "[value" }, new[] { "#value#:value", "#value#:renamed" },
                new[] { "#value\rvalue", "#value\rrenamed" }, new[] { "é_1 = value", "é_1 = renamed" } })
                Assert.AreEqual(pair[1], VbaTextEdits.ReplaceIdentifier(pair[0], "value", "renamed"), pair[0]);
        }
    }
}

using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeDebugWindowsParsingTests
    {
        [TestMethod]
        public void FrenchAndEnglishLocalRowsExposeExpressionValueAndType()
        {
            dynamic french = VbeDebugWindows.ParseDebugRow(
                "Expression compteur Valeur 42 Type Integer", new[] { "compteur" });
            Assert.AreEqual("compteur", (string)french.Expression);
            Assert.AreEqual("42", (string)french.Value);
            Assert.AreEqual("Integer", (string)french.Type);
            Assert.IsTrue((bool)french.Parsed);
            Assert.AreEqual(0, (int)french.Depth);

            dynamic english = VbeDebugWindows.ParseDebugRow(
                "Expression obj Value {Object} Type Worksheet", new[] { "obj", "Name" });
            Assert.AreEqual("obj", (string)english.Expression);
            Assert.AreEqual("{Object}", (string)english.Value);
            Assert.AreEqual("Worksheet", (string)english.Type);
            Assert.AreEqual(1, (int)english.Depth);
            CollectionAssert.AreEqual(new[] { "obj", "Name" }, (string[])english.PathSegments);
        }

        [TestMethod]
        public void FrenchAndEnglishWatchRowsExposeContext()
        {
            dynamic french = VbeDebugWindows.ParseDebugRow(
                "counter Valeur 3 Type Long Contexte VBAProject.Module1.Main", new[] { "counter" });
            Assert.AreEqual("counter", (string)french.Expression);
            Assert.AreEqual("3", (string)french.Value);
            Assert.AreEqual("VBAProject.Module1.Main", (string)french.Context);
            Assert.IsTrue((bool)french.Parsed);

            dynamic english = VbeDebugWindows.ParseDebugRow(
                "counter Value 4 Type Long Context VBAProject.Module1.Main", new[] { "counter" });
            Assert.AreEqual("counter", (string)english.Expression);
            Assert.AreEqual("4", (string)english.Value);
            Assert.AreEqual("VBAProject.Module1.Main", (string)english.Context);
        }

        [TestMethod]
        public void NoVariablesPlaceholderIsOmittedButOtherUnparsedRowsRemainVisible()
        {
            Assert.IsNull(VbeDebugWindows.ParseDebugRow(
                "Expression  Valeur Aucune variable Type ", new string[0]));
            Assert.IsNull(VbeDebugWindows.ParseDebugRow(
                "Expression  Value No variables Type ", new string[0]));

            dynamic unknown = VbeDebugWindows.ParseDebugRow("Provider-specific row", null);
            Assert.IsFalse((bool)unknown.Parsed);
            Assert.AreEqual("Provider-specific row", (string)unknown.Raw);
            Assert.IsNull((string)unknown.Expression);
            Assert.AreEqual(0, ((string[])unknown.PathSegments).Length);
        }

        [TestMethod]
        public void ItemPathSegmentUnderstandsLocalAndWatchRows()
        {
            Assert.AreEqual("child", VbeDebugWindows.ParseItemPathSegment(
                "Expression child Valeur 7 Type Long"));
            Assert.AreEqual("child", VbeDebugWindows.ParseItemPathSegment(
                "Expression child Value 7 Type Long"));
            Assert.AreEqual("counter", VbeDebugWindows.ParseItemPathSegment(
                "counter Value 4 Type Long Context VBAProject.Module1.Main"));
            Assert.IsNull(VbeDebugWindows.ParseItemPathSegment("Unreadable UIA row"));
            Assert.IsNull(VbeDebugWindows.ParseItemPathSegment(null));
        }

        [TestMethod]
        public void WatchContextParserRecognizesLocalizedSuffixOnly()
        {
            Assert.AreEqual("VBAProject.Module1.Main", VbeDebugWindows.ParseWatchContext(
                "counter Valeur 3 Type Long Contexte VBAProject.Module1.Main"));
            Assert.AreEqual("VBAProject.Module1.Main", VbeDebugWindows.ParseWatchContext(
                "counter Value 3 Type Long Context VBAProject.Module1.Main"));
            Assert.IsNull(VbeDebugWindows.ParseWatchContext("counter Value 3 Type Long"));
            Assert.IsNull(VbeDebugWindows.ParseWatchContext(null));
        }

        [TestMethod]
        public void OnlyKnownVbaDiagnosticPrefixesMayAuthorizeButtonResponse()
        {
            foreach (var message in new[] { "Erreur d'exécution '9'", "Run-time error '9'",
                "Erreur de compilation: Syntaxe", "Compile error: Syntax error" })
                Assert.IsTrue(VbeDebugWindows.IsRecognizedDiagnostic(message), message);
            foreach (var message in new[] { null, "", "Other application error",
                "Note: Compile error", "Windows Security" })
                Assert.IsFalse(VbeDebugWindows.IsRecognizedDiagnostic(message), message);
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass, TestCategory("Unit")]
    public sealed class VbaDeclarationIndexTests
    {
        [TestMethod]
        public void IndexSeparatesScopesTypesConstantsAndPhysicalContinuationPositions()
        {
            string code = "Public a, b As Long\r\nPrivate Const Label As String = \"a,b: Dim fake\", Count = 2\r\n" +
                "Public Sub Test(ByVal argument As Long, _\r\n Optional text As String = \"x,y\", ParamArray values() As Variant)\r\n" +
                " Dim local As New Scripting.Dictionary, other$\r\n Static saved As Double\r\nEnd Sub\r\n" +
                "Private Type Record\r\n Field As String * 12\r\nEnd Type\r\nPublic Enum Color\r\n Red = 1\r\nEnd Enum";
            var symbols = VbaDeclarationIndex.Read(code);
            Assert.AreEqual(14, symbols.Length);
            Assert.AreEqual("Variant", symbols.Single(x => x.Name == "a").TypeName);
            Assert.AreEqual("Long", symbols.Single(x => x.Name == "b").TypeName);
            Assert.AreEqual("Constant", symbols.Single(x => x.Name == "Label").Kind);
            Assert.AreEqual("Test", symbols.Single(x => x.Name == "argument").Scope);
            Assert.AreEqual(4, symbols.Single(x => x.Name == "text").Line);
            Assert.AreEqual(11, symbols.Single(x => x.Name == "text").Column);
            Assert.AreEqual("Scripting.Dictionary", symbols.Single(x => x.Name == "local").TypeName);
            Assert.AreEqual("String", symbols.Single(x => x.Name == "other").TypeName);
            Assert.AreEqual("Record", symbols.Single(x => x.Name == "Field").Scope);
            Assert.AreEqual("EnumMember", symbols.Single(x => x.Name == "Red").Kind);
            Assert.IsFalse(symbols.Any(x => x.Name == "fake"));
        }
        [TestMethod]
        public void IndexExcludesCommentsStringsAndMarksExternalAndConditionalDeclarations()
        {
            var symbols = VbaDeclarationIndex.Read("'Dim fake\nRem Dim hidden\nDebug.Print \"Dim falseName: Const nope = 1\"\n" +
                "#If VBA7 Then\nPrivate value As LongPtr\n#Else\nPrivate value As Long\n#End If\n" +
                "Declare PtrSafe Function External Lib \"x\" () As Long\nDim réel As Long: Dim nextValue As String\n");
            Assert.AreEqual(5, symbols.Length);
            Assert.AreEqual(2, symbols.Count(x => x.Conditional));
            Assert.AreEqual(10, symbols.Single(x => x.Name == "réel").Line);
            Assert.AreEqual("String", symbols.Single(x => x.Name == "nextValue").TypeName);
            Assert.AreEqual(0, VbaDeclarationIndex.Read(null).Length);
        }
        [TestMethod]
        public void ArrayDimensionsAndDefaultExpressionsDoNotCreateAdditionalParameters()
        {
            var symbols = VbaDeclarationIndex.Read("Public Sub Run(Optional x As Long = Factory(1, 2), ByRef items() As String)\n" +
                "Dim grid(1 To 2, 1 To 4) As Double, count%\nEnd Sub\nDim moduleOnly As Object");
            Assert.AreEqual(5, symbols.Length);
            Assert.AreEqual("Integer", symbols.Single(x => x.Name == "count").TypeName);
            Assert.AreEqual("Module", symbols.Single(x => x.Name == "moduleOnly").Scope);
        }

        [TestMethod]
        public void IncompleteDeclarationsAndAllContainerAndModifierKindsHaveExplicitResults()
        {
            string source = "#\n#End\n#End Else\n#End If\nEnd\nEnd If\n" +
                "Sub\nSub 1\nSub Bare\nSub Open(\nSub Nested(x As Long, y As Factory(1))\nEnd Sub\n" +
                "Function F(ByRef a As Long)\nEnd Function\nProperty Get P(Optional b As String)\nEnd Property\n" +
                "Event\nEvent 1\nEvent Changed()\nDeclare\nDeclare PtrSafe Other\nDeclare Sub External Lib \"x\" ()\nDeclare Function Fn Lib \"x\" ()\nDeclare Sub 1\n" +
                "Type\nType 1\nType T\n1 Invalid\nField As Long\nEnd Type\nEnum\nEnum 1\nEnum E\nFirst = 1\nEnd Enum\n" +
                "Friend Global Static counter As Long\nDim WithEvents otherListener As Object\nWithEvents listener As Object\nPrivate WithEvents notifier As Object\n" +
                "Dim\nDim Optional, invalid As, other As New, broken As Long = 1, , _bad\n" +
                "Dim s$, i%, l&, f!, d#, c@, ll^\n";
            var symbols = VbaDeclarationIndex.Read(source);
            Assert.IsTrue(symbols.Any(x => x.Name == "Changed" && x.Kind == "Event"));
            Assert.AreEqual(2, symbols.Count(x => x.Kind == "ExternalProcedure"));
            Assert.IsTrue(symbols.Any(x => x.Name == "counter" && x.TypeName == "Long"));
            Assert.IsTrue(symbols.Any(x => x.Name == "Field" && x.Kind == "Field"));
            Assert.IsTrue(symbols.Any(x => x.Name == "First" && x.Kind == "EnumMember"));
            Assert.AreEqual("", symbols.Single(x => x.Name == "invalid").TypeName);
            Assert.AreEqual("", symbols.Single(x => x.Name == "other").TypeName);
            foreach (var pair in new[] { new[] { "s", "String" }, new[] { "i", "Integer" }, new[] { "l", "Long" }, new[] { "f", "Single" }, new[] { "d", "Double" }, new[] { "c", "Currency" }, new[] { "ll", "LongLong" } })
                Assert.AreEqual(pair[1], symbols.Single(x => x.Name == pair[0]).TypeName);
            var suffix = typeof(VbaDeclarationIndex).GetMethod("SuffixType", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
            Assert.AreEqual("Variant", suffix.Invoke(null, new object[] { '?' }));
        }

        [TestMethod]
        public void LexerHandlesEveryLiteralCommentContinuationAndEndOfInputShape()
        {
            foreach (string source in new[] { "", "\n\r\n", ":", "word:", "word:=", "word", "x_2$", "_", "a _\n b", "_\n", "a\tb", "'tail", "Rem tail", "Rem\nDim x", "\"unterminated", "\"a\"\"b\"", "\"last\"", "[name]", "[unterminated\n", "x #date#", "x #unfinished", "x #unfinished\n", "x#", "a!" })
            {
                var tokens = VbaDeclarationIndex.Statements(source).SelectMany(x => x).ToArray();
                Assert.IsTrue(tokens.All(t => t.Line > 0 && t.Column > 0 && t.Text.Length > 0), source);
            }
            var escaped = VbaDeclarationIndex.Statements("Dim x As String = \"a\"\"b\": Dim y As Date = #1/2/2026#").ToArray();
            Assert.AreEqual(2, escaped.Length);
            Assert.AreEqual(2, escaped.SelectMany(x => x).Count(x => x.Text == "<literal>"));
            var continuation = VbaDeclarationIndex.Statements("Dim a, _\r\n b As Long").Single();
            Assert.AreEqual(2, continuation.Single(x => x.Text == "b").Line);
        }
    }
}

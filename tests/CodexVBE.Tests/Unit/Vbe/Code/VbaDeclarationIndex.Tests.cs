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
    }
}

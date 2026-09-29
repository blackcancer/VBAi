using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaLocalScalarCandidatesTests
    {
        [TestMethod]
        public void ExplicitScalarParametersAndLocalDeclarationsKeepTheirSourceExpressions()
        {
            const string source = "Private moduleValue As Long\r\n" +
                "Public Sub Inspect(ByVal count As Long, ByRef label$)\r\n" +
                "    Dim amount As Currency, ready As Boolean\r\n" +
                "    Static retained%\r\n" +
                "    Const threshold As Double = 2.5\r\n" +
                "End Sub\r\n";
            var candidates = VbaLocalScalarCandidates.Read(source, "Inspect", 2, 6);
            CollectionAssert.AreEqual(new[] { "count", "label", "amount", "ready", "retained", "threshold" },
                candidates.Select(c => c.Name).ToArray());
            CollectionAssert.AreEqual(new[] { "count", "label$", "amount", "ready", "retained%", "threshold" },
                candidates.Select(c => c.Expression).ToArray());
            Assert.IsTrue(candidates.All(c => c.Eligible));
            Assert.AreEqual(0, candidates.Count(c => c.Name == "moduleValue"));
            Assert.AreEqual(4, candidates.Single(c => c.Name == "retained").Line);
        }

        [TestMethod]
        public void UnsafeAndAmbiguousDeclarationsRemainVisibleButCannotBeEvaluated()
        {
            const string source = "Public Sub Inspect(ByVal values() As Long, ByVal freeValue As Variant)\r\n" +
                "    Dim objectValue As Object, recordValue As SomeRecord\r\n" +
                "    Dim instantiated As New Widget\r\n" +
                "    Dim numbers(1 To 3) As Long\r\n" +
                "    Dim same As Long, SAME As Long\r\n" +
                "#If Flag Then\r\n" +
                "    Dim conditionalValue As Long\r\n" +
                "#End If\r\n" +
                "End Sub";
            var candidates = VbaLocalScalarCandidates.Read(source, "Inspect", 1, 9);
            Assert.IsFalse(candidates.Any(c => c.Eligible));
            Assert.AreEqual("ArrayDeclaration", candidates.Single(c => c.Name == "values").Reason);
            Assert.AreEqual("NonScalarOrVariantType", candidates.Single(c => c.Name == "freeValue").Reason);
            Assert.AreEqual("NonScalarOrVariantType", candidates.Single(c => c.Name == "objectValue").Reason);
            Assert.AreEqual("NonScalarOrVariantType", candidates.Single(c => c.Name == "recordValue").Reason);
            Assert.AreEqual("AutoInstantiatedDeclaration", candidates.Single(c => c.Name == "instantiated").Reason);
            Assert.AreEqual("ArrayDeclaration", candidates.Single(c => c.Name == "numbers").Reason);
            Assert.IsTrue(candidates.Where(c => c.Name.Equals("same", StringComparison.OrdinalIgnoreCase))
                .All(c => c.Reason == "AmbiguousDeclaration"));
            Assert.AreEqual("ConditionalDeclaration", candidates.Single(c => c.Name == "conditionalValue").Reason);
        }

        [TestMethod]
        public void ProcedureRangeAndScopeExcludeNeighbouringPropertiesAndModuleMembers()
        {
            const string source = "Dim outside As Long\r\n" +
                "Property Get Value() As Long\r\n" +
                "    Dim propertyValue As Long\r\n" +
                "End Property\r\n" +
                "Property Let Value(ByVal incoming As Long)\r\n" +
                "    Dim setterValue As Long\r\n" +
                "End Property\r\n" +
                "Sub Inspect()\r\n" +
                "    Dim inside As Long\r\n" +
                "End Sub";
            var candidates = VbaLocalScalarCandidates.Read(source, "Inspect", 8, 10);
            CollectionAssert.AreEqual(new[] { "inside" }, candidates.Select(c => c.Name).ToArray());
            Assert.IsTrue(candidates[0].Eligible);
        }
    }
}

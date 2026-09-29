using System.Linq;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit.Editor
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorLanguageIndexTests
    {
        [TestMethod]
        public void ProjectIndexRetainsTypesVisibilityScopeAndPhysicalLocations()
        {
            var source = new EditorSource { Module = "Client", ComponentType = 1, Text = "Option Explicit\nPrivate hidden As Long\nPublic Function Calculate(ByVal value As Long, _\n Optional label As String = \"a,b\") As Double\n Dim item As Widget\n Calculate = value\nEnd Function\nPublic Sub Other()\nEnd Sub" };
            var symbols = EditorLanguageIndex.Build(new[] { source, new EditorSource { Module = "Widget", ComponentType = 2, Text = "Public Name As String\nPrivate secret As Long" } });
            var method = symbols.Single(s => s.Name == "Calculate");
            Assert.AreEqual("Double", method.TypeName); Assert.AreEqual(3, method.Line); Assert.AreEqual(7, method.EndLine);
            CollectionAssert.AreEqual(new[] { "value As Long", "label As String" }, method.Parameters);
            Assert.IsTrue(symbols.Single(s => s.Name == "hidden").Private);
            Assert.AreEqual("Calculate", symbols.Single(s => s.Name == "item").Scope);
            Assert.AreEqual("Widget", symbols.Single(s => s.Name == "item").TypeName);
            Assert.AreEqual("Class", symbols.Single(s => s.Name == "Widget").Kind);
        }
        [TestMethod]
        public void ConditionalIncompleteAndPropertyStatementsKeepOwnedPhysicalRanges()
        {
            string code = "#If VBA7 Then\nPrivate Property Get Value()\nDim item As Long\nEnd Property\n#Else\nPublic Sub Another()\nEnd Sub\n#End If\n#End If\nPublic\nPrivate Friend Static\nSub\nProperty Get\nEnd\nEnd Enum\n#Const Flag = True\nPublic Function NoType()\nDim thing As String\nEnd Function";
            var symbols = EditorLanguageIndex.Build(new[] { new EditorSource { Module = "OwnedClass", ComponentType = 2, Text = code } });
            Assert.IsTrue(symbols.All(symbol => !symbol.External));
            var property = symbols.Single(symbol => symbol.Name == "Value"); Assert.AreEqual("Property", property.Kind); Assert.IsTrue(property.Private); Assert.IsTrue(property.Conditional); Assert.AreEqual(4, property.EndLine); Assert.AreEqual("Variant", property.TypeName);
            var another = symbols.Single(symbol => symbol.Name == "Another"); Assert.IsTrue(another.Conditional); Assert.AreEqual(7, another.EndLine);
            var plain = symbols.Single(symbol => symbol.Name == "NoType"); Assert.IsFalse(plain.Conditional); Assert.AreEqual(19, plain.EndLine);
            Assert.IsTrue(symbols.Any(symbol => symbol.Name == "thing" && symbol.Scope == "NoType"));
            Assert.AreEqual(1, EditorLanguageIndex.Build(new[] { new EditorSource { Module = "Empty", ComponentType = 1, Text = "" } }).Length);
        }
    }
}

using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit.Editor
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
    }
}

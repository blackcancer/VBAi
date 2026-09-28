using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit.Editor
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorAttributeGuardTests
    {
        [TestMethod]
        public void ModuleVariableAttributesAreNotDiscardedByReplacement()
        {
            const string source = "Public Value As Long";
            Assert.ThrowsException<InvalidOperationException>(() => EditorVbeModule.EnsureAttributeDeclarationsUntouched("Attribute Value.VB_VarHelpID = 1", source, Tuple.Create(1, 1, "Public Other As Long")));
        }
        [TestMethod]
        public void BodyAndAdjacentInsertionsPreserveAttributedDeclarations()
        {
            string source = "Public Sub Special( _\n ByVal value As Long)\n Debug.Print value\nEnd Sub";
            string export = "Attribute Special.VB_Description = \"Keep\"";
            EditorVbeModule.EnsureAttributeDeclarationsUntouched(export, source, EditorDocument.Difference(source, source.Replace("Print value", "Print 42")));
            EditorVbeModule.EnsureAttributeDeclarationsUntouched(export, source, EditorDocument.Difference(source, "' before\n" + source));
            Assert.ThrowsException<InvalidOperationException>(() => EditorVbeModule.EnsureAttributeDeclarationsUntouched(export, source, EditorDocument.Difference(source, source.Replace("Special", "Other"))));
            Assert.ThrowsException<InvalidOperationException>(() => EditorVbeModule.EnsureAttributeDeclarationsUntouched(export, source, Tuple.Create(2, 0, "' split")));
        }
    }
}

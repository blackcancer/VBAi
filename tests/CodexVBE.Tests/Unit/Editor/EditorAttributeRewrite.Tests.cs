using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit.Editor
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorAttributeRewriteTests
    {
        [TestMethod]
        public void RenamePreservesOpaqueMetadataAndOtherProcedureAttributes()
        {
            string before = "Public Sub Old()\nEnd Sub\nPublic Sub Other()\nEnd Sub\n";
            string exported = "Attribute VB_Name = \"Module1\"\nPublic Sub Old()\nAttribute Old.VB_Description = \"Old text\"\nEnd Sub\nPublic Sub Other()\nAttribute Other.VB_Description = \"Other\"\nEnd Sub\n";
            string result = EditorAttributeRewrite.Prepare(exported, before, Tuple.Create(1, 1, "Public Sub NewName()"));
            StringAssert.Contains(result, "Attribute NewName.VB_Description = \"Old text\"");
            StringAssert.Contains(result, "Attribute Other.VB_Description = \"Other\"");
        }
        [TestMethod]
        public void RefusesStaleSourceMultilineAndChangedProcedureKind()
        {
            string source = "Public Sub Old()\nAttribute Old.VB_Description = \"Keep\"\nEnd Sub";
            string before = "Public Sub Old()\nEnd Sub";
            Assert.IsNull(EditorAttributeRewrite.Prepare(source, before + "\n'changed", Tuple.Create(1, 1, "Public Sub NewName()")));
            Assert.IsNull(EditorAttributeRewrite.Prepare(source, before, Tuple.Create(1, 1, "Public Function NewName()")));
            Assert.IsNull(EditorAttributeRewrite.Prepare(source, before, Tuple.Create(1, 1, "Public Sub NewName(value As Long)")));
            Assert.IsNull(EditorAttributeRewrite.Prepare(source, before, Tuple.Create(1, 1, "Public Sub NewName( _\n)")));
        }
    }
}

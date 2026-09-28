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
        public void RefusesStaleSourceAndChangedProcedureKind()
        {
            string source = "Public Sub Old()\nAttribute Old.VB_Description = \"Keep\"\nEnd Sub";
            string before = "Public Sub Old()\nEnd Sub";
            Assert.IsNull(EditorAttributeRewrite.Prepare(source, before + "\n'changed", Tuple.Create(1, 1, "Public Sub NewName()")));
            Assert.IsNull(EditorAttributeRewrite.Prepare(source, before, Tuple.Create(1, 1, "Public Function NewName()")));
            Assert.IsNotNull(EditorAttributeRewrite.Prepare(source, before, Tuple.Create(1, 1, "Public Sub NewName(value As Long)")));
            Assert.IsNotNull(EditorAttributeRewrite.Prepare(source, before, Tuple.Create(1, 1, "Public Sub NewName( _\n)")));
        }
        [TestMethod]
        public void MultilineSignatureAndClassHeaderPreserveMemberMetadata()
        {
            string before = "Public Function ReadValue( _\n    ByVal count As Long) As Long\nEnd Function\n";
            string exported = "VERSION 1.0 CLASS\nBEGIN\n MultiUse = -1\nEND\nAttribute VB_Name = \"Example\"\nAttribute VB_PredeclaredId = True\nPublic Function ReadValue( _\n    ByVal count As Long) As Long\nAttribute ReadValue.VB_UserMemId = 0\nEnd Function\n";
            string after = before.Replace("ReadValue", "Value").Replace("count As Long", "count As Double");
            string result = EditorAttributeRewrite.Prepare(exported, before, EditorDocument.Difference(before, after));
            Assert.IsNotNull(result); Assert.IsFalse(result.Contains("VERSION 1.0"));
            StringAssert.Contains(result, "Attribute Value.VB_UserMemId = 0");
            StringAssert.Contains(result, "Attribute VB_PredeclaredId = True");
        }
        [TestMethod]
        public void ParameterMetadataAndConditionalDeclarationsAreNotSilentlyRebound()
        {
            string before = "Public Sub Old(value As Long)\nEnd Sub";
            string exported = "Public Sub Old(value As Long)\nAttribute Old.value.VB_VarDescription = \"Keep\"\nEnd Sub";
            string after = before.Replace("value As Long", "other As Long");
            Assert.IsNull(EditorAttributeRewrite.Prepare(exported, before, EditorDocument.Difference(before, after)));
            string conditional = "#If Win64 Then\n" + before + "\n#End If";
            Assert.IsNull(EditorAttributeRewrite.Prepare("#If Win64 Then\n" + exported + "\n#End If", conditional, Tuple.Create(2, 1, "Public Sub NewName(value As Long)")));
        }
    }
}

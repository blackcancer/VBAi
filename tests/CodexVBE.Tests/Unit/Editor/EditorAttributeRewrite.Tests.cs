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
        [TestMethod]
        public void InvalidDeclarationPatchAndMetadataOwnershipNeverRebindAttributes()
        {
            string before = "Public Sub Old()\nEnd Sub";
            string exported = "Public Sub Old()\nAttribute Old.VB_Description = \"Keep\"\nEnd Sub";
            foreach (var patch in new[] { Tuple.Create(0, 1, ""), Tuple.Create(1, -1, ""), Tuple.Create(3, 1, ""), Tuple.Create(1, 1, ""), Tuple.Create(1, 1, "Sub Old"), Tuple.Create(1, 1, "Sub Old(") })
                Assert.IsNull(EditorAttributeRewrite.Prepare(exported, before, patch));
            foreach (string invalid in new[] { "Sub Old", "Sub Old(" })
                Assert.IsNull(EditorAttributeRewrite.Prepare(invalid, invalid, Tuple.Create(1, 1, "Sub Old()")));
            Assert.IsNull(EditorAttributeRewrite.Prepare("Attribute Old.VB_Description = \"Wrong location\"\n" + before, before, Tuple.Create(1, 1, "Public Sub Renamed()")));
            Assert.IsNull(EditorAttributeRewrite.Prepare("Public Sub Old()\nEnd Sub\nAttribute Old.VB_Description = \"Too late\"", before, Tuple.Create(1, 1, "Public Sub Renamed()")));
            Assert.IsNull(EditorAttributeRewrite.Prepare(exported.Replace("Attribute Old.", "Attribute Other."), before, Tuple.Create(1, 1, "Public Sub Renamed()")));
            foreach (string tokens in new[] { "Public", "Sub", "Sub 123()", "Property Get", "Private Property Get 123()" })
                Assert.IsFalse(EditorAttributeRewrite.HasMultilineAttributes(tokens));
            Assert.IsTrue(EditorAttributeRewrite.HasMultilineAttributes("Sub Incomplete("));
            Assert.IsFalse(EditorAttributeRewrite.HasMultilineAttributes("Sub Complete()\nEnd Sub"));
            Assert.AreEqual("plain\r\ncode", EditorAttributeRewrite.FullExport("plain", "plain\ncode"));
        }

        [TestMethod]
        public void MultilineInsertOverlapAndDuplicateSignaturesAreValidatedBeforeRewrite()
        {
            string before = "Public Sub Old( _\n ByVal x As Long)\nEnd Sub";
            string exported = "Public Sub Old( _\n ByVal x As Long)\nAttribute Old.VB_Description = \"Keep\"\nEnd Sub";
            Assert.IsTrue(EditorAttributeRewrite.HasMultilineAttributes(exported));
            Assert.IsFalse(EditorAttributeRewrite.HasMultilineAttributes(before));
            Assert.IsNotNull(EditorAttributeRewrite.Prepare(exported, before, Tuple.Create(2, 0, " ByVal y As Long, _")));
            Assert.IsNull(EditorAttributeRewrite.Prepare(exported, before, Tuple.Create(1, 0, "' before")));
            Assert.IsNull(EditorAttributeRewrite.Prepare(exported, before, Tuple.Create(3, 0, "' after")));
            Assert.IsNull(EditorAttributeRewrite.Prepare(exported, before, Tuple.Create(3, 1, "' no signature overlap")));
            string duplicate = "Sub Old()\nEnd Sub\nSub Old()\nEnd Sub";
            string metadata = "Sub Old()\nAttribute Old.VB_Description = \"First\"\nEnd Sub\nSub Old()\nEnd Sub";
            Assert.IsNull(EditorAttributeRewrite.Prepare(metadata, duplicate, Tuple.Create(1, 1, "Sub Renamed()")));
            Assert.IsNull(EditorAttributeRewrite.Prepare(metadata, duplicate, Tuple.Create(1, 1, "Sub Old(value As Long)")));
            string two = "Sub Old()\nEnd Sub\nSub Other()\nEnd Sub";
            Assert.IsNull(EditorAttributeRewrite.Prepare("Sub Old()\nAttribute Old.VB_Description = \"Keep\"\nEnd Sub\nSub Other()\nEnd Sub", two, Tuple.Create(1, 1, "Sub Other()")));
            string parameter = "Sub Old(value As Long)\nAttribute Old.value.VB_VarDescription = \"Keep\"\nEnd Sub";
            string renamed = EditorAttributeRewrite.Prepare(parameter, "Sub Old(value As Long)\nEnd Sub", Tuple.Create(1, 1, "Sub Renamed(value As Long)"));
            StringAssert.Contains(renamed, "Attribute Renamed.value.VB_VarDescription");
            string prefixed = "' prefix\n" + exported;
            Assert.IsNull(EditorAttributeRewrite.Prepare(prefixed, "' prefix\n" + before, Tuple.Create(1, 1, "' unchanged declaration")));
            Assert.IsNull(EditorAttributeRewrite.Prepare("Sub Old()\nAttribute Old.VB_Description = \"Keep\"\nEnd Sub\nSub Other()\nEnd Sub", two, Tuple.Create(3, 1, "Sub Renamed()")));
            Assert.AreEqual("VERSION 1.0 CLASS\r\nAttribute VB_Name = \"Kept\"", EditorAttributeRewrite.FullExport("VERSION 1.0 CLASS\nAttribute VB_Name = \"Original\"", "Attribute VB_Name = \"Kept\""));
            Assert.AreEqual("Attribute VB_Name = \"Module1\"\nAttribute Old.VB_Description = \"Keep\"", EditorAttributeRewrite.Metadata("Attribute VB_Name = \"Module1\"\n" + exported));
            Assert.AreEqual("Attribute Old.VB_Description = \"Keep\"", EditorAttributeRewrite.MemberMetadata("Attribute VB_Name = \"Module1\"\n" + exported));
            string property = "Public Property Get Value()\nEnd Property";
            Assert.IsNotNull(EditorAttributeRewrite.Prepare("Public Property Get Value()\nAttribute Value.VB_UserMemId = 0\nEnd Property", property, Tuple.Create(1, 1, "Public Property Get Renamed()")));
        }
    }
}

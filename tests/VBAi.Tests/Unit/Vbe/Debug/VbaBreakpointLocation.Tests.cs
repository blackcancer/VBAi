using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass]
    public sealed class VbaBreakpointLocationTests
    {
        [DataTestMethod]
        [DataRow("")]
        [DataRow("' Debug.Print 1")]
        [DataRow("Rem Debug.Print 1")]
        [DataRow("Dim value As Long")]
        [DataRow("Static value As Long")]
        [DataRow("Const value = 1")]
        [DataRow("target:")]
        [DataRow("100:")]
        [DataRow("target: Rem comment")]
        [DataRow("#If VBA7 Then")]
        public void NonExecutableLinesAreRejected(string statement)
        {
            Assert.IsFalse(VbaBreakpointLocation.CanRequest("Sub Example()\n" + statement + "\nEnd Sub", 2));
        }

        [DataTestMethod]
        [DataRow("Debug.Print 1")]
        [DataRow("target: Debug.Print 1")]
        [DataRow("100 Debug.Print 1")]
        [DataRow("Dim value As Long: value = 1")]
        [DataRow("Call Example")]
        [DataRow("Example")]
        [DataRow("Debug.Print \"don't split: this\"")]
        public void ExecutableStatementsAreAccepted(string statement)
        {
            Assert.IsTrue(VbaBreakpointLocation.CanRequest("Private Sub Example()\n" + statement + "\nEnd Sub", 2));
        }

        [TestMethod]
        public void DeclarationsProcedureHeadersAndContinuationsAreRejected()
        {
            string source = "Option Explicit\nPrivate Sub Example( _\nByVal value As Long)\nDebug.Print _\nvalue\nEnd Sub\nDebug.Print 2";
            foreach (int line in new[] { -1, 0, 1, 2, 3, 5, 7, 8 }) Assert.IsFalse(VbaBreakpointLocation.CanRequest(source, line), "Line " + line);
            Assert.IsTrue(VbaBreakpointLocation.CanRequest(source, 4));
            Assert.IsTrue(VbaBreakpointLocation.CanRequest(source, 6));
        }
        [DataTestMethod]
        [DataRow("True", true, false)]
        [DataRow("1", true, false)]
        [DataRow("False", false, true)]
        [DataRow("0", false, true)]
        [DataRow("VBA7", false, false)]
        [DataRow("UnknownConstant", false, false)]
        public void ConditionalBranchesOnlyAllowKnownActiveStatements(string expression, bool ifBranch, bool elseBranch)
        {
            string source = "Sub Example()\n#If " + expression + " Then\nDebug.Print 1\n#Else\nDebug.Print 2\n#End If\nEnd Sub";
            Assert.AreEqual(ifBranch, VbaBreakpointLocation.CanRequest(source, 3), expression + " If branch");
            Assert.AreEqual(elseBranch, VbaBreakpointLocation.CanRequest(source, 5), expression + " Else branch");
            Assert.IsFalse(VbaBreakpointLocation.CanRequest(source, 2));
            Assert.IsFalse(VbaBreakpointLocation.CanRequest(source, 4));
        }

        [TestMethod]
        public void NestedKnownBranchesAndUnknownCompilerConstantsFailClosed()
        {
            string nested = "Sub Example()\n#If True Then\n#If False Then\nDebug.Print 1\n#Else\nDebug.Print 2\n#End If\n#End If\nDebug.Print 3\nEnd Sub";
            Assert.IsFalse(VbaBreakpointLocation.CanRequest(nested, 4));
            Assert.IsTrue(VbaBreakpointLocation.CanRequest(nested, 6));
            Assert.IsTrue(VbaBreakpointLocation.CanRequest(nested, 9));

            string unknown = "Sub Example()\n#If VBA7 Then\nDebug.Print 1\n#ElseIf True Then\nDebug.Print 2\n#Else\nDebug.Print 3\n#End If\nDebug.Print 4\nEnd Sub";
            foreach (int line in new[] { 3, 5, 7 }) Assert.IsFalse(VbaBreakpointLocation.CanRequest(unknown, line), "Line " + line);
            Assert.IsTrue(VbaBreakpointLocation.CanRequest(unknown, 9));
        }

        [TestMethod]
        public void InactiveSignaturesCannotAuthorizeCodeOutsideProcedures()
        {
            string source = "#If False Then\nSub Hidden()\nDebug.Print 1\nEnd Sub\n#End If\nDebug.Print 2\nSub Active()\nDebug.Print 3\nEnd Sub\nDebug.Print 4";
            foreach (int line in new[] { 2, 3, 4, 6, 7, 10 }) Assert.IsFalse(VbaBreakpointLocation.CanRequest(source, line), "Line " + line);
            Assert.IsTrue(VbaBreakpointLocation.CanRequest(source, 8));
            Assert.IsTrue(VbaBreakpointLocation.CanRequest(source, 9));
        }
    }
}

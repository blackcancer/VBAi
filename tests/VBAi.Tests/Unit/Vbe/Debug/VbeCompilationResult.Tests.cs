namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

    [TestClass]
    public sealed class VbeCompilationResultTests
    {
        [DataTestMethod]
        [DataRow("Absent")]
        [DataRow("Disabled")]
        public void UnavailableCommandNeverClaimsSuccessfulCompilation(string capability)
        {
            dynamic result = VbeCompilationResult.Create("P", new { Executed = false, Available = false, Capability = capability, Reason = "not available" }, null);
            Assert.IsFalse((bool)result.Available); Assert.IsFalse((bool)result.Compiled);
            Assert.AreEqual("NativeCompile" + capability, (string)result.Verification);
            Assert.AreEqual("not available", (string)result.Diagnostic); Assert.IsNull((object)result.NextRead);
        }

        [TestMethod]
        public void CommandAdmissionAndNativeDiagnosticAreBothRequired()
        {
            dynamic absentProof = VbeCompilationResult.Create("P", new { Available = true }, null);
            Assert.IsFalse((bool)absentProof.Compiled);
            dynamic completed = VbeCompilationResult.Create("P", new { Executed = true }, null);
            Assert.IsTrue((bool)completed.Compiled);
            dynamic diagnostic = VbeCompilationResult.Create("P", new { Executed = true }, "Syntax error");
            Assert.IsFalse((bool)diagnostic.Compiled); Assert.AreEqual("NativeDiagnosticCaptured", (string)diagnostic.Verification);
        }
    }
}

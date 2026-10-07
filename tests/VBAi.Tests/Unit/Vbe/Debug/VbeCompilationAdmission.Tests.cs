namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

    [TestClass]
    public sealed class VbeCompilationAdmissionTests
    {
        [TestMethod]
        public void CancellationRefusesLateAdmissionAndCannotCancelEnteredWork()
        {
            var pending = new VbeCompilationAdmission();
            Assert.IsTrue(pending.CancelPending()); Assert.IsFalse(pending.TryBegin());
            var entered = new VbeCompilationAdmission();
            Assert.IsTrue(entered.TryBegin()); Assert.IsFalse(entered.CancelPending()); Assert.IsFalse(entered.TryBegin());
        }
    }
}

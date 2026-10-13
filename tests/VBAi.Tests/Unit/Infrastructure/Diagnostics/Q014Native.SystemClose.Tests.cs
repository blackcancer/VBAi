using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class Q014NativeSystemCloseTests
    {
        [TestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void NativeBoundaryReceivesExactlyOneSystemMenuCloseWithTheAdmittedHandle(bool delivered)
        {
            int entries = 0;
            bool result = Q014Native.DeliverVbeSystemClose(12345, (handle, message, command) =>
            {
                entries++;
                Assert.AreEqual(12345L, handle);
                Assert.AreEqual(0x0112u, message, "The user close route starts at WM_SYSCOMMAND.");
                Assert.AreEqual(0xF060L, command, "The command must be SC_CLOSE, not host Quit.");
                return delivered;
            });
            Assert.AreEqual(delivered, result);
            Assert.AreEqual(1, entries);
        }

        [TestMethod]
        [DataRow(0L)]
        [DataRow(-1L)]
        public void InvalidTargetIsRejectedBeforeNativeEntry(long handle)
        {
            int entries = 0;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
                Q014Native.DeliverVbeSystemClose(handle, (target, message, command) => { entries++; return true; }));
            Assert.AreEqual(0, entries);
        }

        [TestMethod]
        public void MissingOrThrowingDeliveryNeverFallsBackToAnotherCloseRoute()
        {
            Assert.ThrowsException<ArgumentNullException>(() => Q014Native.DeliverVbeSystemClose(1, null));
            int entries = 0;
            var cause = new InvalidOperationException("Native delivery failed");
            var observed = Assert.ThrowsException<InvalidOperationException>(() =>
                Q014Native.DeliverVbeSystemClose(1, (handle, message, command) => { entries++; throw cause; }));
            Assert.AreSame(cause, observed);
            Assert.AreEqual(1, entries);
        }
    }
}

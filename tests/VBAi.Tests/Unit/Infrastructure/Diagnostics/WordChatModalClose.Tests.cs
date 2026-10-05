using System;
using System.ComponentModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks the one-shot native close boundary without an Office host or window.</summary>
    [TestClass]
    public sealed class WordChatModalCloseTests
    {
        [TestMethod]
        public void InvalidOrChangedModalCannotDispatchClose()
        {
            int calls = 0;
            Func<IntPtr, bool> post = _ => { calls++; return true; };
            Assert.ThrowsException<ArgumentException>(() => WordChatModalClose.PostOnce(IntPtr.Zero, () => { }, post));
            Assert.ThrowsException<ArgumentNullException>(() => WordChatModalClose.PostOnce(new IntPtr(41), null, post));
            Assert.ThrowsException<ArgumentNullException>(() => WordChatModalClose.PostOnce(new IntPtr(41), () => { }, null));
            Assert.ThrowsException<InvalidOperationException>(() => WordChatModalClose.PostOnce(new IntPtr(41),
                () => { throw new InvalidOperationException("Owner changed."); }, post));
            Assert.AreEqual(0, calls);
        }

        [TestMethod]
        public void GuardPrecedesExactSingleDelivery()
        {
            bool guarded = false;
            int calls = 0;
            WordChatModalClose.PostOnce(new IntPtr(41), () => guarded = true, window => {
                Assert.IsTrue(guarded); Assert.AreEqual(new IntPtr(41), window); calls++; return true;
            });
            Assert.AreEqual(1, calls);
        }

        [TestMethod]
        public void FailedAndUncertainDeliveryAreNeverReplayed()
        {
            int calls = 0;
            Assert.ThrowsException<Win32Exception>(() => WordChatModalClose.PostOnce(new IntPtr(41), () => { },
                _ => { calls++; return false; }));
            Assert.AreEqual(1, calls);
            calls = 0;
            Assert.ThrowsException<TimeoutException>(() => WordChatModalClose.PostOnce(new IntPtr(41), () => { },
                _ => { calls++; throw new TimeoutException("Delivery uncertain."); }));
            Assert.AreEqual(1, calls);
        }
    }
}

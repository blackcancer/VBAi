using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorBufferTests
    {
        [TestMethod]
        public void ApplyChecksConflictsAndPreservesDraft()
        {
            string source = "Option Explicit";
            var buffer = new EditorBuffer(() => source, value => source = value, () => true);
            buffer.Draft = "Option Explicit\n' Draft";
            source += "\r\n' External";
            Assert.ThrowsException<InvalidOperationException>(() => buffer.Apply());
            Assert.AreEqual("Option Explicit\n' Draft", buffer.Draft);
            Assert.AreEqual("Option Explicit\r\n' External", source);
        }

        [TestMethod]
        public void ApplyRefusesExecutionModeWithoutWriting()
        {
            var buffer = new EditorBuffer(() => "original", value => Assert.Fail("Unexpected write"), () => false);
            buffer.Draft = "changed";
            Assert.ThrowsException<InvalidOperationException>(() => buffer.Apply());
        }

        [TestMethod]
        public void SuccessfulApplyVerifiesAndAcceptsLineEndingNormalization()
        {
            string source = "original";
            var buffer = new EditorBuffer(() => source, value => source = value, () => true);
            buffer.Draft = "first\nsecond\n";
            buffer.Apply();
            Assert.AreEqual("first\r\nsecond", source);
            Assert.IsFalse(buffer.Dirty);
        }

        [TestMethod]
        public void PartialFailureKeepsBaselineAndDraftForRecovery()
        {
            string source = "original";
            var buffer = new EditorBuffer(() => source, value => { source = "partial"; throw new InvalidOperationException(); }, () => true);
            buffer.Draft = "changed";
            Assert.ThrowsException<InvalidOperationException>(() => buffer.Apply());
            Assert.AreEqual("original", buffer.Baseline);
            Assert.AreEqual("changed", buffer.Draft);
            Assert.ThrowsException<InvalidOperationException>(() => buffer.Apply());
        }
    }
}

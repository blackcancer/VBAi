using System.Runtime.InteropServices;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeCodePaneLifetimeTests
    {
        public sealed class Pane
        {
            public int Error;
            public object Module = new object();
            public object CodeModule { get { if (Error != 0) throw new COMException("native pane error", Error); return Module; } }
        }
        [TestMethod]
        public void LivePanePreservesExactModuleIdentity()
        {
            var pane = new Pane();
            Assert.IsTrue(VbeDebug.TryLivePaneModule(pane, out object module, out string error));
            Assert.AreSame(pane.Module, module);
            Assert.IsNull(error);
        }
        [TestMethod]
        public void DestroyedSplitPaneIsReportedWithoutAUsableModule()
        {
            Assert.IsFalse(VbeDebug.TryLivePaneModule(new Pane { Error = unchecked((int)0x80020010) }, out object module, out string error));
            Assert.IsNull(module);
            Assert.AreEqual("native pane error", error);
        }
        [TestMethod]
        public void OtherNativeFailuresAreNotMisclassifiedAsDestroyedPanes()
        {
            Assert.ThrowsException<COMException>(() => VbeDebug.TryLivePaneModule(new Pane { Error = unchecked((int)0x80004005) }, out object module, out string error));
        }
    }
}

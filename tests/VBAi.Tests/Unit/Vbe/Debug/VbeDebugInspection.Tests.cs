using System;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeDebugInspectionTests
    {
        [TestMethod]
        public void NestedScopesKeepNavigationSuspendedUntilTheOuterScopeEnds()
        {
            Assert.IsFalse(VbeDebugInspection.IsActive);
            var outer = new VbeDebugInspection();
            try
            {
                Assert.IsTrue(VbeDebugInspection.IsActive);
                var inner = new VbeDebugInspection();
                try { Assert.IsTrue(VbeDebugInspection.IsActive); }
                finally { inner.Dispose(); }
                Assert.IsTrue(VbeDebugInspection.IsActive);
                inner.Dispose();
                Assert.IsTrue(VbeDebugInspection.IsActive);
            }
            finally { outer.Dispose(); }
            Assert.IsFalse(VbeDebugInspection.IsActive);
            outer.Dispose();
            Assert.IsFalse(VbeDebugInspection.IsActive);
        }

        [TestMethod]
        public void WrongThreadCannotDisposeOwnersInspection()
        {
            var lease = new VbeDebugInspection();
            try
            {
                Exception failure = null;
                var worker = new Thread(() => {
                    try { lease.Dispose(); }
                    catch (Exception error) { failure = error; }
                });
                worker.Start(); worker.Join();
                Assert.IsInstanceOfType(failure, typeof(InvalidOperationException));
                Assert.IsTrue(VbeDebugInspection.IsActive);
            }
            finally { lease.Dispose(); }
            Assert.IsFalse(VbeDebugInspection.IsActive);
        }
    }
}

using System;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Verifies stop confirmation and preservation of native lifetime guards.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeNativeRendererTests
    {
        [TestMethod]
        public void MissingStopExportSucceedsOnlyWhenInactive()
        {
            foreach (bool active in new[] { false, true })
                using (var fixture = new NativeRendererStopFixture(active, false))
                {
                    Assert.AreEqual(!active, VbeNativeRenderer.Stop()); Assert.AreEqual(active, VbeNativeRenderer.Active);
                    Assert.AreEqual(0, fixture.Calls); Assert.AreEqual(fixture.PinnedModule, fixture.Read("module"));
                    Assert.IsNull(fixture.Read("stop"));
                    if (active) StringAssert.Contains(fixture.Messages.Single(), "managed guards must be retained");
                    else Assert.AreEqual(0, fixture.Messages.Count);
                }
        }

        [TestMethod]
        public void ConfirmedStopClearsActiveWithoutUnloadingPinnedModuleOrExports()
        {
            using (var fixture = new NativeRendererStopFixture())
            {
                Assert.IsTrue(VbeNativeRenderer.Stop()); Assert.IsFalse(VbeNativeRenderer.Active); Assert.AreEqual(1, fixture.Calls);
                Assert.AreEqual(fixture.PinnedModule, fixture.Read("module")); Assert.AreSame(fixture.StopExport, fixture.Read("stop"));
                StringAssert.Contains(fixture.Messages.Single(), "stopped: status=0");
                Assert.IsTrue(VbeNativeRenderer.Stop()); Assert.IsFalse(VbeNativeRenderer.Active); Assert.AreEqual(2, fixture.Calls);
            }
        }

        [TestMethod]
        public void WrongThreadStopRetainsActiveAndExplainsOwningVbeThread()
        {
            using (var fixture = new NativeRendererStopFixture { Result = 1444 })
            {
                Assert.IsFalse(VbeNativeRenderer.Stop()); Assert.IsTrue(VbeNativeRenderer.Active); Assert.AreEqual(1, fixture.Calls);
                Assert.AreEqual(fixture.PinnedModule, fixture.Read("module")); Assert.AreSame(fixture.StopExport, fixture.Read("stop"));
                StringAssert.Contains(fixture.Messages.Single(), "ERROR_INVALID_THREAD_ID");
                StringAssert.Contains(fixture.Messages.Single(), "owning VBE UI thread");
                Assert.IsFalse(fixture.Messages.Any(message => message.Contains("renderer stopped")));
            }
        }

        [TestMethod]
        public void GenericNativeErrorPreservesPreviousStateWithoutRetry()
        {
            foreach (bool active in new[] { false, true })
                using (var fixture = new NativeRendererStopFixture(active) { Result = 5 })
                {
                    Assert.IsFalse(VbeNativeRenderer.Stop()); Assert.AreEqual(active, VbeNativeRenderer.Active); Assert.AreEqual(1, fixture.Calls);
                    Assert.AreEqual(fixture.PinnedModule, fixture.Read("module")); Assert.AreSame(fixture.StopExport, fixture.Read("stop"));
                    StringAssert.Contains(fixture.Messages.Single(), "status=5"); StringAssert.Contains(fixture.Messages.Single(), "restoration was not confirmed");
                }
        }

        [TestMethod]
        public void ExplicitSuccessfulRetryIsRequiredBeforeClearingActive()
        {
            using (var fixture = new NativeRendererStopFixture { Result = 1444 })
            {
                Assert.IsFalse(VbeNativeRenderer.Stop()); Assert.IsTrue(VbeNativeRenderer.Active); Assert.AreEqual(1, fixture.Calls);
                fixture.Result = 0;
                Assert.IsTrue(VbeNativeRenderer.Stop()); Assert.IsFalse(VbeNativeRenderer.Active); Assert.AreEqual(2, fixture.Calls);
                Assert.AreEqual(fixture.PinnedModule, fixture.Read("module")); Assert.AreSame(fixture.StopExport, fixture.Read("stop"));
            }
        }

        [TestMethod]
        public void InvocationExceptionRetainsStateAndDoesNotCrossNativeCallbackBoundary()
        {
            using (var fixture = new NativeRendererStopFixture { Failure = new InvalidOperationException("injected invocation failure") })
            {
                Assert.IsFalse(VbeNativeRenderer.Stop()); Assert.IsTrue(VbeNativeRenderer.Active); Assert.AreEqual(1, fixture.Calls);
                Assert.AreEqual(fixture.PinnedModule, fixture.Read("module")); Assert.AreSame(fixture.StopExport, fixture.Read("stop"));
                StringAssert.Contains(fixture.Messages.Single(), "stop could not be confirmed: InvalidOperationException");
                StringAssert.Contains(fixture.Messages.Single(), "managed guards must be retained");
            }
        }
    }
}
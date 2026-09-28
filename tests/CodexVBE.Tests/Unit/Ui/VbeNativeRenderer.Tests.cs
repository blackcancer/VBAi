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
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeNativeRendererLifecycleTests
    {
        [TestMethod]
        public void SuccessfulStartUsesValidatedModuleAndPopulatedStatus()
        {
            using (var fixture = new NativeRendererStopFixture(false))
            {
                fixture.ConfigureLifecycle();
                VbeNativeRenderer.Start(new IntPtr(91));
                Assert.IsTrue(VbeNativeRenderer.Active);
                Assert.AreEqual(1, fixture.StartCalls); Assert.AreEqual(new IntPtr(91), fixture.StartWindow);
                Assert.AreEqual(0, fixture.Calls); Assert.AreEqual(1, fixture.QueryCalls);
                Assert.AreEqual((uint)System.Runtime.InteropServices.Marshal.SizeOf(typeof(VbeNativeRenderer.RendererStatus)), fixture.QuerySize);
                StringAssert.Contains(fixture.Messages.Single(), "active=1, windows=2, imports=3, restored=4, patterns=5, icons=6, text=7, fills=8, unsupported=9, failures=10");
                Assert.AreEqual(fixture.PinnedModule, fixture.Read("module"));
            }
        }

        [TestMethod]
        public void FailedStartsRequireExplicitStopAndPreserveStateWhenStopFails()
        {
            foreach (bool invocationFailure in new[] { false, true })
                foreach (uint stopResult in new uint[] { 0, 1444 })
                    using (var fixture = new NativeRendererStopFixture(true) { Result = stopResult })
                    {
                        fixture.ConfigureLifecycle(); fixture.StartResult = 5;
                        if (invocationFailure) fixture.StartFailure = new InvalidOperationException("synthetic start failure");
                        VbeNativeRenderer.Start(new IntPtr(92));
                        Assert.AreEqual(1, fixture.StartCalls); Assert.AreEqual(1, fixture.Calls);
                        Assert.AreEqual(stopResult != 0, VbeNativeRenderer.Active);
                        Assert.AreEqual(fixture.PinnedModule, fixture.Read("module"));
                        StringAssert.Contains(fixture.Messages.Last(), "legacy recovery retained");
                        if (invocationFailure) StringAssert.Contains(fixture.Messages.Last(), "synthetic start failure");
                    }
        }

        [TestMethod]
        public void RegisterCallsOnlyActiveRendererAndStopsOnlyForUnexpectedError()
        {
            using (var fixture = new NativeRendererStopFixture(false))
            {
                fixture.ConfigureLifecycle(); VbeNativeRenderer.Register(new IntPtr(93));
                Assert.AreEqual(0, fixture.RegisterCalls); Assert.AreEqual(0, fixture.Calls);
            }
            foreach (uint result in new uint[] { 0, 50, 5 })
                using (var fixture = new NativeRendererStopFixture() { RegisterResult = result })
                {
                    fixture.ConfigureLifecycle(); VbeNativeRenderer.Register(new IntPtr(94));
                    Assert.AreEqual(1, fixture.RegisterCalls); Assert.AreEqual(new IntPtr(94), fixture.RegisteredWindow);
                    Assert.AreEqual(result == 5 ? 1 : 0, fixture.Calls);
                    Assert.AreEqual(result != 5, VbeNativeRenderer.Active);
                    if (result == 5) StringAssert.Contains(fixture.Messages.First(), "registration failed: 5");
                    else Assert.AreEqual(0, fixture.Messages.Count);
                }
        }

        [TestMethod]
        public void RefreshCallsOnlyActiveRendererAndStopsOnError()
        {
            using (var fixture = new NativeRendererStopFixture(false))
            {
                fixture.ConfigureLifecycle(); VbeNativeRenderer.Refresh();
                Assert.AreEqual(0, fixture.RefreshCalls); Assert.AreEqual(0, fixture.Calls);
            }
            foreach (uint result in new uint[] { 0, 5 })
                using (var fixture = new NativeRendererStopFixture() { RefreshResult = result })
                {
                    fixture.ConfigureLifecycle(); VbeNativeRenderer.Refresh();
                    Assert.AreEqual(1, fixture.RefreshCalls); Assert.AreEqual(result == 5 ? 1 : 0, fixture.Calls);
                    Assert.AreEqual(result == 0, VbeNativeRenderer.Active);
                    if (result == 5) StringAssert.Contains(fixture.Messages.First(), "import refresh failed: 5");
                    else Assert.AreEqual(0, fixture.Messages.Count);
                }
        }

        [TestMethod]
        public void StatusFailureIsReportedWithoutUndoingSuccessfulStart()
        {
            using (var fixture = new NativeRendererStopFixture(false) { QueryResult = 5 })
            {
                fixture.ConfigureLifecycle(); VbeNativeRenderer.Start(new IntPtr(95));
                Assert.IsTrue(VbeNativeRenderer.Active); Assert.AreEqual(0, fixture.Calls);
                StringAssert.Contains(fixture.Messages.Single(), "query status=5");
            }
            using (var fixture = new NativeRendererStopFixture(false))
            {
                fixture.ConfigureLifecycle(false); VbeNativeRenderer.Start(new IntPtr(96));
                StringAssert.Contains(fixture.Messages.Single(), "not loaded");
            }
        }

        [TestMethod]
        public void PayloadHashMatchesKnownSha256Vector()
        {
            var method = typeof(VbeNativeRenderer).GetMethod("Hash", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
            Assert.AreEqual("BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD", method.Invoke(null, new object[] { System.Text.Encoding.ASCII.GetBytes("abc") }));
        }
    }
}

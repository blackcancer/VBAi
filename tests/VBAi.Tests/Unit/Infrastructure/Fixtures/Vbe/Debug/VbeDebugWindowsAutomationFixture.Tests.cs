using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Windows.Automation;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks the documented low-32-bit Windows protocol representation independently of native HWND allocation.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbeDebugWindowsAutomationFixtureTests
    {
        [TestMethod]
        [DataRow(0L, 0), DataRow(2147483647L, 2147483647)]
        [DataRow(2147483648L, -2147483648), DataRow(4294967295L, -1)]
        [DataRow(-25L, -25), DataRow(4294967271L, -25)]
        [DataRow(-256L, -256), DataRow(4294967040L, -256)]
        public void NativeIdentifiersKeepTheirBitsRegardlessOfSignExtension(long value, int expected)
        {
            Assert.AreEqual(8, IntPtr.Size, "The add-in and its HWND qualification run on x64.");
            Assert.AreEqual(expected, VbeDebugWindowsSystemTests.NativeProtocolInt32(new IntPtr(value)));
        }

        [TestMethod]
        public void AZeroExtendedUiaRootIdentifierReproducesTheOldCheckedNarrowingFailure()
        {
            var value = new IntPtr(4294967271L);
            Assert.ThrowsException<OverflowException>(() => value.ToInt32());
            Assert.AreEqual(System.Windows.Automation.Provider.AutomationInteropProvider.RootObjectId,
                VbeDebugWindowsSystemTests.NativeProtocolInt32(value));
        }

        [STATestMethod]
        [DataRow(false), DataRow(true)]
        public void TeardownDisconnectsCachedOwnedProvidersAndObservesTheOriginalThreadExit(bool nativeDialog)
        {
            var root = new VbeDebugWindowsSystemTests.AutomationNode { Name = "Owned teardown root", Kind = ControlType.Window };
            root.Add(new VbeDebugWindowsSystemTests.AutomationNode { Name = "Owned teardown child", Kind = ControlType.Edit });
            var host = new VbeDebugWindowsSystemTests.AutomationHost(root, optionsDialog: nativeDialog, noActivate: true);
            AutomationElement cachedRoot, cachedChild;
            try
            {
                cachedRoot = AutomationElement.FromHandle(host.Handle);
                Assert.AreEqual(root.Name, cachedRoot.Current.Name);
                cachedChild = cachedRoot.FindFirst(TreeScope.Children,
                    new PropertyCondition(AutomationElement.NameProperty, "Owned teardown child"));
                Assert.IsNotNull(cachedChild);
                Assert.AreEqual("Owned teardown child", cachedChild.Current.Name);
            }
            finally { host.Dispose(); }
            Assert.IsTrue(host.OwnerThreadExited);
            Assert.AreEqual(2, host.DisconnectedProviderCount);
            Assert.ThrowsException<ElementNotAvailableException>(() => { _ = cachedRoot.Current.Name; });
            Assert.ThrowsException<ElementNotAvailableException>(() => { _ = cachedChild.Current.Name; });
        }

        [STATestMethod]
        public void DisposalIsIdempotentAndDoesNotRepeatTheNativeProviderRelease()
        {
            var root = new VbeDebugWindowsSystemTests.AutomationNode { Name = "Owned idempotent teardown", Kind = ControlType.Window };
            var host = new VbeDebugWindowsSystemTests.AutomationHost(root, noActivate: true);
            var cached = AutomationElement.FromHandle(host.Handle);
            Assert.AreEqual(root.Name, cached.Current.Name);
            host.Dispose();
            host.Dispose();
            Assert.IsTrue(host.OwnerThreadExited);
            Assert.AreEqual(1, host.DisconnectedProviderCount);
        }

        [STATestMethod]
        [DataRow(false, false), DataRow(false, true)]
        [DataRow(true, false), DataRow(true, true)]
        public void TeardownClearsTheHostHandleEvenWhenNativeIdentityIsUnavailable(bool nativeDialog, bool missingIdentity)
        {
            var root = new VbeDebugWindowsSystemTests.AutomationNode
            {
                Name = "Owned native identity teardown",
                Kind = ControlType.ComboBox,
                NativeHandle = missingIdentity ? (int?)0 : null
            };
            root.Add(new VbeDebugWindowsSystemTests.AutomationNode { Name = "Owned item" });
            var host = new VbeDebugWindowsSystemTests.AutomationHost(root, optionsDialog: nativeDialog, noActivate: true);
            AutomationElement cached;
            try
            {
                cached = AutomationElement.FromHandle(host.Handle);
                Assert.AreEqual(root.Name, cached.Current.Name);
                Assert.AreNotEqual(IntPtr.Zero, root.Window);
                Assert.IsNotNull(root.HostRawElementProvider);
            }
            finally { host.Dispose(); }
            Assert.AreEqual(IntPtr.Zero, root.Window);
            Assert.IsNull(root.HostRawElementProvider);
            Assert.IsTrue(host.OwnerThreadExited);
            Assert.AreEqual(2, host.DisconnectedProviderCount);
            Assert.ThrowsException<ElementNotAvailableException>(() => { _ = cached.Current.Name; });
        }
    }
}
